# 企业微信 → Office 365 联系人同步 — 设计 Spec

**日期**: 2026-06-05
**作者**: Claude (brainstorming 流程)
**状态**: 已批准, 等待实施

## Context

`SyncContactsControl` 已有「企业微信API」下拉项 (idx 3), 但只占位, 真实同步逻辑未实现, 用户在 UI 中选完后点「全量同步」只会走到默认兜底, 没有可用的导入路径。

本 spec 完整实现「企业微信 API → Office 365 联系人」同步链路: 拉企业微信成员, 字段映射后用 `Office365ImportService.UpsertContact` 写入 O365 邮箱。零新依赖, 风格与现有 `CardDavService` 对齐。

## Architecture

### 新增文件
- `src\MailConverter\Services\WeChatWork\WeChatWorkContactService.cs` — 企业微信联系人拉取服务

### 修改文件
- `src\MailConverter\MainForm.cs` — 新增 `SyncContactsFromWeChatWork` 公开方法
- `src\MailConverter\Services\Contacts\SyncContactsControl.cs` — `BtnStartSync_Click` 在 idx 3 分支调用新方法; `SetSourceFieldLabels` 补 idx 3 的标签切换

### 关键类

```csharp
public class WeChatWorkContactService
{
    private const string DefaultApiBase = "https://qyapi.weixin.qq.com/cgi-bin";

    public string GetAccessToken(string corpId, string corpSecret, string apiBase = null);
    public bool TestConnection(string corpId, string corpSecret, out string error);
    public List<WeChatWorkUser> GetAllMembers(string accessToken, string apiBase = null, Action<int, int> progress = null);
}

public class WeChatWorkUser
{
    public string UserId { get; set; }
    public string Name { get; set; }
    public string Mobile { get; set; }
    public string Email { get; set; }
    public string Position { get; set; }
    public List<int> Department { get; set; }
    public int Status { get; set; }       // 1=激活, 2=禁用, 4=未激活
    public string Avatar { get; set; }
    public Dictionary<string, string> ExtAttr { get; set; }
}
```

## Data Flow

```
BtnStartSync_Click (idx 3)
  → 校验 corpId/corpSecret 非空, 校验 O365 已登录
  → MainForm.SyncContactsFromWeChatWork(corpId, corpSecret, apiBase, progress, log)
      1. svc.GetAccessToken → 失败抛错返回
      2. svc.GetAllMembers → 递归拉所有部门成员
         - /cgi-bin/department/list 拿根部门
         - 深度优先遍历子部门
         - 每个部门 /cgi-bin/user/simplelist?department_id=X&fetch_child=1
         - 按 userid 去重 (一个员工可能在多个部门)
         - 过滤 status == 1 (激活)
      3. for each user (N 个):
         - displayName = user.Name
         - email = user.Email (空则用 mobile)
         - phone = user.Mobile
         - company = $"企业微信 ID: {user.UserId}"
         - title = user.Position
         - _office365Service.UpsertContact(displayName, email, phone, company, title)
         - progress?.Invoke(idx, total)
      4. 返回 created/updated/skipped/failed 统计
```

## Field Mapping

| 企业微信字段 | O365 Contact 字段 | 备注 |
|---|---|---|
| `name` | `displayName` | 必填, 缺失跳过 |
| `email` | `emailAddresses[0]` | 可空 |
| `mobile` | `businessPhones[0]` | 可空 |
| `position` | `jobTitle` | 可空 |
| `department[0]` | `department` | 多部门取第一个 |
| `userid` | `companyName` | `"企业微信 ID: {userid}"` |

`UpsertContact` 内部以 email 为唯一键查找; 若 email 为空, 则用 mobile 作为键。两者都空 → 跳过。

## Error Handling

| 错误 | 处理 |
|---|---|
| access_token 错误码 40014/41001/42001 | 重试 gettoken 1 次, 仍失败抛错 |
| 部门 API 限流 45009 | 退避 1s 重试, 3 次后跳过该部门 |
| 单个 user upsert 失败 | catch 异常, 计入 failed, 继续下一个 |
| 进度回调 | 每 10 个 user 一次 (避免 Invoke 风暴) |
| 网络异常 | catch, 返回错误消息, 不崩溃 |

## UI Changes

`SetSourceFieldLabels(int sourceType)` 新增 idx 3 分支:
- lblServerUrl = "API 地址:" (默认值: https://qyapi.weixin.qq.com/cgi-bin/)
- lblUsername = "CorpID:"
- lblPassword = "CorpSecret:"

`UpdateControlsVisibility(int sourceType)` idx 3 已有分支, 不动。

`BtnStartSync_Click` idx 3 派发 (替换原有 "其他" else 兜底):
```csharp
else if (sourceType == 3)
{
    var corpId = txtUsername.Text;
    var corpSecret = txtPassword.Text;
    var apiBase = string.IsNullOrWhiteSpace(txtServerUrl.Text) ? null : txtServerUrl.Text;
    if (string.IsNullOrWhiteSpace(corpId) || string.IsNullOrWhiteSpace(corpSecret))
    {
        lblSyncStatus.Text = "请输入 CorpID 和 CorpSecret";
        lblSyncStatus.ForeColor = Color.Red;
        return;
    }
    resultMessage = MainForm.SyncContactsFromWeChatWork(corpId, corpSecret, apiBase, progressCallback, logCallback);
}
```

`CmbSourceType_SelectedIndexChanged` idx 3: auto-fill txtServerUrl 默认值 (如果用户没改过)。

## MainForm.SyncContactsFromWeChatWork 签名

```csharp
public string SyncContactsFromWeChatWork(
    string corpId,
    string corpSecret,
    string apiBase = null,
    Action<int, int> progress = null,
    Action<string> log = null)
```

返回 `"同步完成, 共 {total} 个成员 (新建 {created}, 更新 {updated}, 跳过 {skipped}, 失败 {failed})"`。

注: `UpsertContact` 现有实现不返回 created/updated 计数, 仅返回 bool。**因此需要在 service 层**记录:
- 调用前 `FindContactByEmail` 检查存在与否 → 计数 "update" vs "create"
- 失败 → 计数 "failed"
- 缺 email + mobile → 计数 "skipped"

## Verification

1. `dotnet build -c Debug` → 0 errors
2. 启动 → 同步联系人 → 选「企业微信API」, 验证标签变为 "API 地址/CorpID/CorpSecret"
3. 输入测试 CorpID + 错误 Secret → 点同步 → 应看到清晰的鉴权错误 (不崩溃)
4. 输入正确凭据 → 拉取成员 → 进度条 0% → N%, 最终显示统计
5. 在 O365 联系人中查看是否新增/更新

## Out of Scope

- 部门映射 (不映射到 O365, 仅 user 级)
- 自定义属性 extattr 读取
- 外部联系人 (external_contact API)
- 反向同步 (O365 → 企业微信)
- 定时同步 / 增量同步 (仅一次性全量)
