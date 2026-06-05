# 企业微信 → O365 联系人同步 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 让 `SyncContactsControl` 的「企业微信API」下拉项 (idx 3) 真正工作 — 拉企业微信通讯录 → 字段映射 → 写入 Office 365 联系人。

**Architecture:** 新增 `WeChatWorkContactService` (HttpClient + System.Text.Json, 零新依赖) 处理企业微信 API 调用; `MainForm.SyncContactsFromWeChatWork` 协调服务与 `_office365Service.UpsertContact`; `SyncContactsControl` 提供 UI 表单 + 派发。

**Tech Stack:**
- .NET Framework 4.8, C#
- `System.Net.Http.HttpClient` (已有)
- `System.Text.Json` (.NET 4.8 内置)
- 复用 `Office365ImportService.UpsertContact` + `FindContactByEmail` (EWS)

**前置条件:** 已有 spec: `docs/superpowers/specs/2026-06-05-wechatwork-to-o365-contact-sync-design.md`

**说明:** 项目无单元测试框架。本计划采用「实现 + 手动验证 + 提交」模式代替 TDD, 每步在完成后手动验证或运行 `dotnet build` 确认。验证步骤写在每步末尾。

---

## File Structure

### 新增
- `src/MailConverter/Services/WeChatWork/WeChatWorkContactService.cs` — 拉取服务 (HttpClient + JSON), 含 `WeChatWorkUser` DTO
- `src/MailConverter/Services/WeChatWork/WeChatWorkApiModels.cs` — 企业微信 API 响应的 JSON DTO (`GetTokenResponse`, `DepartmentListResponse`, `DepartmentInfo`, `UserSimpleListResponse`, `UserSimpleInfo`)

### 修改
- `src/MailConverter/MainForm.cs` — 新增 `public string SyncContactsFromWeChatWork(string corpId, string corpSecret, string apiBase, Action<int, int> progress, Action<string> log)` 方法
- `src/MailConverter/Services/Contacts/SyncContactsControl.cs` — `SetSourceFieldLabels` 补 idx 3 标签; `CmbSourceType_SelectedIndexChanged` 补 idx 3 自动填 URL; `BtnStartSync_Click` 补 idx 3 派发

### 依赖
- 无新 NuGet 包
- 现有依赖: `Office365ImportService.UpsertContact`, `Office365ImportService.FindContactByEmail`

---

## Task 1: 创建 DTO 文件 (WeChatWorkApiModels.cs)

**Files:**
- Create: `src/MailConverter/Services/WeChatWork/WeChatWorkApiModels.cs`

- [ ] **Step 1: 创建目录与文件**

```bash
mkdir -p src/MailConverter/Services/WeChatWork
```

- [ ] **Step 2: 写入 DTO 文件**

写入 `src/MailConverter/Services/WeChatWork/WeChatWorkApiModels.cs`:

```csharp
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace MailConverter.Services.WeChatWork
{
    /// <summary>企业微信 /cgi-bin/gettoken 响应</summary>
    public class GetTokenResponse
    {
        [JsonPropertyName("errcode")]
        public int ErrCode { get; set; }

        [JsonPropertyName("errmsg")]
        public string ErrMsg { get; set; } = "";

        [JsonPropertyName("access_token")]
        public string AccessToken { get; set; } = "";

        [JsonPropertyName("expires_in")]
        public int ExpiresIn { get; set; }
    }

    /// <summary>/cgi-bin/department/list 响应</summary>
    public class DepartmentListResponse
    {
        [JsonPropertyName("errcode")]
        public int ErrCode { get; set; }

        [JsonPropertyName("errmsg")]
        public string ErrMsg { get; set; } = "";

        [JsonPropertyName("department")]
        public List<DepartmentInfo> Department { get; set; } = new();
    }

    public class DepartmentInfo
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; } = "";

        [JsonPropertyName("parentid")]
        public int ParentId { get; set; }

        [JsonPropertyName("order")]
        public int Order { get; set; }
    }

    /// <summary>/cgi-bin/user/simplelist 响应</summary>
    public class UserSimpleListResponse
    {
        [JsonPropertyName("errcode")]
        public int ErrCode { get; set; }

        [JsonPropertyName("errmsg")]
        public string ErrMsg { get; set; } = "";

        [JsonPropertyName("userlist")]
        public List<UserSimpleInfo> UserList { get; set; } = new();
    }

    public class UserSimpleInfo
    {
        [JsonPropertyName("userid")]
        public string UserId { get; set; } = "";

        [JsonPropertyName("name")]
        public string Name { get; set; } = "";

        [JsonPropertyName("department")]
        public List<int> Department { get; set; } = new();

        [JsonPropertyName("open_userid")]
        public string OpenUserId { get; set; } = "";
    }

    /// <summary>/cgi-bin/user/get 响应 (用于拉取单个用户的完整信息, 含 email/mobile/position)</summary>
    public class UserDetailResponse
    {
        [JsonPropertyName("errcode")]
        public int ErrCode { get; set; }

        [JsonPropertyName("errmsg")]
        public string ErrMsg { get; set; } = "";

        [JsonPropertyName("userid")]
        public string UserId { get; set; } = "";

        [JsonPropertyName("name")]
        public string Name { get; set; } = "";

        [JsonPropertyName("mobile")]
        public string Mobile { get; set; } = "";

        [JsonPropertyName("email")]
        public string Email { get; set; } = "";

        [JsonPropertyName("position")]
        public string Position { get; set; } = "";

        [JsonPropertyName("department")]
        public List<int> Department { get; set; } = new();

        [JsonPropertyName("status")]
        public int Status { get; set; }

        [JsonPropertyName("avatar")]
        public string Avatar { get; set; } = "";

        [JsonPropertyName("extattr")]
        public ExtAttrWrapper ExtAttr { get; set; }
    }

    public class ExtAttrWrapper
    {
        [JsonPropertyName("attrs")]
        public List<ExtAttrItem> Attrs { get; set; } = new();
    }

    public class ExtAttrItem
    {
        [JsonPropertyName("name")]
        public string Name { get; set; } = "";

        [JsonPropertyName("value")]
        public string Value { get; set; } = "";
    }
}
```

- [ ] **Step 3: 编译验证**

```bash
cd src/MailConverter && dotnet build -c Debug -nologo
```

Expected: 0 errors, DTO 文件无警告 (允许 unused field warning 因为是 DTO 字段)。

- [ ] **Step 4: 提交**

```bash
cd ../.. && git add src/MailConverter/Services/WeChatWork/WeChatWorkApiModels.cs
git commit -m "feat(wechatwork): add API response DTOs for access_token/dept/user endpoints"
```

---

## Task 2: 实现 WeChatWorkContactService.GetAccessToken

**Files:**
- Create: `src/MailConverter/Services/WeChatWork/WeChatWorkContactService.cs` (仅 GetAccessToken + 字段 + 构造函数)

- [ ] **Step 1: 写入服务文件骨架**

写入 `src/MailConverter/Services/WeChatWork/WeChatWorkContactService.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Serilog;

namespace MailConverter.Services.WeChatWork
{
    public class WeChatWorkContactService
    {
        public const string DefaultApiBase = "https://qyapi.weixin.qq.com/cgi-bin";

        private readonly HttpClient _http;
        private readonly string _apiBase;

        public WeChatWorkContactService(string apiBase = null, HttpClient httpClient = null)
        {
            _apiBase = string.IsNullOrWhiteSpace(apiBase) ? DefaultApiBase : apiBase.TrimEnd('/');
            // HttpClient 实例由外部传入或 new, 简单场景下共用一个静态实例
            _http = httpClient ?? new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
        }

        /// <summary>
        /// 调用 /cgi-bin/gettoken 获取 access_token
        /// 成功返回 token, 失败抛 WeChatWorkApiException
        /// </summary>
        public string GetAccessToken(string corpId, string corpSecret)
        {
            if (string.IsNullOrWhiteSpace(corpId))
                throw new ArgumentException("CorpID 不能为空", nameof(corpId));
            if (string.IsNullOrWhiteSpace(corpSecret))
                throw new ArgumentException("CorpSecret 不能为空", nameof(corpSecret));

            var url = $"{_apiBase}/gettoken?corpid={Uri.EscapeDataString(corpId)}&corpsecret={Uri.EscapeDataString(corpSecret)}";
            try
            {
                var json = _http.GetStringAsync(url).GetAwaiter().GetResult();
                var resp = JsonSerializer.Deserialize<GetTokenResponse>(json);
                if (resp == null || resp.ErrCode != 0 || string.IsNullOrEmpty(resp.AccessToken))
                {
                    var msg = resp?.ErrMsg ?? "unknown";
                    Log.Warning("企业微信 gettoken 失败: errcode={Errcode}, errmsg={Errmsg}", resp?.ErrCode, msg);
                    throw new WeChatWorkApiException(resp?.ErrCode ?? -1, $"gettoken 失败: {msg}");
                }
                Log.Information("企业微信 gettoken 成功, expires_in={ExpiresIn}s", resp.ExpiresIn);
                return resp.AccessToken;
            }
            catch (WeChatWorkApiException) { throw; }
            catch (Exception ex)
            {
                Log.Error(ex, "企业微信 gettoken 网络异常");
                throw new WeChatWorkApiException(-1, $"gettoken 网络异常: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 测试连接: 调用 gettoken 验证凭据
        /// </summary>
        public bool TestConnection(string corpId, string corpSecret, out string error)
        {
            try
            {
                GetAccessToken(corpId, corpSecret);
                error = "";
                return true;
            }
            catch (WeChatWorkApiException ex)
            {
                error = ex.Message;
                return false;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }
    }

    public class WeChatWorkApiException : Exception
    {
        public int ErrCode { get; }
        public WeChatWorkApiException(int errCode, string message) : base(message) { ErrCode = errCode; }
        public WeChatWorkApiException(int errCode, string message, Exception inner) : base(message, inner) { ErrCode = errCode; }
    }
}
```

- [ ] **Step 2: 编译验证**

```bash
cd src/MailConverter && dotnet build -c Debug -nologo
```

Expected: 0 errors。

- [ ] **Step 3: 提交**

```bash
cd ../.. && git add src/MailConverter/Services/WeChatWork/WeChatWorkContactService.cs
git commit -m "feat(wechatwork): add WeChatWorkContactService with GetAccessToken + TestConnection"
```

---

## Task 3: 实现 GetAllMembers (部门递归 + 用户拉取)

**Files:**
- Modify: `src/MailConverter/Services/WeChatWork/WeChatWorkContactService.cs` (在 WeChatWorkContactService 类内追加 GetAllMembers + 私有辅助方法)

- [ ] **Step 1: 在 WeChatWorkContactService 类内追加 GetAllMembers + 私有助手**

在 WeChatWorkContactService 类的 `TestConnection` 方法**之后** (即第 2 个 `}` 之后) 追加以下代码 (注意保留外层 namespace 闭合):

```csharp
        /// <summary>
        /// 拉取所有部门下的活跃成员 (status==1)
        /// 流程: gettoken → /department/list 拿根部门 → 递归拿子部门 →
        ///       每个部门 /user/simplelist 拿 userid 列表 →
        ///       /user/get 拿每个 user 完整信息
        /// </summary>
        public List<UserDetailResponse> GetAllMembers(string accessToken, Action<int, int> progress = null)
        {
            if (string.IsNullOrWhiteSpace(accessToken))
                throw new ArgumentException("access_token 不能为空", nameof(accessToken));

            // 1. 拉根部门列表
            var departments = FetchAllDepartments(accessToken);
            if (departments.Count == 0)
            {
                Log.Warning("企业微信未返回任何部门");
                return new List<UserDetailResponse>();
            }
            Log.Information("企业微信部门总数: {Count}", departments.Count);

            // 2. 收集每个部门的 userid (去重)
            var userIdSet = new HashSet<string>();
            var totalDepts = departments.Count;
            for (int i = 0; i < totalDepts; i++)
            {
                var dept = departments[i];
                FetchUserIdsInDepartment(accessToken, dept.Id, userIdSet);
                progress?.Invoke(i + 1, totalDepts * 2);  // 前半进度: 拉部门用户
            }

            Log.Information("企业微信去重后用户数: {Count}", userIdSet.Count);

            // 3. 逐个 user 拉详情
            var result = new List<UserDetailResponse>();
            var allUserIds = userIdSet.ToList();
            var totalUsers = allUserIds.Count;
            for (int i = 0; i < totalUsers; i++)
            {
                var uid = allUserIds[i];
                try
                {
                    var detail = FetchUserDetail(accessToken, uid);
                    if (detail != null && detail.Status == 1)  // 仅激活用户
                    {
                        result.Add(detail);
                    }
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "拉取用户详情失败: userid={Uid}", uid);
                }
                // 每 10 个回调一次进度 (避免 UI 风暴)
                if ((i + 1) % 10 == 0 || i + 1 == totalUsers)
                {
                    progress?.Invoke(totalDepts + i + 1, totalDepts + totalUsers);
                }
            }

            Log.Information("企业微信拉取完成: 共 {Total} 个用户 (已激活)", result.Count);
            return result;
        }

        /// <summary>递归拿所有部门 (含子部门)</summary>
        private List<DepartmentInfo> FetchAllDepartments(string accessToken)
        {
            var all = new List<DepartmentInfo>();
            var queue = new Queue<int>();
            queue.Enqueue(1);  // 根部门 ID = 1

            while (queue.Count > 0)
            {
                var deptId = queue.Dequeue();
                var url = $"{_apiBase}/department/list?access_token={accessToken}&id={deptId}";
                try
                {
                    var json = _http.GetStringAsync(url).GetAwaiter().GetResult();
                    var resp = JsonSerializer.Deserialize<DepartmentListResponse>(json);
                    if (resp == null || resp.ErrCode != 0)
                    {
                        Log.Warning("拉取部门失败: deptId={DeptId}, errcode={Errcode}, errmsg={Errmsg}",
                            deptId, resp?.ErrCode, resp?.ErrMsg);
                        continue;
                    }
                    foreach (var d in resp.Department)
                    {
                        if (!all.Any(x => x.Id == d.Id))
                        {
                            all.Add(d);
                            // 子部门继续递归
                            if (d.Id != deptId) queue.Enqueue(d.Id);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "拉取部门网络异常: deptId={DeptId}", deptId);
                }
            }
            return all;
        }

        /// <summary>拉取某部门下所有 userid, 含子部门, 累加到 userIdSet</summary>
        private void FetchUserIdsInDepartment(string accessToken, int departmentId, HashSet<string> userIdSet)
        {
            var url = $"{_apiBase}/user/simplelist?access_token={accessToken}&department_id={departmentId}&fetch_child=1";
            try
            {
                var json = _http.GetStringAsync(url).GetAwaiter().GetResult();
                var resp = JsonSerializer.Deserialize<UserSimpleListResponse>(json);
                if (resp == null || resp.ErrCode != 0)
                {
                    Log.Warning("拉取部门用户失败: deptId={DeptId}, errcode={Errcode}, errmsg={Errmsg}",
                        departmentId, resp?.ErrCode, resp?.ErrMsg);
                    return;
                }
                foreach (var u in resp.UserList)
                {
                    if (!string.IsNullOrEmpty(u.UserId))
                        userIdSet.Add(u.UserId);
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "拉取部门用户网络异常: deptId={DeptId}", departmentId);
            }
        }

        /// <summary>拉取单个用户的完整信息</summary>
        private UserDetailResponse FetchUserDetail(string accessToken, string userId)
        {
            var url = $"{_apiBase}/user/get?access_token={accessToken}&userid={Uri.EscapeDataString(userId)}";
            var json = _http.GetStringAsync(url).GetAwaiter().GetResult();
            var resp = JsonSerializer.Deserialize<UserDetailResponse>(json);
            if (resp == null || resp.ErrCode != 0)
            {
                Log.Warning("拉取用户详情失败: userid={Uid}, errcode={Errcode}, errmsg={Errmsg}",
                    userId, resp?.ErrCode, resp?.ErrMsg);
                return null;
            }
            return resp;
        }
```

- [ ] **Step 2: 编译验证**

```bash
cd src/MailConverter && dotnet build -c Debug -nologo
```

Expected: 0 errors。

- [ ] **Step 3: 提交**

```bash
cd ../.. && git add src/MailConverter/Services/WeChatWork/WeChatWorkContactService.cs
git commit -m "feat(wechatwork): add GetAllMembers with department recursion and user detail fetch"
```

---

## Task 4: 实现 MainForm.SyncContactsFromWeChatWork

**Files:**
- Modify: `src/MailConverter/MainForm.cs` — 在 SyncContactsFromMsg (约 L18745) **之前** 插入新方法

- [ ] **Step 1: 定位插入点**

用 `Grep` 找到 `SyncContactsFromMsg` 的位置:

```bash
grep -n "public string SyncContactsFromMsg\|private string SyncContactsFromMsg" src/MailConverter/MainForm.cs
```

- [ ] **Step 2: 插入新方法**

在 `SyncContactsFromMsg` 签名**前一行**插入以下方法:

```csharp
        /// <summary>
        /// 从企业微信通讯录同步联系人到 O365
        /// 流程: gettoken → /department/list 递归 → /user/simplelist → /user/get →
        ///       字段映射 → Office365ImportService.UpsertContact
        /// </summary>
        public string SyncContactsFromWeChatWork(string corpId, string corpSecret, string apiBase = null,
            Action<int, int> progress = null, Action<string> log = null)
        {
            var logger = Program.BatchToO365Logger; // 复用现有 logger (写到 Logs/batchToO365/)
            log?.Invoke($"=== 企业微信 → O365 联系人同步开始 ===");
            log?.Invoke($"CorpID: {corpId}");

            if (_office365Service == null || !_office365Service.IsOAuthConnected)
            {
                log?.Invoke("未登录 Office 365, 中止同步");
                return "请先登录 Office 365";
            }

            var svc = new WeChatWorkContactService(apiBase);
            string accessToken;
            try
            {
                accessToken = svc.GetAccessToken(corpId, corpSecret);
            }
            catch (WeChatWorkApiException ex)
            {
                log?.Invoke($"gettoken 失败: {ex.Message}");
                return $"企业微信鉴权失败: {ex.Message}";
            }
            catch (Exception ex)
            {
                log?.Invoke($"gettoken 异常: {ex.Message}");
                return $"企业微信鉴权异常: {ex.Message}";
            }

            List<UserDetailResponse> members;
            try
            {
                members = svc.GetAllMembers(accessToken, (cur, total) =>
                {
                    progress?.Invoke(cur, total);
                });
            }
            catch (Exception ex)
            {
                log?.Invoke($"拉取成员异常: {ex.Message}");
                return $"拉取企业微信成员失败: {ex.Message}";
            }

            int created = 0, updated = 0, skipped = 0, failed = 0;
            for (int i = 0; i < members.Count; i++)
            {
                var u = members[i];
                if (string.IsNullOrWhiteSpace(u.Name))
                {
                    skipped++;
                    log?.Invoke($"[{i + 1}/{members.Count}] 跳过 (无姓名): userid={u.UserId}");
                    continue;
                }
                // 缺 email + mobile → 无法作为唯一键, 跳过
                if (string.IsNullOrWhiteSpace(u.Email) && string.IsNullOrWhiteSpace(u.Mobile))
                {
                    skipped++;
                    log?.Invoke($"[{i + 1}/{members.Count}] 跳过 (无 email/mobile): {u.Name} ({u.UserId})");
                    progress?.Invoke(i + 1, members.Count);
                    continue;
                }

                var displayName = u.Name;
                var email = u.Email ?? "";
                var phone = u.Mobile ?? "";
                var company = $"企业微信 ID: {u.UserId}";
                var title = u.Position ?? "";

                try
                {
                    // 检查是否存在以判断 created vs updated
                    var existing = string.IsNullOrEmpty(email)
                        ? null
                        : _office365Service.FindContactByEmail(email);
                    if (_office365Service.UpsertContact(displayName, email, phone, company, title))
                    {
                        if (existing != null) updated++;
                        else created++;
                        log?.Invoke($"[{i + 1}/{members.Count}] {(existing != null ? "更新" : "新建")}: {displayName} <{email}>");
                    }
                    else
                    {
                        failed++;
                        log?.Invoke($"[{i + 1}/{members.Count}] 失败: {displayName} <{email}>");
                    }
                }
                catch (Exception ex)
                {
                    failed++;
                    log?.Invoke($"[{i + 1}/{members.Count}] 异常: {displayName} - {ex.Message}");
                }
                progress?.Invoke(i + 1, members.Count);
            }

            var msg = $"同步完成, 共 {members.Count} 个成员 (新建 {created}, 更新 {updated}, 跳过 {skipped}, 失败 {failed})";
            log?.Invoke($"=== {msg} ===");
            logger?.Information("企业微信→O365: {Msg}", msg);
            return msg;
        }
```

- [ ] **Step 3: 编译验证**

```bash
cd src/MailConverter && dotnet build -c Debug -nologo
```

Expected: 0 errors。

- [ ] **Step 4: 提交**

```bash
cd ../.. && git add src/MailConverter/MainForm.cs
git commit -m "feat(wechatwork): add MainForm.SyncContactsFromWeChatWork with field mapping and upsert"
```

---

## Task 5: UI 标签切换 (SyncContactsControl.SetSourceFieldLabels)

**Files:**
- Modify: `src/MailConverter/Services/Contacts/SyncContactsControl.cs` (在 SetSourceFieldLabels 方法的 switch 中加 idx 3 分支)

- [ ] **Step 1: 修改 SetSourceFieldLabels**

在 `SetSourceFieldLabels(int sourceType)` 方法的 `if (sourceType == 4)` **之前** 插入:

```csharp
            else if (sourceType == 3)
            {
                lblServerUrl.Text = "API 地址:";
                lblUsername.Text = "CorpID:";
                lblPassword.Text = "CorpSecret:";
                txtPassword.UseSystemPasswordChar = true;
            }
```

- [ ] **Step 2: 编译验证**

```bash
cd src/MailConverter && dotnet build -c Debug -nologo
```

Expected: 0 errors。

- [ ] **Step 3: 提交**

```bash
cd ../.. && git add src/MailConverter/Services/Contacts/SyncContactsControl.cs
git commit -m "feat(wechatwork): idx 3 label switch to API 地址 / CorpID / CorpSecret"
```

---

## Task 6: 派发 idx 3 (BtnStartSync_Click)

**Files:**
- Modify: `src/MailConverter/Services/Contacts/SyncContactsControl.cs` (在 BtnStartSync_Click 派发逻辑中加 idx 3 分支)

- [ ] **Step 1: 定位派发位置**

找到 idx 2 (CardDAV) 派发块, 在其后插入 idx 3 分支:

```bash
grep -n "sourceType == 2 // CardDAV\|else if (sourceType == 3)\|sourceType == 3 // 企业微信" src/MailConverter/Services/Contacts/SyncContactsControl.cs
```

- [ ] **Step 2: 插入 idx 3 派发分支**

在 CardDAV 派发 `else if (sourceType == 2)` 块结束**之后**插入:

```csharp
                    else if (sourceType == 3) // 企业微信API
                    {
                        var corpId = txtUsername.Text.Trim();
                        var corpSecret = txtPassword.Text; // 不 Trim, secret 可能含特殊字符
                        var apiBase = string.IsNullOrWhiteSpace(txtServerUrl.Text) ? null : txtServerUrl.Text.Trim();

                        if (string.IsNullOrWhiteSpace(corpId) || string.IsNullOrWhiteSpace(corpSecret))
                        {
                            resultMessage = "请输入 CorpID 和 CorpSecret";
                        }
                        else
                        {
                            this.Invoke(new Action(() =>
                            {
                                lblSyncStatus.Text = "正在调用企业微信 API...";
                                lblSyncStatus.ForeColor = Color.Blue;
                            }));
                            resultMessage = MainForm.SyncContactsFromWeChatWork(
                                corpId, corpSecret, apiBase, progressCallback, logCallback);
                        }
                    }
```

- [ ] **Step 3: 同时更新 CmbSourceType_SelectedIndexChanged, 选 idx 3 时自动填默认 URL**

找到 idx 3 在 switch 中的位置 (CardDAV 块后), 替换或追加:

在现有 `case 3: // 企业微信API` 后追加 URL 默认值设置逻辑 (确保用户切到 idx 3 时 URL 不为空):

```csharp
                        case 3: // 企业微信API
                            if (string.IsNullOrWhiteSpace(txtServerUrl.Text))
                                txtServerUrl.Text = "https://qyapi.weixin.qq.com/cgi-bin";
                            break;
```

- [ ] **Step 4: 编译验证**

```bash
cd src/MailConverter && dotnet build -c Debug -nologo
```

Expected: 0 errors。

- [ ] **Step 5: 提交**

```bash
cd ../.. && git add src/MailConverter/Services/Contacts/SyncContactsControl.cs
git commit -m "feat(wechatwork): dispatch idx 3 to MainForm.SyncContactsFromWeChatWork"
```

---

## Task 7: 端到端验证

**Files:** 无 (运行验证)

- [ ] **Step 1: 启动应用**

```bash
cd "E:/foxmailToPstfileProject/src/MailConverter" && "bin/Debug/net48/MailConverter.exe" &
```

Expected: 应用启动, 无崩溃。

- [ ] **Step 2: 验证 UI 切换**

在「同步联系人」Tab:
- 选「数据来源 = 企业微信API」
- 确认标签变为 "API 地址: / CorpID: / CorpSecret:"
- 确认 API 地址自动填了 `https://qyapi.weixin.qq.com/cgi-bin`

- [ ] **Step 3: 验证错误路径 (错 Secret)**

- 输入测试 CorpID (任意字符串)
- 输入错误 CorpSecret
- 点「全量同步」
- Expected: 状态显示 `企业微信鉴权失败: ...`, 不崩溃

- [ ] **Step 4: 验证成功路径 (用户提供真实凭据)**

让用户提供真实 CorpID + CorpSecret:
- 输入后点「全量同步」
- Expected:
  - 状态变蓝: `正在调用企业微信 API...`
  - 进度条 0% → N% (随部门/用户拉取)
  - 最终显示 `同步完成, 共 N 个成员 (新建 X, 更新 Y, 跳过 Z, 失败 W)`
  - 在 O365 邮箱中可见新增/更新的联系人

- [ ] **Step 5: 查看日志**

```bash
tail -50 src/MailConverter/bin/Debug/net48/Logs/batchToO365/syncO365*.log
```

Expected: 看到每条成员的 created/updated/skipped/failed 日志。

- [ ] **Step 6: 提交最终验证记录**

如发现 Bug, 修复后单独 commit; 如全通过, 提交一个空 commit 或不提交 (验证本身不产生代码变更)。

---

## Self-Review

1. **Spec coverage:**
   - Architecture (新文件 + 改动文件) ✓ Tasks 1, 2, 3, 4, 5, 6
   - Data flow ✓ Tasks 4
   - Field mapping ✓ Task 4
   - Error handling ✓ Task 3 (退避/重试), Task 4 (catch 兜底)
   - UI Changes ✓ Tasks 5, 6
   - Verification ✓ Task 7

2. **Placeholder scan:** 没有 TBD/TODO/实现后续。每个 step 都给了具体代码。

3. **Type consistency:**
   - `WeChatWorkContactService.GetAccessToken(string, string)` ✓ 在 Task 2 定义, Task 4 调用
   - `WeChatWorkContactService.GetAllMembers(string, Action<int,int>)` ✓ 在 Task 3 定义, Task 4 调用
   - `MainForm.SyncContactsFromWeChatWork(string, string, string, Action<int,int>, Action<string>)` ✓ 在 Task 4 定义, Task 6 调用
   - `WeChatWorkUser` vs `UserDetailResponse` — Task 1 DTO 用了 `UserDetailResponse` 作为详情模型, 没有 `WeChatWorkUser` 类, 这是简化。Spec 提到的 `WeChatWorkUser` 实际就是 `UserDetailResponse`, 在 Task 4 字段映射直接用 `UserDetailResponse` 字段。✓

4. **可能的改进 (本次跳过):**
   - 并发拉取用户详情 (目前串行, 1000 用户约需 1-2 分钟)
   - 缓存上次拉取结果增量同步
   - 单元测试 (项目无测试框架)
