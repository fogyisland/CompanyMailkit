# Azure AD 应用注册指南 - OAuth2 认证

本文档说明如何在 Azure Active Directory 中注册应用以获取 Client ID，用于 Office 365 OAuth2 认证。

## 前提条件

- Microsoft 365 管理员账户
- 访问 Azure 门户 (https://portal.azure.com)

## 注册步骤

### 1. 访问 Azure 门户

打开 https://portal.azure.com ，使用你的 Microsoft 365 管理员账户登录。

### 2. 注册新应用

1. 搜索 **"Azure Active Directory"** 或 **"Azure AD"**
2. 在左侧菜单点击 **"应用注册"**
3. 点击 **"新建注册"**
4. 填写以下信息：
   - **名称**: `MailConverter`（或任意名称）
   - **支持的账户类型**: "仅此组织目录中的账户"
5. 点击 **"注册"**

### 3. 获取 Client ID

注册成功后，在应用概览页面可以看到：
- **应用程序(客户端) ID** - 这就是 Client ID，复制保存备用

### 4. 配置重定向 URI

1. 在左侧点击 **"身份验证"**
2. 点击 **"添加平台"**
3. 选择 **"移动和桌面应用程序"**
4. 选择 **"http://localhost"**
5. 在自定义 URI 输入: `http://localhost:5555/`
6. 点击 **"配置"**

### 5. 添加 API 权限

**方法一：从 Microsoft API 添加**
1. 点击 **"API 权限"**
2. 点击 **"添加权限"**
3. 选择 **"Microsoft API"**
4. 选择 **"Exchange"**
5. 选择 **"Exchange Online"**
6. 展开 **"其他"** 或 **"EWS"**
7. 勾选 **"EWS.AccessAsUser.All"**
8. 点击 **"添加权限"**

**方法二：从组织 API 添加（如果方法一找不到）**
1. 点击 **"API 权限"**
2. 点击 **"添加权限"**
3. 选择 **"我所在的组织使用的 API"**
4. 在搜索框输入 `Exchange Online` 或 `EWS`
5. 选择找到的 **Exchange Online API**
6. 展开 **"权限"**
7. 勾选 **"EWS.AccessAsUser.All"**
8. 点击 **"添加权限"**

**方法三：手动添加（如果以上都找不到）**
1. 点击 **"API 权限"**
2. 点击 **"添加权限"**
3. 选择 **"Microsoft Graph"**（不是 Exchange）
4. 选择 **"委托的权限"**
5. 在搜索框输入 `EWS`
6. 勾选 **"EWS.AccessAsUser.All"**
7. 点击 **"添加权限"**

**重要：授予管理员同意**
- 如果看到 "需要管理员同意" 按钮，点击 **"为代表 <你的组织> 授予管理员同意"**
- 这需要 Microsoft 365 管理员权限

### 6. 验证配置

确认以下内容：
- API 权限中包含 `EWS.AccessAsUser.All`
- 重定向 URI 包含 `http://localhost:5555/`

## 在应用程序中使用

1. 启动 MailConverter
2. 选择 **"导入到 Office 365"** 面板
3. 在认证方式下拉框选择 **"OAuth 2.0"**
4. 输入你的 **Client ID**（从 Azure 获取）
5. 输入目标邮箱地址
6. 点击 **"使用 Microsoft 登录"** 按钮
7. 在浏览器中完成 Microsoft 登录
8. 登录成功后会自动获取访问令牌

## 常见问题

### Q: 找不到 Exchange API 怎么办？
A: 尝试以下方法：
1. 确保使用的是 **"委托的权限"**（不是应用程序权限）
2. 选择 **"我所在的组织使用的 API"** 然后搜索 "Exchange"
3. 如果实在找不到，使用 **Microsoft Graph** API，搜索 `EWS.AccessAsUser.All`

### Q: 提示 "需要管理员同意"
A: 需要 Microsoft 365 管理员登录 Azure 门户，点击 "为代表 <你的组织> 授予管理员同意" 按钮。

### Q: OAuth 登录失败，提示权限不足
A: 检查 API 权限中是否添加了 `EWS.AccessAsUser.All`，如果没有请按上述步骤添加。

### Q: OAuth 登录失败，提示 redirect_uri 不匹配
A: 检查重定向 URI 是否配置正确，应为 `http://localhost:5555/`

### Q: 如何查看已注册的应用？
A: 在 Azure 门户 → Azure Active Directory → 应用注册，可以查看所有已注册的应用。

### Q: 没有管理员权限怎么办？
A: 可以让管理员帮忙：
1. 管理员登录 Azure 门户
2. 找到你注册的应用
3. 在 "API 权限" 中点击 "为代表...授予管理员同意"

## PST 批量同步联系人到 O365 所需权限

使用 Client Secret（应用程序权限）同步 PST 联系人的功能，需要额外的 Graph API 权限：

### 所需权限

| 权限名称 | 类型 | 说明 |
|---------|------|------|
| `Contacts.ReadWrite` | 应用程序权限 | 读写用户个人联系人 |
| `OrgContact.ReadWrite.All` | 应用程序权限 | 读写组织全体联系人（可选，更全面）|
| `Calendars.ReadWrite` | 应用程序权限 | 读写用户日历 |
| `Calendars.ReadWrite.All` | 应用程序权限 | 读写组织全部日历 |

### 添加步骤

1. 在 Azure 门户打开你的应用注册
2. 进入 **"API 权限"**
3. 点击 **"添加权限"**
4. 选择 **"Microsoft Graph"**
5. 选择 **"应用程序权限"**（不是"委托的权限"）
6. 搜索并添加：
   - `Contacts.ReadWrite`（联系人同步用）
   - `Calendars.ReadWrite`（日历同步用）
   - 或更高级的 `OrgContact.ReadWrite.All` / `Calendars.ReadWrite.All`
7. **重要：** 点击 **"为代表 [你的组织] 授予管理员同意"** 按钮
8. 需要租户管理员点击同意才能生效

### 注意事项

- **应用程序权限** 不需要用户登录，适用于后台批量处理
- **管理员同意** 是必须的，否则会报 `Access is denied` 错误
- 联系人同步和日历同步需要分别添加对应权限
- 如果使用 `*ReadWrite.All` 权限，还需要额外管理员同意
