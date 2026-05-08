# 小铭邮件百宝箱 功能文档

## 一、项目概述

- **项目名称**: 小铭邮件百宝箱 (xiaomingMailtoolkitCompany)
- **版本**: v1.1.7
- **技术栈**: .NET Framework 4.8, WinForms, MimeKit, MailKit, EWS Managed API, Microsoft Graph SDK
- **用途**: 多功能邮件转换工具，支持 EML/OST/IMAP 转 PST、Exchange 服务器集成、Office 365 导入

---

## 二、主要功能模块

### 2.1 邮件格式转换

| 功能 | 说明 |
|------|------|
| EML → PST | 将 EML 邮件文件批量转换为 PST 格式 |
| OST → PST | 将 Outlook OST 数据文件转换为 PST |
| IMAP → PST | 从 IMAP 服务器同步邮件并导出为 PST |
| 批量导出 PST | 支持过滤 GUID 命名和系统隐藏文件夹 |

### 2.2 Exchange 服务器集成

| 功能 | 说明 |
|------|------|
| EWS 投递 | 通过 Exchange Web Services 投递邮件到目标邮箱 |
| IMAP/EWS 同步 | 本地 EX 百宝箱模式，支持 Impersonation 访问其他邮箱 |
| 邮件夹遍历 | 使用 MsgFolderRoot + Deep 遍历获取邮件文件夹 |
| 过滤器 | 跳过 SearchFolder 和 IPF.Note 以外的文件夹 |

### 2.3 Office 365 / Microsoft Graph

| 功能 | 说明 |
|------|------|
| Graph API 导入 | 使用 Microsoft Graph SDK 5.57.0 导入邮件到 Office 365 |
| Azure 存储 | 支持 Azure Blob Storage 操作 |
| OAuth 认证 | 支持 OAuth 2.0 认证流程 |
| **邮件搜索导出** | 使用 Graph API `$search` 搜索邮件，导出到 PST 文件 |
| **邮件搜索删除** | 搜索邮件并硬删除（permanentDelete），不经过回收站 |

### 2.4 日历和联系人

| 功能 | 说明 |
|------|------|
| 日历导出 CSV/ICS | 从 EWS 获取日历信息并导出 |
| 联系人导出 | 支持 vCard/CardDAV 格式 |
| C# Outlook Interop | 日历/联系人通过 C# Outlook Interop 写入 PST |

### 2.5 Exchange Online 百宝箱 - 邮件搜索导出

**位置**: Exchange Online 百宝箱 → 邮件搜索导出

**功能说明**:

| 功能 | 说明 |
|------|------|
| 关键字搜索 | 使用 `$search="subject:关键字"` 搜索邮件主题 |
| 附件名搜索 | 使用 `$search="attachment:文件名"` 搜索附件名 |
| 日期范围 | 支持按收件日期范围筛选 |
| 导出 PST | 将搜索结果导出为 PST 文件 |
| 导出并删除 | 导出 PST 后硬删除源邮件（permanentDelete） |

**技术实现**:

1. **搜索**: Graph API `$search` 参数，支持分页
   ```
   GET https://graph.microsoft.com/v1.0/users/{email}/messages?$search="subject:关键字"
   ```

2. **导出流程**:
   - Graph API `$value` 获取邮件原始 EML 内容
   - 保存到临时目录
   - 调用 `create_pst.py` 合并到 PST 文件

3. **删除**: Graph API `permanentDelete` 硬删除
   ```
   POST https://graph.microsoft.com/v1.0/users/{email}/messages/{id}/permanentDelete
   ```

**权限要求**: `Mail.ReadWrite`

---

## 三、技术特性

### 3.1 邮件处理

- **MimeKit 4.7.0**: EML 解析，MIME 使用字节数组避免编码损坏
- **MailKit 4.7.0**: IMAP 客户端，支持 InternalDate 保留原始时间
- **EWS Managed API**: Exchange 服务器操作

### 3.2 日志系统

- **Serilog**: 结构化日志记录
- **日志目录结构**:
  ```
  bin/Debug/net48/logs/
  ├── transferPST/           # PST转换日志
  │   ├── EML2PST/
  │   ├── OST2PST/
  │   ├── IMAP2PST/
  │   └── IMAPMULTI2PST/
  └── mailpick/             # 邮件提取日志
      ├── imap2eml/
      └── extract/
  ```

### 3.3 配置管理

- **SettingsService**: 读取 `settings.txt`（账户配置）
- **ConfigService**: 读取 `registration.inf`（注册信息）
- **注册信息双重存储**: INF 文件 + HKLM 注册表

---

## 四、注册与激活

### 4.1 云端注册 API

- **API 地址**:
  - 安装: `https://www.booming.one/api/install`
  - 激活: `https://www.booming.one/api/activate-by-code`
  - 检查状态: `https://www.booming.one/api/install/check`

- **软件名称**: `xiaomingMailtoolkitCompany`（云端注册用）

### 4.2 注册流程

1. 启动时检查注册表并与云端同步注册状态
2. 未注册时自动提交注册
3. 已注册用户可使用"激活软件"功能（ActivationForm）
4. 未注册用户使用"注册软件"功能（RegistrationForm）

### 4.3 剩余天数计算

- 优先使用 `activation.expireDate`（已购买/激活用户）
- 回退使用 `installation.expireDate`（仅注册试用用户）
- 当服务器返回 `remainingDays=0` 时，从 `expireDate` 重算

---

## 五、Python 环境

### 5.1 脚本目录

- **位置**: `src/MailConverter/script/`
- **Python 版本**: 3.10+
- **编码**: UTF-8（通过 `PYTHONIOENCODING=UTF-8` 环境变量设置）

### 5.2 主要脚本

| 脚本 | 用途 |
|------|------|
| `create_pst.py` | 创建 PST 文件（处理邮件） |
| `convert_ost.py` | 转换 OST 文件 |
| `check_env.py` | 检查 Python 环境 |
| `deliver_mail.py` | 投递邮件 |
| `ews_deliver.py` | EWS 投递功能 |
| `extract_emails.py` | 提取邮件 |
| `extract_emails_to_msg.py` | 提取邮件到 MSG 格式 |
| `generate_key.py` | 生成密钥 |
| `check_ost.py` / `check_pst.py` | 环境检查 |

---

## 六、UI/UX 特性

### 6.1 侧边栏导航

- 现代化可折叠动画侧边栏
- Bootstrap Icons 图标支持
- Segoe MDL2 Assets 内置图标作为默认

### 6.2 窗体特性

- DPI 自适应字体缩放
- 智能窗体尺寸（屏幕 80%，保底 1024x768）
- 工业级按钮样式（TextImageRelation.ImageBeforeText、AutoEllipsis）
- 进度条移至窗体底部，新增"进度"标签

### 6.3 DPI 缩放

- ⚠️ 当前存在问题：管理员模式下运行窗口启动时很小
- 已尝试方案：app.manifest DPI 配置、SetProcessDPIAware()、SetProcessDpiAwareness(2)
- **当前状态**: 代码已还原，等待进一步排查

---

## 七、构建与发布

### 7.1 构建配置

- **目标框架**: .NET Framework 4.8
- **输出类型**: WinExe
- **AssemblyName**: MailConverter
- **ApplicationIcon**: app.ico

### 7.2 依赖包

| 包 | 版本 |
|----|------|
| MimeKit | 4.7.0 |
| MailKit | 4.7.0 |
| Microsoft.Graph | 5.57.0 |
| Microsoft.Exchange.WebServices.NETStandard | 1.1.3 |
| Serilog | 2.12.0 |
| ClosedXML | 0.102.2 |
| CsvHelper | 31.0.0 |

### 7.3 发布注意事项

- Python 环境位于 `src/MailConverter/python/`（项目目录内）
- Bootstrap Icons 位于项目根目录 `Icon/` 文件夹
- 删除 `bin/` 目录会丢失运行时生成的用户配置（config/*.inf, settings.txt 等）

---

## 八、待优化项

### 8.1 高优先级

- [ ] DPI 缩放窗口变小问题（管理员模式下）
- [ ] 注册/激活表单体验优化

### 8.2 中优先级

- [ ] 日历导出性能优化
- [ ] 大文件 PST 处理优化

### 8.3 低优先级

- [ ] 配置文件加密
- [ ] 多语言支持

---

## 九、版本历史

| 版本 | 日期 | 主要变更 |
|------|------|----------|
| v1.1.8 | 2026-05-08 | 新增 Exchange Online 邮件搜索导出删除功能（Graph API） |
| v1.1.7 | 2026-04-06 | 修复 Exchange EWS URL 自动添加、EWS 投递增强 |
| v1.1.6 | 2026-04-05 | EWS 投递到 IMAP 功能、MailKit 升级到 4.7.0 |
| v1.1.5 | 2026-04-04 | 现代化侧边栏、Bootstrap Icons、DPI 自适应 |
| v1.1.4 | 2026-04-03 | 进度条改进、启动时注册检查、RegistryService 增强 |