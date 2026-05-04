# MailConverter 邮件转换工具

**生成时间**: 2026-03-20

## 项目概述

MailConverter 是一款多功能的邮件转换工具，支持将各种格式的邮件转换到 PST 文件或同步到 Exchange/Office 365。

## 功能特性

### 1. EML 转 PST
- 将 Foxmail 导出的 EML 目录批量转换为 PST 文件
- 支持递归扫描子目录
- 保留原始文件夹结构
- 多线程并行转换

### 2. OST 转 PST
- 将 Outlook OST 文件转换为 PST 文件
- 支持保留文件夹结构
- 使用 Outlook COM 接口进行转换

### 3. 提取邮件
- 从 PST/OST 文件中提取邮件
- 导出为 EML 格式
- 保留原始文件夹结构

### 4. IMAP 收件
- 通过 IMAP 协议从邮件服务器获取邮件
- 支持自动发现常见邮件服务商配置（Gmail, Outlook, QQ, 163 等）
- 支持 SSL 连接
- 可选择特定文件夹
- 支持时间范围过滤

### 5. IMAP 多线程收件
- 支持多个 IMAP 账户并行收信
- 每个账户独立 PST 文件
- 可配置最大并行数
- 实时显示进度

### 6. 实时收件
- 收到一封立即投递一封到 PST
- 实时处理，无需等待全部下载完成

### 7. Excel/CSV 导入
- 从 Excel 文件导入联系人/邮件到 PST
- 从 CSV 文件导入
- 支持自定义列映射

### 8. IMAP 转 EWS 同步
- 从 IMAP 源服务器读取邮件
- 同步写入 Exchange/Office 365 (EWS)
- **支持按源文件夹结构创建目标文件夹**
- 多账户并行同步
- 自动创建不存在的目标文件夹

## 系统要求

### 必需组件
- Windows 10/11
- .NET Framework 4.8
- Microsoft Outlook 2016/2019/365 (32位或64位)
- Python 3.8+ (用于某些转换功能)

### Python 依赖
```bash
pip install pywin32
pip install openpyxl
pip install exchangelib
```

## 项目结构

```
src/MailConverter/
├── MainForm.cs                    # 主界面 (WinForms)
├── Program.cs                     # 程序入口
├── Services/
│   ├── AutoDiscoverService.cs           # IMAP 自动发现服务
│   ├── ImapFetchService.cs              # 单账户 IMAP 收信
│   ├── MultiThreadImapService.cs        # 多线程 IMAP 收信
│   ├── ImapToEwsService.cs              # IMAP 转 EWS 同步
│   └── ExcelCsvImportService.cs         # Excel/CSV 导入
├── emlToPst/                    # EML 转 PST 目录
│   └── create_pst.py
├── ostToPst/                    # OST 转 PST 目录
│   └── convert_ost.py
├── imapToPst/                   # IMAP 转 PST 目录
│   └── create_pst.py
├── extract_emails.py            # 提取邮件到 EML
├── deliver_mail.py              # 实时投递邮件到 PST
├── ews_deliver.py              # EWS 投递到 Exchange/Office 365
├── check_env.py                 # 环境检查
├── check_ost.py                 # 检查 OST 文件
├── check_pst.py                 # 检查 PST 文件
├── generate_readme.py           # 生成自述文件
└── MailConverter.csproj        # 项目文件
```

## 使用说明

### EML 转 PST
1. 点击 "EML转PST" 标签页
2. 选择源目录（包含 EML 文件的文件夹）
3. 选择输出 PST 文件路径
4. 设置选项（递归扫描、保留文件夹结构、线程数）
5. 点击 "开始转换"

### OST 转 PST
1. 点击 "OST转PST" 标签页
2. 选择 OST 文件
3. 选择输出 PST 文件路径
4. 点击 "开始转换"

### IMAP 收件
1. 点击 "IMAP收件" 标签页
2. 输入邮件服务器地址、端口、用户名、密码
3. 或点击 "自动发现" 选择常用服务商
4. 选择要获取的文件夹
5. 设置时间范围
6. 点击 "开始获取"

### IMAP 多线程收件
1. 点击 "IMAP多线程收件" 标签页
2. 添加多个账户（服务器、端口、用户名、密码）
3. 设置每个账户的输出 PST 文件
4. 设置最大并行数
5. 点击 "开始批量收件"

### IMAP 转 EWS 同步
1. 点击 "IMAP→EWS同步" 标签页
2. 配置源 IMAP 账户（读取邮件的服务器）
3. 配置目标 EWS 账户（写入邮件的服务器）
4. 设置目标文件夹路径
5. **注意：目标文件夹会自动按照源 IMAP 文件夹结构创建**
   - 例如：源 IMAP 有 "Inbox"、"Sent" 文件夹
   - 目标会创建 "Archive/Inbox"、"Archive/Sent"
6. 点击 "开始同步"

### Excel/CSV 导入
1. 点击 "Excel/CSV导入" 标签页
2. 选择导入文件（xlsx 或 csv）
3. 选择目标 PST 文件
4. 设置列映射
5. 点击 "开始导入"

## IMAP 自动发现支持的服务商

| 服务商 | 服务器 | 端口 | SSL |
|--------|--------|------|-----|
| Gmail | imap.gmail.com | 993 | 是 |
| Outlook | outlook.office365.com | 993 | 是 |
| QQ 邮箱 | imap.qq.com | 993 | 是 |
| 163 邮箱 | imap.163.com | 993 | 是 |
| 126 邮箱 | imap.126.com | 993 | 是 |
| 企业邮箱 | (自动检测) | - | 是 |

## 技术细节

### 多线程处理
- 使用 SemaphoreSlim 控制并发数
- 每个线程独立连接和操作
- 线程安全的进度回调

### EWS 投递
- 使用 exchangelib 库连接 Office 365
- 支持自动发现
- 支持嵌套文件夹创建（使用 "/" 分隔路径）
- 失败时自动尝试 Outlook COM 备选方案

### 文件夹映射规则

在 IMAP 转 EWS 同步时，目标文件夹按以下规则创建：

```
源 IMAP 文件夹                    目标 EWS 文件夹
─────────────────────────────────────────────────
Inbox                        →    Archive/Inbox (如果设置了Archive)
Sent                         →    Archive/Sent
Drafts                       →    Archive/Drafts
自定义文件夹                   →    Archive/自定义文件夹

如果未设置目标文件夹根目录，则直接使用源文件夹名：
Inbox                        →    Inbox
Sent                         →    Sent
```

## 配置文件

### 环境检查
运行 `check_env.py` 检查环境配置：
```bash
python check_env.py
```

### 日志文件
日志保存在 `logs/mailconverter.log`

## 注意事项

1. **Outlook 运行时**: 某些功能需要 Outlook 在后台运行
2. **权限**: 确保有足够的权限访问 PST 文件和邮件服务器
3. **网络**: IMAP/EWS 功能需要网络连接
4. **大文件**: 处理大量邮件时可能需要较长时间

## 常见问题

### Q: 转换失败怎么办？
A: 检查日志文件 `logs/mailconverter.log` 获取详细错误信息

### Q: EWS 投递失败怎么办？
A:
- 确认 EWS 账户有写入权限
- 检查服务器 URL 是否正确
- 确认目标文件夹路径格式正确（使用 "/" 分隔）

### Q: 如何查看进度？
A: 界面底部显示进度条和实时日志

## 版本信息

- 当前版本: 1.0.1
- 开发框架: .NET Framework 4.8 + WinForms
- 核心依赖:
  - MimeKit 4.0.0 (EML 解析)
  - MailKit 4.0.0 (IMAP 客户端)
  - ClosedXML 0.102.2 (Excel 处理)
  - CsvHelper 31.0.0 (CSV 处理)

---

# EML/IMAP 转 PST 技术文档 (2026-03-20)

## 问题背景

在将 EML 文件或 IMAP 邮件导入 PST 时，遇到了以下问题：
- **邮件处于"可编辑/草稿"状态** - 双击邮件像是写新邮件，而不是已接收邮件
- **发件人信息丢失** - 显示为"未知发件人"
- **时间戳丢失** - 发送/接收时间显示为导入时间，而非原始时间

## 根本原因

Outlook COM 自动化创建新邮件时，默认会创建"草稿"状态的邮件项目，需要通过 MAPI 属性来修复。

## 解决方案

### 核心技术要点

1. **草稿箱中转**
   ```python
   local_drafts = ns.GetDefaultFolder(16)  # 16 = 草稿箱
   mail = local_drafts.Items.Add(0)
   ```

2. **super_fix_sender() 函数 - 双重修复**
   - 在 `mail.Save()` **之前**调用一次
   - 在 `mail.Move(target_folder)` **之后**再调用一次
   - 关键：移动操作会重置属性，必须在目的地再刷一遍

3. **完整的 MAPI 属性设置**
   ```python
   properties = {
       "0x0C1A001F": from_str,  # PR_SENDER_NAME_W
       "0x0C1F001F": from_str,  # PR_SENDER_EMAIL_ADDRESS_W
       "0x0C1E001F": "SMTP",    # PR_SENDER_ADDRTYPE_W
       "0x0042001F": from_str,  # PR_SENT_REPRESENTING_NAME_W
       "0x0065001F": from_str,  # PR_SENT_REPRESENTING_EMAIL_ADDRESS_W
       "0x0064001F": "SMTP",    # PR_SENT_REPRESENTING_ADDRTYPE_W
   }
   ```

4. **修复草稿标志**
   ```python
   prop_accessor.SetProperty(f"{tag_prefix}0x0E070003", 1)  # MSGFLAG_READ=1
   ```

5. **时间戳设置**
   ```python
   prop_accessor.SetProperty(f"{tag_prefix}0x00390040", dt)  # PR_CLIENT_SUBMIT_TIME
   prop_accessor.SetProperty(f"{tag_prefix}0x0E060040", dt)  # PR_MESSAGE_DELIVERY_TIME
   ```

## 项目目录结构

```
src/MailConverter/
├── emlToPst/              # EML 转 PST 脚本
│   └── create_pst.py
├── ostToPst/              # OST 转 PST 脚本
│   └── convert_ost.py
├── imapToPst/             # IMAP 转 PST 脚本 (与 EML 共用)
│   └── create_pst.py
└── ...
```

**注意**: C# 代码已更新为指向新的子目录：
- `emlToPst/create_pst.py`
- `ostToPst/convert_ost.py`
- `imapToPst/create_pst.py`

## Python 依赖

内置 Python 3.10.6 环境，位于 `python/` 目录：
- pywin32 (COM 接口)
- dateutil (日期解析)
- email (EML 解析)

## 关键 MAPI 属性参考

| 属性名 | Tag | 说明 |
|--------|-----|------|
| PR_MESSAGE_FLAGS | 0x0E070003 | 邮件标志，设为1=已读 |
| PR_SENDER_NAME_W | 0x0C1A001F | 发件人显示名 |
| PR_SENDER_EMAIL_ADDRESS_W | 0x0C1F001F | 发件人邮箱 |
| PR_SENDER_ADDRTYPE_W | 0x0C1E001F | 发件人地址类型 (SMTP) |
| PR_CLIENT_SUBMIT_TIME | 0x00390040 | 客户端提交时间 |
| PR_MESSAGE_DELIVERY_TIME | 0x0E060040 | 邮件接收时间 |

## 相关日志

日志文件: `logs/mailconverter.log`

成功导入的关键日志输出：
```
Outlook connected
Creating PST...
Using drafts as intermediate
Total emails imported: XX
SUCCESS
```

## 构建项目

```bash
cd src/MailConverter
dotnet build
```

## 许可

### 商业授权

本软件为商业软件，受版权法保护。

| 授权类型 | 价格 | 说明 |
|---------|------|------|
| 个人授权 | ¥299 | 单用户永久授权 |
| 企业授权 | ¥2999 | 企业内无限用户 |
| 定制服务 | 另议 | 源代码定制和技术支持 |

详细条款请参阅 [LICENSE_ZH.md](LICENSE_ZH.md) (中文) 或 [LICENSE.md](LICENSE.md) (English)。

### 第三方组件许可

本软件使用的开源组件遵循其各自的开源许可证（MIT、Apache 2.0、BSD 等），允许商业使用。
