# PST 邮件转换技术文档

**日期**: 2026-03-20
**问题**: EML/IMAP 导入 PST 后邮件显示为"可编辑/草稿"状态

---

## 问题现象

1. 邮件图标显示为"打开的信封"（草稿状态）
2. 双击邮件后显示"发送"按钮，而非"回复/回复全部"
3. 发件人显示为"未知发件人"
4. 时间戳为导入时间，而非原始发送/接收时间

---

## 根本原因

Outlook COM 自动化使用 `Items.Add(0)` 创建新邮件时，默认创建的是"草稿"（IPM.Note）项目，具有以下特征：
- `MSGFLAG_UNSENT` (0x08) 标志被设置
- 未设置发件人 MAPI 属性
- 时间戳为创建时间

---

## 解决方案

### 1. 使用草稿箱中转

```python
# 获取草稿箱
local_drafts = ns.GetDefaultFolder(16)  # 16 = olFolderDrafts

# 在草稿箱创建邮件
mail = local_drafts.Items.Add(0)
```

### 2. super_fix_sender() 函数

```python
def super_fix_sender(mail_item, msg):
    """最强力的发件人注入"""
    prop_accessor = mail_item.PropertyAccessor
    tag_prefix = "http://schemas.microsoft.com/mapi/proptag/"

    from_str = str(msg.get('from', ''))

    # 设置完整的发件人信息
    properties = {
        f"{tag_prefix}0x0C1A001F": from_str,  # 发件人姓名
        f"{tag_prefix}0x0C1F001F": from_str,  # 发件人地址
        f"{tag_prefix}0x0C1E001F": "SMTP",   # 地址类型
        f"{tag_prefix}0x0042001F": from_str,  # 代表姓名
        f"{tag_prefix}0x0065001F": from_str,  # 代表地址
        f"{tag_prefix}0x0064001F": "SMTP",   # 代表地址类型
    }

    for tag, value in properties.items():
        try:
            prop_accessor.SetProperty(tag, value)
        except: pass

    # 关键：移除草稿标志 (MSGFLAG_READ=1 清除 UNSENT)
    prop_accessor.SetProperty(f"{tag_prefix}0x0E070003", 1)

    # 设置时间戳
    date_str = msg.get('date')
    if date_str:
        dt = parsedate_to_datetime(str(date_str))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        prop_accessor.SetProperty(f"{tag_prefix}0x00390040", dt)  # 发送时间
        prop_accessor.SetProperty(f"{tag_prefix}0x0E060040", dt)  # 接收时间

    mail_item.Save()
```

### 3. 双重修复（关键！）

```python
# 1. 在草稿箱创建并修复
mail = local_drafts.Items.Add(0)
# ... 设置主题、正文、附件 ...
super_fix_sender(mail, msg)  # 第一次修复
mail.Save()

# 2. 移动到目标文件夹
moved_mail = mail.Move(target_folder)

# 3. 【关键】移动后再修复一次！
super_fix_sender(moved_mail, msg)  # 第二次修复
```

**为什么需要两次？**
- `Move()` 操作会重置某些 MAPI 属性
- 在目标位置再次修复确保属性持久化

---

## MAPI 属性参考

### 发件人相关

| 属性名 | Tag | 类型 | 说明 |
|--------|-----|------|------|
| PR_SENDER_NAME | 0x0C1A001F | Unicode | 发件人显示名 |
| PR_SENDER_EMAIL_ADDRESS | 0x0C1F001F | Unicode | 发件人邮箱地址 |
| PR_SENDER_ADDRTYPE | 0x0C1E001F | Unicode | 地址类型 (SMTP/EXCHANGE) |
| PR_SENT_REPRESENTING_NAME | 0x0042001F | Unicode | 代表发件人姓名 |
| PR_SENT_REPRESENTING_EMAIL_ADDRESS | 0x0065001F | Unicode | 代表发件人邮箱 |
| PR_SENT_REPRESENTING_ADDRTYPE | 0x0064001F | Unicode | 代表发件人地址类型 |

### 时间相关

| 属性名 | Tag | 类型 | 说明 |
|--------|-----|------|------|
| PR_CLIENT_SUBMIT_TIME | 0x00390040 | FILETIME | 客户端提交时间（发送时间） |
| PR_MESSAGE_DELIVERY_TIME | 0x0E060040 | FILETIME | 邮件接收时间 |

### 标志相关

| 属性名 | Tag | 类型 | 说明 |
|--------|-----|------|------|
| PR_MESSAGE_FLAGS | 0x0E070003 | Integer | 邮件标志 |
| MSGFLAG_READ | 0x00000001 | Bit | 已读标志 |
| MSGFLAG_UNSENT | 0x00000008 | Bit | 未发送/草稿标志 |

---

## 完整工作流程

```python
def create_pst_and_import(pst_path, eml_dir):
    # 1. 初始化 Outlook COM
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")

    # 2. 创建 PST 文件
    ns.AddStore(pst_path)

    # 3. 获取草稿箱和 PST 根文件夹
    local_drafts = ns.GetDefaultFolder(16)
    pst_root = find_pst_root(ns, pst_path)

    # 4. 遍历 EML 文件
    for eml_file in eml_files:
        msg = parse_eml(eml_file)

        # 5. 在草稿箱创建邮件
        mail = local_drafts.Items.Add(0)
        mail.Subject = msg['subject']
        mail.Body = msg.get_body()
        # ... 设置其他属性 ...

        # 6. 第一次修复
        super_fix_sender(mail, msg)
        mail.Save()

        # 7. 移动到目标文件夹
        target_folder = get_or_create_folder(pst_root, rel_path)
        moved_mail = mail.Move(target_folder)

        # 8. 第二次修复（关键！）
        super_fix_sender(moved_mail, msg)
```

---

## 调试技巧

### 检查邮件属性

```python
# 在 Outlook VBA 中检查
Sub CheckMailProperties()
    Dim mail As Outlook.MailItem
    Set mail = ActiveExplorer.Selection(1)

    Dim propAccessor As Outlook.PropertyAccessor
    Set propAccessor = mail.PropertyAccessor

    Debug.Print "Sender: " & mail.Sender
    Debug.Print "SentOn: " & mail.SentOn
    Debug.Print "ReceivedTime: " & mail.ReceivedTime
    Debug.Print "MessageClass: " & mail.MessageClass

    ' 检查 MAPI 属性
    On Error Resume Next
    Debug.Print "PR_MESSAGE_FLAGS: " & _
        propAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E070003")
End Sub
```

### 常见错误

1. **服务器运行失败** (-2146959355)
   - 原因：Outlook COM 对象被占用
   - 解决：关闭所有 Outlook 进程

2. **属性设置失败** (-2147024891)
   - 原因：某些属性在特定状态下不可写
   - 解决：使用 try-except 捕获，继续执行

---

## 文件位置

```
src/MailConverter/
├── emlToPst/create_pst.py      # EML 转 PST (最终版)
├── ostToPst/convert_ost.py      # OST 转 PST
└── imapToPst/create_pst.py     # IMAP 转 PST (复用 EML 脚本)
```

备份版本：
```
backup_scripts/
├── create_pst_2026-03-19_eml_worked.py
├── create_pst_2026-03-19_v23_state_and_time_fixed.py
├── create_pst_2026-03-19_v25_final.py
├── create_pst_2026-03-19_v26.py
└── create_pst_2026-03-20_perfect.py
```

---

## 参考资料

- [Outlook MSGFLAG 常量](https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders)
- [MAPI 属性列表](https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/mapi-properties)
- [PropertyAccessor 对象](https://docs.microsoft.com/en-us/office/vba/api/outlook.propertyaccessor)
