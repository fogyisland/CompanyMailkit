# 单用户同步 Tab 拆分 (联系人 / 日历分离)

**日期**: 2026-06-02
**状态**: 设计已批准, 待用户审 spec 后写 plan

## 背景

"单用户同步（O365)" Tab 内的"同步联系人/日历"页签将两类功能合并在一个 `Panel` 中
(`MainForm.CreateSyncContactsCalendarPanel`, 11725-12300 行), 体积大、控件多、事件处理
分支复杂 (按 `cmbSyncType.SelectedIndex` 切换联系人/日历两套逻辑), 不利于维护和测试。

联系人代码与日历代码在以下服务中并存:
- `Services/PstExtractService.cs` (1083 行) - 同时含 PST 联系人 / 日历提取
- `Services/Office365ImportService.cs` (2755 行) - 同时含 Graph 联系人 / 日历同步

## 目标

1. **UI 拆分**: "同步联系人/日历" Tab 拆为两个独立 Tab, 各自独立登录
2. **代码拆分**: 联系人 / 日历的数据类 + 服务代码彻底分离
3. **保留兼容**: PST 批量同步 Tab (`CreatePstSyncContactsPanel` / `CreatePstSyncCalendarPanel`) 保持现状
4. **不破坏现有调用点**: MainForm 中使用 `PstExtractService` / `Office365ImportService` 的旧代码,
   通过 facade 转发到新服务, 改动最小

## 文件结构

```
src/MailConverter/
├── Services/
│   ├── Contacts/                          [新增子目录]
│   │   ├── ContactData.cs                 [新, 从 PstExtractService 抽出 PstContactData]
│   │   ├── ContactSyncService.cs          [新, 整合 PST 提取 + Graph 同步]
│   │   ├── ContactSelectionForm.cs        [从根目录移入]
│   │   └── SyncContactsControl.cs         [新, UserControl 拆分自原面板]
│   ├── Calendars/                         [新增子目录]
│   │   ├── CalendarData.cs                [新, 从 PstExtractService 抽出 PstCalendarData]
│   │   ├── CalendarSyncService.cs         [新, 整合 PST 提取 + Graph 同步]
│   │   ├── CalendarSelectionForm.cs       [新, 镜像 ContactSelectionForm]
│   │   └── SyncCalendarControl.cs         [新, UserControl 拆分自原面板]
│   ├── PstExtractService.cs               [留 facade, 联系人/日历代码删除]
│   └── Office365ImportService.cs          [留 facade, 联系人/日历代码删除]
└── MainForm.cs
```

## 重命名

| 旧名 | 新名 | 新位置 |
|---|---|---|
| `PstExtractService.PstContactData` | `ContactData` | `Services/Contacts/ContactData.cs` |
| `PstExtractService.PstCalendarData` | `CalendarData` | `Services/Calendars/CalendarData.cs` |

引用点 (`MainForm.cs`, `Office365ImportService.cs`) 全部更新.

## Facade 转发

`PstExtractService` 中联系人/日历相关方法保留为薄 facade, 内部委托到新服务:

| 旧方法 (PstExtractService) | 转发到 |
|---|---|
| `ExtractContactsToVcf` | `ContactSyncService.ExtractContactsToVcf` |
| `ExtractContactsFromPst` | `ContactSyncService.ExtractContactsFromPst` |
| `ExtractCalendarToIcs` | `CalendarSyncService.ExtractCalendarToIcs` |
| `PstContactData` 类型引用 | `ContactData` (MailConverter.Services.Contacts) |
| `PstCalendarData` 类型引用 | `CalendarData` (MailConverter.Services.Calendars) |

`Office365ImportService` 同理:

| 旧方法 (Office365ImportService) | 转发到 |
|---|---|
| `ImportContactsToGraph` | `ContactSyncService.ImportContactsToGraph` |
| `ImportSingleContactDirectWithRetryAsync` | `ContactSyncService.ImportSingleContactDirectWithRetryAsync` |
| `ConvertToGraphContact` | `ContactSyncService.ConvertToGraphContact` |
| `ImportCalendarsToGraph` | `CalendarSyncService.ImportCalendarsToGraph` |
| `ImportSingleCalendarDirectWithRetryAsync` | `CalendarSyncService.ImportSingleCalendarDirectWithRetryAsync` |
| `ConvertToGraphEvent` | `CalendarSyncService.ConvertToGraphEvent` |

## UI 拆分 (单用户同步)

### 当前
```
单用户同步（O365)
├── EML导入
├── PST导入
└── 同步联系人/日历  ← 合并面板
```

### 目标
```
单用户同步（O365)
├── EML导入
├── PST导入
├── 同步联系人      ← 独立 Tab, 独立登录
└── 同步日历        ← 独立 Tab, 独立登录
```

### 控件拆分

`SyncContactsControl` 包含:
- 认证 (Client ID / 租户 / 邮箱 / 登录按钮 / 当前登录邮箱)
- 同步类型: 固定 "同步个人通讯录"
- 数据来源: CSV / VCF / CardDAV / 企业微信API
- 源文件 / 服务器 / 用户名 / 密码
- 选择联系人按钮 (CardDAV 模式显示)
- 增量同步按钮 (CardDAV 模式显示)
- 全量同步按钮
- 进度条 + 状态标签

`SyncCalendarControl` 包含:
- 认证 (独立登录)
- 同步类型: 固定 "同步日历"
- 数据来源: ICS / VCS / CalDAV
- 源文件 / 服务器 / 用户名 / 密码
- 全量同步按钮
- 进度条 + 状态标签

## 范围

### 包含
- UI 拆分 (单用户同步 Tab)
- 数据类重命名 + 文件移动
- 新建 ContactSyncService / CalendarSyncService
- 旧服务留 facade
- MainForm 引用点更新
- 新建 CalendarSelectionForm

### 不包含
- PST 批量同步 Tab 改动 (保持现状, 通过 facade 复用新服务)
- `ParseVcfToContacts` (MainForm 私有, 保持原位)
- 邮件相关代码
- 任何功能行为变更 (纯结构重构)

## 验证

1. 编译 0 错误
2. UI 启动后, "单用户同步" 下显示 4 个 Tab: EML导入 / PST导入 / 同步联系人 / 同步日历
3. 同步联系人 Tab: 独立登录, 数据源限定 CSV/VCF/CardDAV/企业微信
4. 同步日历 Tab: 独立登录, 数据源限定 ICS/VCS/CalDAV
5. 现有 PST 批量同步功能不受影响
