# 单用户同步 Tab 拆分 - 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将"单用户同步"Tab 内的"同步联系人/日历"合并面板拆为两个独立 Tab (联系人/日历), 联系人与日历的数据类和服务代码彻底分离, 旧服务通过 facade 转发保持兼容.

**Architecture:**
- 数据类: `PstContactData` / `PstCalendarData` 抽到独立文件, 重命名为 `ContactData` / `CalendarData`
- 服务: 新建 `ContactSyncService` / `CalendarSyncService` 整合提取 + 同步逻辑 (static 类, 接 GraphServiceClient 参数)
- 旧服务: `PstExtractService` / `Office365ImportService` 保留方法签名, 内部委托到新服务
- UI: 拆为 `SyncContactsControl` / `SyncCalendarControl` (UserControl), MainForm 加两个 Tab

**Tech Stack:** .NET Framework 4.8, WinForms, Microsoft.Graph 5.57.0, Microsoft.Office.Interop.Outlook

---

## 任务依赖图

```
Task 1 → Task 3
Task 2 → Task 4
Task 3, 4 → Task 5
Task 3, 4 → Task 6
Task 1, 2 → Task 7
Task 7 → Task 12
Task 8 → Task 12
Task 9 → Task 12
Task 10, 11 → Task 12
Task 12 → Task 13
```

---

## Task 1: 抽 `ContactData` 到独立文件

**Files:**
- Create: `src/MailConverter/Services/Contacts/ContactData.cs`
- Modify: `src/MailConverter/Services/PstExtractService.cs:479-500` (删除内嵌类)

- [ ] **Step 1: 创建 `Services/Contacts/ContactData.cs`**

完整内容:

```csharp
using System;

namespace MailConverter.Services.Contacts
{
    /// <summary>
    /// PST 联系人数据模型 (从 Outlook ContactItem 抽取后用于 Graph 同步)
    /// </summary>
    public class ContactData
    {
        public string DisplayName { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string MiddleName { get; set; }
        public string Title { get; set; }
        public string Suffix { get; set; }
        public string Email { get; set; }
        public string Email2 { get; set; }
        public string Email3 { get; set; }
        public string Phone { get; set; }
        public string Phone2 { get; set; }
        public string MobilePhone { get; set; }
        public string CompanyName { get; set; }
        public string Department { get; set; }
        public string JobTitle { get; set; }
        public string BusinessAddress { get; set; }
        public string HomeAddress { get; set; }
        public string PersonalNotes { get; set; }
        public DateTime? Birthday { get; set; }
    }
}
```

- [ ] **Step 2: 从 `PstExtractService.cs` 删除内嵌 `PstContactData` 类**

删除 479-500 行的 `public class PstContactData { ... }` 整段 (含前面的 `/// <summary>` 注释).

- [ ] **Step 3: 构建验证**

```bash
dotnet build src/MailConverter/MailConverter.csproj -c Debug -v minimal
```

预期: 大量编译错误, 因为 `PstExtractService.PstContactData` 引用全部失效. **这是预期的**, 暂时接受, 后续 Task 会修.

- [ ] **Step 4: 提交**

```bash
git add src/MailConverter/Services/Contacts/ContactData.cs src/MailConverter/Services/PstExtractService.cs
git commit -m "refactor: 抽出 ContactData 到独立文件 (PstContactData 暂时未替换)"
```

---

## Task 2: 抽 `CalendarData` 到独立文件

**Files:**
- Create: `src/MailConverter/Services/Calendars/CalendarData.cs`
- Modify: `src/MailConverter/Services/PstExtractService.cs:502-521` (删除内嵌类)

- [ ] **Step 1: 创建 `Services/Calendars/CalendarData.cs`**

```csharp
using System;

namespace MailConverter.Services.Calendars
{
    /// <summary>
    /// PST 日历数据模型 (从 Outlook AppointmentItem 抽取后用于 Graph 同步)
    /// </summary>
    public class CalendarData
    {
        public string Subject { get; set; }
        public string Body { get; set; }
        public DateTime? StartTime { get; set; }
        public DateTime? EndTime { get; set; }
        public string Location { get; set; }
        public bool IsAllDayEvent { get; set; }
        public string ReminderMinutesBeforeStart { get; set; }
        public bool ReminderSet { get; set; }
        public string Categories { get; set; }
        public string RequiredAttendees { get; set; }
        public string OptionalAttendees { get; set; }
        public string ResourceAttendees { get; set; }
        public bool IsRecurring { get; set; }
        public string RecurrencePattern { get; set; }
    }
}
```

- [ ] **Step 2: 从 `PstExtractService.cs` 删除内嵌 `PstCalendarData` 类**

删除 502-521 行 (含 `/// <summary>`) 整段.

- [ ] **Step 3: 提交**

```bash
git add src/MailConverter/Services/Calendars/CalendarData.cs src/MailConverter/Services/PstExtractService.cs
git commit -m "refactor: 抽出 CalendarData 到独立文件"
```

---

## Task 3: 创建 `ContactSyncService`

**Files:**
- Create: `src/MailConverter/Services/Contacts/ContactSyncService.cs`

- [ ] **Step 1: 创建 `ContactSyncService.cs`**

将 `PstExtractService` 的联系人相关方法 (含 `ExtractContactsToVcf`, `ExtractContactsFromPst`, `ExtractContactsRecursive`, `ExtractContactsRecursiveToMemory`) 和 `Office365ImportService` 的 `ImportContactsBatchDirectAsync`, `ImportSingleContactDirectWithRetryAsync`, `ConvertToGraphContact` 全部移到此处, 类型 `PstContactData` 替换为 `ContactData` (`using MailConverter.Services.Contacts;`).

具体行数参考:
- `ExtractContactsToVcf` (PstExtractService.cs:421-474)
- `ExtractContactsFromPst` (PstExtractService.cs:526-605)
- `ExtractContactsRecursiveToMemory` (PstExtractService.cs:607-680)
- `ExtractContactsRecursive` (PstExtractService.cs:681-730)
- `ImportContactsBatchDirectAsync` (Office365ImportService.cs:2160-2249)
- `ImportSingleContactDirectWithRetryAsync` (Office365ImportService.cs:2250-2285)
- `ConvertToGraphContact` (Office365ImportService.cs:2290-2400)

签名: `ConvertToGraphContact` 改为 `public static Microsoft.Graph.Models.Contact ConvertToGraphContact(ContactData data)` (从 private 改 public static).

类签名:

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Graph;
using Outlook = Microsoft.Office.Interop.Outlook;
using Serilog;

namespace MailConverter.Services.Contacts
{
    public static class ContactSyncService
    {
        // ExtractContactsToVcf, ExtractContactsFromPst, ExtractContactsRecursive,
        // ExtractContactsRecursiveToMemory, ImportContactsBatchDirectAsync,
        // ImportSingleContactDirectWithRetryAsync, ConvertToGraphContact
        // (所有方法签名中 PstContactData → ContactData, 补 GraphServiceClient 参数)
    }
}
```

注意: 移到 static 类后, 原实例方法需要:
1. `Log.Information(...)` 改为 `Log.Information(...)` (Serilog 静态访问仍然 OK)
2. `Program.BatchToO365Logger.*` 保持原样 (静态)
3. 接收 `GraphServiceClient graphClient` 作为参数, 不再使用 `this._graphClient`

签名示例:

```csharp
public static async Task<int> ImportContactsBatchDirectAsync(
    GraphServiceClient graphClient,
    string targetEmail,
    IEnumerable<ContactData> contacts,
    Action<int, int, string> progressCallback = null,
    int maxDegreeOfParallelism = 10)
```

- [ ] **Step 2: 提交 (不构建)**

```bash
git add src/MailConverter/Services/Contacts/ContactSyncService.cs
git commit -m "refactor: 新建 ContactSyncService (整合联系人提取+同步)"
```

预期编译错误, 因为 `PstExtractService.ExtractContactsToVcf` 等已删除但仍被引用. 后续 Task 5 修.

---

## Task 4: 创建 `CalendarSyncService`

**Files:**
- Create: `src/MailConverter/Services/Calendars/CalendarSyncService.cs`

- [ ] **Step 1: 创建 `CalendarSyncService.cs`**

将 `PstExtractService` 的日历相关方法 (`ExtractCalendarToIcs`, `ExtractCalendarRecursive`, `SaveAppointmentAsIcs`) 和 `Office365ImportService` 的 `ImportCalendarBatchDirectAsync`, `ImportSingleCalendarDirectWithRetryAsync`, `ConvertToGraphEvent` 移到此处. 类型替换 `PstCalendarData` → `CalendarData`.

行数参考:
- `ExtractCalendarToIcs` (PstExtractService.cs:735-788)
- `ExtractCalendarRecursive` (PstExtractService.cs:790-839)
- `SaveAppointmentAsIcs` (PstExtractService.cs:841-1083)
- `ImportCalendarBatchDirectAsync` (Office365ImportService.cs:2470-2546)
- `ImportSingleCalendarDirectWithRetryAsync` (Office365ImportService.cs:2547-2633)
- `ConvertToGraphEvent` (Office365ImportService.cs:2634-end of event section)
- `GetTimeZoneFromUtcOffset` (private helper, 也移入)

类签名:

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;
using Outlook = Microsoft.Office.Interop.Outlook;
using Serilog;

namespace MailConverter.Services.Calendars
{
    public static class CalendarSyncService
    {
        // ExtractCalendarToIcs, ExtractCalendarRecursive, SaveAppointmentAsIcs,
        // ImportCalendarBatchDirectAsync, ImportSingleCalendarDirectWithRetryAsync,
        // ConvertToGraphEvent, GetTimeZoneFromUtcOffset
    }
}
```

签名示例 (与 Task 3 类似, 接收 GraphServiceClient):

```csharp
public static async Task<int> ImportCalendarBatchDirectAsync(
    GraphServiceClient graphClient,
    string targetEmail,
    IEnumerable<CalendarData> calendars,
    Action<int, int, string> progressCallback = null,
    int maxDegreeOfParallelism = 10,
    string timeZone = "China Standard Time")

public static Microsoft.Graph.Models.Event ConvertToGraphEvent(CalendarData data, string timeZone = "China Standard Time")
```

- [ ] **Step 2: 提交 (不构建)**

```bash
git add src/MailConverter/Services/Calendars/CalendarSyncService.cs
git commit -m "refactor: 新建 CalendarSyncService (整合日历提取+同步)"
```

预期编译错误. 后续 Task 5, 6 修.

---

## Task 5: `PstExtractService` 留 facade

**Files:**
- Modify: `src/MailConverter/Services/PstExtractService.cs`

- [ ] **Step 1: 删除 `PstExtractService` 中所有联系人/日历方法**

删除以下方法 (整段, 含 `/// <summary>`):
- `ExtractContactsToVcf` (421-474)
- `ExtractContactsFromPst` (526-605)
- `ExtractContactsRecursiveToMemory` (607-680)
- `ExtractContactsRecursive` (681-730)
- `ExtractCalendarToIcs` (735-788)
- `ExtractCalendarRecursive` (790-839)
- `SaveAppointmentAsIcs` (841-1083)

保留: 邮件提取方法 (`ExtractToEml` 等)

- [ ] **Step 2: 添加 facade 方法**

在 `PstExtractService` 类内 (邮件方法之后) 添加:

```csharp
using MailConverter.Services.Contacts;
using MailConverter.Services.Calendars;
using System.Collections.Generic;
using System.Runtime.InteropServices;  // 如果需要

// 联系人 facade
public bool ExtractContactsToVcf(string pstPath, string outputDir, IProgress<int> progress = null)
{
    return ContactSyncService.ExtractContactsToVcf(pstPath, outputDir, progress);
}

public List<ContactData> ExtractContactsFromPst(string pstPath, IProgress<int> progress = null)
{
    return ContactSyncService.ExtractContactsFromPst(pstPath, progress);
}

// 日历 facade
public bool ExtractCalendarToIcs(string pstPath, string outputDir, IProgress<int> progress = null)
{
    return CalendarSyncService.ExtractCalendarToIcs(pstPath, outputDir, progress);
}
```

注意: facade 方法的返回类型必须是新类型 (`ContactData` 而非 `PstContactData`). 这意味着调用点的类型需要更新 (Task 7 处理).

- [ ] **Step 3: 提交**

```bash
git add src/MailConverter/Services/PstExtractService.cs
git commit -m "refactor: PstExtractService 联系人/日历方法改为 facade 转发"
```

---

## Task 6: `Office365ImportService` 留 facade

**Files:**
- Modify: `src/MailConverter/Services/Office365ImportService.cs`

- [ ] **Step 1: 删除 `Office365ImportService` 中所有联系人/日历方法**

删除以下方法:
- `ImportContactsBatchDirectAsync` (2160-2249)
- `ImportSingleContactDirectWithRetryAsync` (2250-2285)
- `ConvertToGraphContact` (2290-2400)
- `ImportCalendarBatchDirectAsync` (2470-2546)
- `ImportSingleCalendarDirectWithRetryAsync` (2547-2633)
- `ConvertToGraphEvent` (2634-~2700)
- `GetTimeZoneFromUtcOffset` (private helper)

保留: `_graphClient` 字段, `ConnectWithClientSecret` / `ConnectWithOAuth` 等, 邮件导入方法.

- [ ] **Step 2: 添加 facade 方法**

在 `Office365ImportService` 类内 (邮件方法之后) 添加:

```csharp
using MailConverter.Services.Contacts;
using MailConverter.Services.Calendars;
using System.Collections.Generic;
using System.Threading.Tasks;

// 联系人 facade
public Task<int> ImportContactsBatchDirectAsync(
    string targetEmail,
    IEnumerable<ContactData> contacts,
    Action<int, int, string> progressCallback = null,
    int maxDegreeOfParallelism = 10)
{
    return ContactSyncService.ImportContactsBatchDirectAsync(
        _graphClient, targetEmail, contacts, progressCallback, maxDegreeOfParallelism);
}

// 日历 facade
public Task<int> ImportCalendarBatchDirectAsync(
    string targetEmail,
    IEnumerable<CalendarData> calendars,
    Action<int, int, string> progressCallback = null,
    int maxDegreeOfParallelism = 10,
    string timeZone = "China Standard Time")
{
    return CalendarSyncService.ImportCalendarBatchDirectAsync(
        _graphClient, targetEmail, calendars, progressCallback, maxDegreeOfParallelism, timeZone);
}
```

- [ ] **Step 3: 构建验证**

```bash
dotnet build src/MailConverter/MailConverter.csproj -c Debug -v minimal
```

预期: 0 错误. (MainForm 引用 PstContactData/PstCalendarData 的错误会在 Task 7 处理, 但因为 facade 方法签名已经改为新类型, MainForm 必须更新才能编译过 — 这是预期)

如果仍有 0 错误, 说明 `PstContactData` / `PstCalendarData` 引用已不存在 (可能 MainForm 已经用的是类型, 但 facade 已经更新). 此时应继续 Task 7.

如果错误仅来自 MainForm 引用 (PstContactData/PstCalendarData), 继续 Task 7.

- [ ] **Step 4: 提交**

```bash
git add src/MailConverter/Services/Office365ImportService.cs
git commit -m "refactor: Office365ImportService 联系人/日历方法改为 facade 转发"
```

---

## Task 7: 更新 MainForm 引用新类型

**Files:**
- Modify: `src/MailConverter/MainForm.cs`

- [ ] **Step 1: 查找所有 `PstContactData` / `PstCalendarData` 引用**

```bash
grep -n "PstContactData\|PstCalendarData" src/MailConverter/MainForm.cs
```

预期行号 (基于已读代码):
- 14930, 14965, 14984: 联系人相关
- 15557, 15602, 15617: 日历相关
- 15886, 15892, 15900: `ParseVcfToContacts`
- 16058, 16072: 批量处理

- [ ] **Step 2: 批量替换**

在 MainForm.cs 顶部添加 `using`:

```csharp
using MailConverter.Services.Contacts;
using MailConverter.Services.Calendars;
```

将 `PstExtractService.PstContactData` → `ContactData` (全文替换, 注意 Verify=False)
将 `PstExtractService.PstCalendarData` → `CalendarData`

```bash
# 旧 (PST 已经删, 需要 Edit tool 一个个改, 不用 sed)
```

- [ ] **Step 3: 构建验证**

```bash
dotnet build src/MailConverter/MailConverter.csproj -c Debug -v minimal
```

预期: 0 错误, 0 个引用 PstContactData/PstCalendarData 的错误.

- [ ] **Step 4: 提交**

```bash
git add src/MailConverter/MainForm.cs
git commit -m "refactor: MainForm 引用 ContactData/CalendarData 替代 PstContactData/PstCalendarData"
```

---

## Task 8: 移动 `ContactSelectionForm` 到子目录

**Files:**
- Move: `src/MailConverter/ContactSelectionForm.cs` → `src/MailConverter/Services/Contacts/ContactSelectionForm.cs`

- [ ] **Step 1: 用 git mv 移动文件**

```bash
git mv src/MailConverter/ContactSelectionForm.cs src/MailConverter/Services/Contacts/ContactSelectionForm.cs
```

- [ ] **Step 2: 修改 namespace**

在 `Services/Contacts/ContactSelectionForm.cs` 第 6 行 (namespace 声明):

```csharp
namespace MailConverter.Services.Contacts
```

替换原 `namespace MailConverter`.

- [ ] **Step 3: 更新 MainForm 引用**

在 `MainForm.cs` 顶部添加 `using MailConverter.Services.Contacts;` (如果还没有).

将 `new ContactSelectionForm(contacts)` → `new ContactSelectionForm(contacts)` (无需改, 引用类型已通过 using 找到). 验证:

```bash
grep -n "ContactSelectionForm" src/MailConverter/MainForm.cs
```

- [ ] **Step 4: 构建验证**

```bash
dotnet build src/MailConverter/MailConverter.csproj -c Debug -v minimal
```

预期: 0 错误.

- [ ] **Step 5: 提交**

```bash
git add src/MailConverter/Services/Contacts/ContactSelectionForm.cs src/MailConverter/MainForm.cs
git commit -m "refactor: 移动 ContactSelectionForm 到 Services/Contacts/"
```

---

## Task 9: 创建 `CalendarSelectionForm`

**Files:**
- Create: `src/MailConverter/Services/Calendars/CalendarSelectionForm.cs`

- [ ] **Step 1: 读取 `ContactSelectionForm.cs` 作为模板**

```bash
cat src/MailConverter/Services/Contacts/ContactSelectionForm.cs
```

- [ ] **Step 2: 创建 `CalendarSelectionForm.cs` (镜像)**

基于 ContactSelectionForm 复制一份, 修改:
- 类名: `CalendarSelectionForm`
- 命名空间: `MailConverter.Services.Calendars`
- 字段类型: `List<ContactData>` → `List<CalendarData>`
- 字段名: `_contacts` → `_calendars`
- 标题: "选择要导入的联系人" → "选择要导入的日历"
- 列名: "姓名" → "主题", "邮箱" → "开始时间", "电话" → "结束时间"
- DataGridView 列结构调整为日历字段 (Subject, StartTime, EndTime, Location, IsAllDayEvent)
- 属性: `SelectedContactData` → `SelectedCalendarData` (返回 `List<CalendarData>`)

```csharp
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace MailConverter.Services.Calendars
{
    public class CalendarSelectionForm : Form
    {
        private List<CalendarData> _allCalendars;
        private List<CalendarData> _selectedCalendars;
        private CheckBox chkSelectAll;
        private DataGridView dgvCalendars;
        private Button btnOK;
        private Button btnCancel;

        public List<CalendarData> SelectedCalendarData
        {
            get { return _selectedCalendars; }
        }

        public CalendarSelectionForm(List<CalendarData> calendars)
        {
            _allCalendars = calendars ?? new List<CalendarData>();
            _selectedCalendars = new List<CalendarData>();
            InitializeComponent();
            LoadCalendars();
        }

        private void InitializeComponent()
        {
            this.Text = "选择要导入的日历";
            this.Size = new Size(700, 500);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new Size(600, 400);

            chkSelectAll = new CheckBox
            {
                Text = "全选",
                Location = new Point(20, 15),
                AutoSize = true,
                Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold)
            };
            chkSelectAll.CheckedChanged += ChkSelectAll_CheckedChanged;

            dgvCalendars = new DataGridView
            {
                Location = new Point(20, 50),
                Size = new Size(640, 350),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                RowHeadersVisible = false
            };

            var colSelect = new DataGridViewCheckBoxColumn
            {
                Name = "Select",
                HeaderText = "选择",
                Width = 60
            };
            var colSubject = new DataGridViewTextBoxColumn { Name = "Subject", HeaderText = "主题", ReadOnly = true };
            var colStart = new DataGridViewTextBoxColumn { Name = "Start", HeaderText = "开始时间", ReadOnly = true };
            var colEnd = new DataGridViewTextBoxColumn { Name = "End", HeaderText = "结束时间", ReadOnly = true };
            var colLocation = new DataGridViewTextBoxColumn { Name = "Location", HeaderText = "地点", ReadOnly = true };
            var colAllDay = new DataGridViewTextBoxColumn { Name = "AllDay", HeaderText = "全天", ReadOnly = true };

            dgvCalendars.Columns.AddRange(new DataGridViewColumn[] {
                colSelect, colSubject, colStart, colEnd, colLocation, colAllDay
            });
            dgvCalendars.CellValueChanged += DgvCalendars_CellValueChanged;
            dgvCalendars.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (dgvCalendars.IsCurrentCellDirty)
                    dgvCalendars.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };

            btnOK = new Button
            {
                Text = "确定",
                Location = new Point(490, 415),
                Size = new Size(80, 30),
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(580, 415),
                Size = new Size(80, 30),
                DialogResult = DialogResult.Cancel
            };

            this.Controls.Add(chkSelectAll);
            this.Controls.Add(dgvCalendars);
            this.Controls.Add(btnOK);
            this.Controls.Add(btnCancel);
            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void LoadCalendars()
        {
            dgvCalendars.Rows.Clear();
            foreach (var cal in _allCalendars)
            {
                dgvCalendars.Rows.Add(
                    false,
                    cal.Subject ?? "(无主题)",
                    cal.StartTime?.ToString("yyyy-MM-dd HH:mm") ?? "",
                    cal.EndTime?.ToString("yyyy-MM-dd HH:mm") ?? "",
                    cal.Location ?? "",
                    cal.IsAllDayEvent ? "是" : "否"
                );
            }
        }

        private void DgvCalendars_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex != 0) return;
            // 同步 Select 列状态到业务数据 (UI-only, 提交时再筛选)
        }

        private void ChkSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            bool sel = chkSelectAll.Checked;
            foreach (DataGridViewRow row in dgvCalendars.Rows)
            {
                row.Cells["Select"].Value = sel;
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            _selectedCalendars.Clear();
            for (int i = 0; i < dgvCalendars.Rows.Count; i++)
            {
                bool isSelected = Convert.ToBoolean(dgvCalendars.Rows[i].Cells["Select"].Value ?? false);
                if (isSelected && i < _allCalendars.Count)
                {
                    _selectedCalendars.Add(_allCalendars[i]);
                }
            }
        }
    }
}
```

- [ ] **Step 3: 构建验证**

```bash
dotnet build src/MailConverter/MailConverter.csproj -c Debug -v minimal
```

预期: 0 错误 (新文件, 无引用方).

- [ ] **Step 4: 提交**

```bash
git add src/MailConverter/Services/Calendars/CalendarSelectionForm.cs
git commit -m "feat: 新增 CalendarSelectionForm (镜像 ContactSelectionForm)"
```

---

## Task 10: 创建 `SyncContactsControl` (UserControl)

**Files:**
- Create: `src/MailConverter/Services/Contacts/SyncContactsControl.cs`

- [ ] **Step 1: 创建 UserControl**

```csharp
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using MailConverter.Services.Contacts;

namespace MailConverter.Services.Contacts
{
    public class SyncContactsControl : UserControl
    {
        private TextBox txtClientId;
        private TextBox txtTenantId;
        private TextBox txtEmail;
        private Button btnOAuthLogin;
        private Label lblCurrentEmail;
        private ComboBox cmbSourceType;
        private ComboBox cmbCardDavAccounts;
        private Label lblCardDavTip;
        private TextBox txtServerUrl;
        private TextBox txtUsername;
        private TextBox txtPassword;
        private TextBox txtSourceFile;
        private Button btnBrowse;
        private Button btnSelectContacts;
        private Button btnIncrementalSync;
        private Button btnStartSync;
        private ProgressBar progressSync;
        private Label lblSyncStatus;

        private string _accessToken;
        private Office365ImportService _office365Service;
        private List<string> _selectedContactUrls = new List<string>();
        private bool _isO365Connected;

        public string CurrentEmail { get; private set; }
        public bool IsConnected => _isO365Connected;

        public SyncContactsControl()
        {
            Dock = DockStyle.Fill;
            BuildUI();
        }

        private void BuildUI()
        {
            // 标题
            var lblTitle = new Label
            {
                Text = "个人同步联系人到 Office 365",
                Location = new Point(20, 15),
                AutoSize = true,
                Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold)
            };

            // ===== 认证信息 =====
            var lblSectionAuth = new Label
            {
                Text = "1. 认证信息",
                Location = new Point(20, 45),
                AutoSize = true,
                Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 215)
            };

            var lblClientId = new Label { Text = "Client ID:", Location = new Point(30, 75), AutoSize = true };
            txtClientId = new TextBox { Location = new Point(110, 73), Size = new Size(160, 22), Name = "txtSyncContactsClientId" };

            var lblTenantId = new Label { Text = "租户ID:", Location = new Point(280, 75), AutoSize = true };
            txtTenantId = new TextBox { Location = new Point(330, 73), Size = new Size(120, 22), Name = "txtSyncContactsTenantId" };

            var lblEmail = new Label { Text = "邮箱:", Location = new Point(30, 100), AutoSize = true };
            txtEmail = new TextBox { Location = new Point(110, 98), Size = new Size(220, 22), Name = "txtSyncContactsEmail" };

            btnOAuthLogin = new Button
            {
                Text = "登录",
                Location = new Point(350, 96),
                Size = new Size(80, 25),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                Name = "btnSyncContactsLogin"
            };

            lblCurrentEmail = new Label
            {
                Text = "未登录",
                Location = new Point(110, 125),
                AutoSize = true,
                ForeColor = Color.Gray,
                Name = "lblSyncContactsCurrent"
            };

            // ===== 同步配置 =====
            var lblSectionSync = new Label
            {
                Text = "2. 同步配置",
                Location = new Point(20, 155),
                AutoSize = true,
                Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 215)
            };

            var lblSourceType = new Label { Text = "数据来源:", Location = new Point(30, 185), AutoSize = true };
            cmbSourceType = new ComboBox
            {
                Location = new Point(110, 183),
                Size = new Size(140, 22),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Name = "cmbSyncContactsSourceType"
            };
            cmbSourceType.Items.AddRange(new object[] { "本地文件(CSV)", "本地文件(VCF)", "CardDAV", "企业微信API" });
            cmbSourceType.SelectedIndex = 0;
            cmbSourceType.SelectedIndexChanged += CmbSourceType_SelectedIndexChanged;

            var lblCardDavAccount = new Label { Text = "提供商:", Location = new Point(30, 213), AutoSize = true, Visible = false, Name = "lblSyncContactsCardDavAcc" };
            cmbCardDavAccounts = new ComboBox
            {
                Location = new Point(110, 211),
                Size = new Size(180, 22),
                Visible = false,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Name = "cmbSyncContactsCardDav"
            };
            lblCardDavTip = new Label
            {
                Text = "请注意：QQ邮箱需使用授权码",
                Location = new Point(295, 213),
                Size = new Size(200, 20),
                ForeColor = Color.Red,
                Visible = false,
                Name = "lblSyncContactsCardDavTip"
            };

            var lblServerUrl = new Label { Text = "服务器:", Location = new Point(30, 240), AutoSize = true, Name = "lblSyncContactsServer" };
            txtServerUrl = new TextBox { Location = new Point(110, 238), Size = new Size(320, 22), Name = "txtSyncContactsServerUrl" };

            var lblUsername = new Label { Text = "用户名:", Location = new Point(30, 268), AutoSize = true, Name = "lblSyncContactsUsername" };
            txtUsername = new TextBox { Location = new Point(110, 266), Size = new Size(180, 22), Name = "txtSyncContactsUsername" };
            var lblPassword = new Label { Text = "密码:", Location = new Point(310, 268), AutoSize = true, Name = "lblSyncContactsPassword" };
            txtPassword = new TextBox { Location = new Point(350, 266), Size = new Size(120, 22), UseSystemPasswordChar = true, Name = "txtSyncContactsPassword" };

            var lblSourceFile = new Label { Text = "本地文件:", Location = new Point(30, 240), AutoSize = true, Name = "lblSyncContactsSourceFile" };
            txtSourceFile = new TextBox { Location = new Point(110, 238), Size = new Size(270, 22), Name = "txtSyncContactsSourceFile" };
            btnBrowse = new Button
            {
                Text = "浏览...",
                Location = new Point(390, 236),
                Size = new Size(75, 25),
                Name = "btnSyncContactsBrowse"
            };
            btnBrowse.Click += BtnBrowse_Click;

            // ===== 按钮 =====
            btnSelectContacts = new Button
            {
                Text = "选择联系人",
                Location = new Point(110, 305),
                Size = new Size(100, 30),
                BackColor = Color.FromArgb(255, 193, 7),
                ForeColor = Color.Black,
                Visible = false,
                Name = "btnSyncContactsSelect"
            };
            btnIncrementalSync = new Button
            {
                Text = "增量同步",
                Location = new Point(220, 305),
                Size = new Size(100, 30),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                Visible = false,
                Name = "btnSyncContactsIncremental"
            };
            btnStartSync = new Button
            {
                Text = "全量同步",
                Location = new Point(330, 305),
                Size = new Size(100, 30),
                Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                Name = "btnSyncContactsStart"
            };
            btnStartSync.Click += BtnStartSync_Click;

            // ===== 进度 =====
            progressSync = new ProgressBar
            {
                Location = new Point(20, 355),
                Size = new Size(570, 18),
                Style = ProgressBarStyle.Continuous,
                Name = "progressSyncContacts"
            };
            lblSyncStatus = new Label
            {
                Location = new Point(20, 380),
                Size = new Size(570, 18),
                ForeColor = Color.Gray,
                Name = "lblSyncContactsStatus"
            };

            // 添加所有控件
            this.Controls.Add(lblTitle);
            this.Controls.Add(lblSectionAuth);
            this.Controls.Add(lblClientId); this.Controls.Add(txtClientId);
            this.Controls.Add(lblTenantId); this.Controls.Add(txtTenantId);
            this.Controls.Add(lblEmail); this.Controls.Add(txtEmail);
            this.Controls.Add(btnOAuthLogin);
            this.Controls.Add(lblCurrentEmail);
            this.Controls.Add(lblSectionSync);
            this.Controls.Add(lblSourceType); this.Controls.Add(cmbSourceType);
            this.Controls.Add(lblCardDavAccount); this.Controls.Add(cmbCardDavAccounts);
            this.Controls.Add(lblCardDavTip);
            this.Controls.Add(lblServerUrl); this.Controls.Add(txtServerUrl);
            this.Controls.Add(lblUsername); this.Controls.Add(txtUsername);
            this.Controls.Add(lblPassword); this.Controls.Add(txtPassword);
            this.Controls.Add(lblSourceFile); this.Controls.Add(txtSourceFile);
            this.Controls.Add(btnBrowse);
            this.Controls.Add(btnSelectContacts);
            this.Controls.Add(btnIncrementalSync);
            this.Controls.Add(btnStartSync);
            this.Controls.Add(progressSync);
            this.Controls.Add(lblSyncStatus);
        }

        // 事件占位 - Task 11 不实现, Task 12 集成到 MainForm 时由 MainForm 提供
        private void CmbSourceType_SelectedIndexChanged(object sender, EventArgs e) { }
        private void BtnBrowse_Click(object sender, EventArgs e) { }
        private void BtnStartSync_Click(object sender, EventArgs e) { }

        // 由 MainForm 设置 OAuth 登录结果
        public void SetOAuthResult(bool success, string email, string accessToken, Office365ImportService service)
        {
            _isO365Connected = success;
            CurrentEmail = email;
            _accessToken = accessToken;
            _office365Service = service;
            if (success)
            {
                lblCurrentEmail.Text = email;
                lblCurrentEmail.ForeColor = Color.Green;
                lblSyncStatus.Text = "登录成功!";
                lblSyncStatus.ForeColor = Color.Green;
            }
            else
            {
                lblCurrentEmail.Text = "未登录";
                lblCurrentEmail.ForeColor = Color.Gray;
                lblSyncStatus.Text = "登录失败，请重试";
                lblSyncStatus.ForeColor = Color.Red;
            }
        }
    }
}
```

- [ ] **Step 2: 构建验证**

```bash
dotnet build src/MailConverter/MailConverter.csproj -c Debug -v minimal
```

预期: 0 错误 (UserControl 暂时无引用方).

- [ ] **Step 3: 提交**

```bash
git add src/MailConverter/Services/Contacts/SyncContactsControl.cs
git commit -m "feat: 新增 SyncContactsControl (UserControl 占位)"
```

---

## Task 11: 创建 `SyncCalendarControl` (UserControl)

**Files:**
- Create: `src/MailConverter/Services/Calendars/SyncCalendarControl.cs`

- [ ] **Step 1: 复制 `SyncContactsControl` 改为日历版**

```bash
cp src/MailConverter/Services/Contacts/SyncContactsControl.cs src/MailConverter/Services/Calendars/SyncCalendarControl.cs
```

修改:
- 命名空间: `MailConverter.Services.Contacts` → `MailConverter.Services.Calendars`
- 类名: `SyncContactsControl` → `SyncCalendarControl`
- 标题: "个人同步联系人到 Office 365" → "个人同步日历到 Office 365"
- 数据来源选项: `"本地文件(CSV)", "本地文件(VCF)", "CardDAV", "企业微信API"` → `"本地文件(ICS)", "本地文件(VCS)", "CalDAV"`
- 移除 CardDAV 账户选择 + 选择联系人按钮 + 增量同步按钮 (日历不需)
- 按钮名称: `btnSyncContacts*` → `btnSyncCalendar*`

- [ ] **Step 2: 构建验证**

```bash
dotnet build src/MailConverter/MailConverter.csproj -c Debug -v minimal
```

预期: 0 错误.

- [ ] **Step 3: 提交**

```bash
git add src/MailConverter/Services/Calendars/SyncCalendarControl.cs
git commit -m "feat: 新增 SyncCalendarControl (UserControl 占位)"
```

---

## Task 12: MainForm 集成两个 Tab

**Files:**
- Modify: `src/MailConverter/MainForm.cs` (Tab 创建 + Panel 构造)

- [ ] **Step 1: 修改 Tab 创建逻辑 (line 1374-1386 区域)**

原代码:
```csharp
_syncTab = new TabPage("同步联系人/日历");
var syncPanel = CreateSyncContactsCalendarPanel();
_syncTab.Controls.Add(syncPanel);
...
_o365NestedTabControl.TabPages.Add(_syncTab);
```

替换为:
```csharp
// 同步联系人 Tab
var syncContactsTab = new TabPage("同步联系人");
var syncContactsControl = new MailConverter.Services.Contacts.SyncContactsControl();
syncContactsTab.Controls.Add(syncContactsControl);
_o365NestedTabControl.TabPages.Add(syncContactsTab);

// 同步日历 Tab
var syncCalendarTab = new TabPage("同步日历");
var syncCalendarControl = new MailConverter.Services.Calendars.SyncCalendarControl();
syncCalendarTab.Controls.Add(syncCalendarControl);
_o365NestedTabControl.TabPages.Add(syncCalendarTab);
```

- [ ] **Step 2: 删除 `CreateSyncContactsCalendarPanel` 方法**

删除 MainForm.cs 11725-12300 行的 `CreateSyncContactsCalendarPanel` 整段 (约 575 行).

- [ ] **Step 3: 更新 Feature_SingleUserSync_Contacts 引用**

line 1386:
```csharp
if (!_featureSettings.Feature_SingleUserSync_Contacts) _syncTab.SetVisible(false);
```

改为分别处理两个 Tab (如果需要, 但因为 FeatureSettingsService 是单联系人开关, 日历可以另开一个 flag 或复用同一个). 最小改动: 保留 `Feature_SingleUserSync_Contacts` 控制两个 Tab 同时显示/隐藏.

- [ ] **Step 4: 构建验证**

```bash
dotnet build src/MailConverter/MailConverter.csproj -c Debug -v minimal
```

预期: 0 错误.

- [ ] **Step 5: 启动验证**

```bash
tasklist | grep -i mailconverter
# 如果在运行, 杀掉
taskkill //F //IM MailConverter.exe
./src/MailConverter/bin/Debug/net48/MailConverter.exe &
sleep 3
tasklist | grep -i mailconverter
```

预期: MailConverter 进程存在.

- [ ] **Step 6: 提交**

```bash
git add src/MailConverter/MainForm.cs
git commit -m "refactor: 单用户同步 Tab 拆为同步联系人 + 同步日历两个独立 Tab"
```

---

## Task 13: 完整端到端验证

- [ ] **Step 1: 启动应用**

```bash
tasklist | grep -i mailconverter
# 如果在运行, 杀掉
taskkill //F //IM MailConverter.exe
./src/MailConverter/bin/Debug/net48/MailConverter.exe &
sleep 3
```

- [ ] **Step 2: UI 检查**

- 点击 "单用户同步" 边栏按钮
- 确认嵌套 Tab 显示: EML导入 / PST导入 / 同步联系人 / 同步日历
- 点击 "同步联系人" Tab: 看到认证表单 + 数据源选择 (CSV/VCF/CardDAV/企业微信)
- 点击 "同步日历" Tab: 看到认证表单 + 数据源选择 (ICS/VCS/CalDAV)
- 两个 Tab 的登录状态独立

- [ ] **Step 3: PST 批量同步回归**

- 点击 "PST批量同步" 边栏按钮
- 进入 "PST同步联系人" Tab: 选 PST 文件, 执行同步, 验证仍能导入
- 进入 "PST同步日历" Tab: 同样验证

- [ ] **Step 4: 关闭应用 + 最终提交 (如有改动)**

```bash
taskkill //F //IM MailConverter.exe
git status
# 如果有改动, commit
```

---

## 备注

- **未实现功能**: Task 10/11 中的 UserControl 事件处理 (`CmbSourceType_SelectedIndexChanged`, `BtnBrowse_Click`, `BtnStartSync_Click`) 是占位. Task 12 集成时, MainForm 可以选择:
  - A) 完整实现这些事件 (复制原 `CreateSyncContactsCalendarPanel` 内的逻辑)
  - B) 暂时只验证 UI 框架, 事件逻辑后续补
- **建议**: 选 A. 完整功能保持不变, 用户体验连续. 实施时把 `CreateSyncContactsCalendarPanel` 内的 `btnStartSync_Click` 等事件处理逻辑拆分到两个 Control 内, 联系人 Control 处理联系人分支, 日历 Control 处理日历分支.
- **PstContactData 引用**: 全部 6 处引用在 MainForm.cs (14930, 14965, 14984, 15886, 15892, 15900, 16058, 16072) 都要更新. 用 Edit tool 一个个改, 每处都需要 Read 周围上下文.
