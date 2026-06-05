# 单用户同步（O365）界面重构实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将单用户同步（O365）页面的 EML 导入和 PST 导入重新设计为卡片式布局，联系人/日历同步统一样式

**Architecture:** 使用分组卡片式布局，通过自定义 Panel + Paint 事件实现阴影效果，保留现有字段引用和事件处理逻辑

**Tech Stack:** .NET Framework 4.8 WinForms, C#

---

## 文件结构

```
src/MailConverter/MainForm.cs
├── CreateCardPanel(title)          [新建 - 辅助方法]
├── StyleTextBox(txt)               [新建 - 统一样式]
├── StyleComboBox(cmb)              [新建 - 统样式]
├── StylePrimaryButton(btn)         [新建 - 统一样式]
├── StyleSecondaryButton(btn)        [新建 - 统一样式]
├── CreateEmImportPanel()           [新建 - EML导入]
├── CreatePstImportPanelV2()        [新建 - PST导入]
└── (现有联系人/日历方法)           [仅添加样式调用]
```

---

## Task 1: 创建样式辅助方法

**Files:**
- Modify: `src/MailConverter/MainForm.cs` (在现有方法区域添加)

- [ ] **Step 1: 添加 CreateCardPanel 辅助方法**

```csharp
/// <summary>
/// 创建带阴影效果的卡片面板
/// </summary>
private Panel CreateCardPanel(string title, int height)
{
    var card = new Panel
    {
        BackColor = Color.White,
        Size = new Size(550, height),
        Padding = new Padding(15)
    };

    // 标题
    var lblTitle = new Label
    {
        Text = title,
        Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold),
        ForeColor = Color.FromArgb(50, 50, 50),
        Location = new Point(15, 10),
        AutoSize = true
    };

    // 内容区域（标题下方）
    var contentPanel = new Panel
    {
        Location = new Point(15, 35),
        Size = new Size(card.Width - 30, height - 50)
    };

    card.Controls.Add(lblTitle);
    card.Controls.Add(contentPanel);

    // 阴影效果
    card.Paint += (s, e) =>
    {
        using (var shadowPen = new Pen(Color.FromArgb(180, 180, 180), 1))
        using (var whitePen = new Pen(Color.White, 1))
        {
            // 外层阴影边框
            e.Graphics.DrawRectangle(shadowPen, 0, 0, card.Width - 1, card.Height - 1);
            // 内层白色边框
            e.Graphics.DrawRectangle(whitePen, 0, 0, card.Width - 2, card.Height - 2);
        }
    };

    return card;
}
```

- [ ] **Step 2: 添加 StyleTextBox 辅助方法**

```csharp
/// <summary>
/// 统一样式文本框
/// </summary>
private void StyleTextBox(TextBox txt)
{
    txt.Height = 26;
    txt.BorderStyle = BorderStyle.FixedSingle;
    txt.Font = new Font("Microsoft Sans Serif", 9F);
}
```

- [ ] **Step 3: 添加 StyleComboBox 辅助方法**

```csharp
/// <summary>
/// 统一样式下拉框
/// </summary>
private void StyleComboBox(ComboBox cmb)
{
    cmb.Height = 26;
    cmb.DropDownStyle = ComboBoxStyle.DropDownList;
    cmb.Font = new Font("Microsoft Sans Serif", 9F);
}
```

- [ ] **Step 4: 添加 StylePrimaryButton 辅助方法**

```csharp
/// <summary>
/// 统一样式主按钮
/// </summary>
private void StylePrimaryButton(Button btn)
{
    btn.Height = 36;
    btn.FlatStyle = FlatStyle.Flat;
    btn.BackColor = Color.FromArgb(0, 120, 215);
    btn.ForeColor = Color.White;
    btn.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold);
    btn.Cursor = Cursors.Hand;
}
```

- [ ] **Step 5: 添加 StyleSecondaryButton 辅助方法**

```csharp
/// <summary>
/// 统一样式次按钮
/// </summary>
private void StyleSecondaryButton(Button btn)
{
    btn.Height = 26;
    btn.FlatStyle = FlatStyle.Flat;
    btn.BackColor = Color.FromArgb(245, 245, 245);
    btn.ForeColor = Color.FromArgb(50, 50, 50);
    btn.Font = new Font("Microsoft Sans Serif", 9F);
    btn.Cursor = Cursors.Hand;
}
```

- [ ] **Step 6: 编译验证**

Run: `dotnet build src/MailConverter/MailConverter.csproj`
Expected: 编译成功（仅有原有警告）

---

## Task 2: 创建 EML 导入页面 CreateEmImportPanel()

**Files:**
- Modify: `src/MailConverter/MainForm.cs` (在 CreateOffice365Panel 方法后添加)

- [ ] **Step 1: 添加 CreateEmImportPanel 方法框架**

```csharp
/// <summary>
/// EML 邮件导入页面（卡片式布局）
/// </summary>
private Panel CreateEmImportPanel()
{
    var mainPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(240, 240, 240) };

    var container = new Panel
    {
        Location = new Point(20, 20),
        Size = new Size(580, 580)
    };
    mainPanel.Controls.Add(container);

    int cardY = 0;
    const int cardHeight = 100;
    const int cardGap = 15;

    // 卡片1: 账户配置
    var cardAccount = CreateCardPanel("账户配置", 100);
    cardAccount.Location = new Point(0, cardY);
    SetupAccountCard(cardAccount);
    container.Controls.Add(cardAccount);
    cardY += cardHeight + cardGap;

    // 卡片2: 源文件配置
    var cardSource = CreateCardPanel("源文件配置", 80);
    cardSource.Location = new Point(0, cardY);
    SetupSourceCard(cardSource, false);
    container.Controls.Add(cardSource);
    cardY += cardSource.Height + cardGap;

    // 卡片3: 目标配置
    var cardTarget = CreateCardPanel("目标配置", 70);
    cardTarget.Location = new Point(0, cardY);
    SetupTargetCard(cardTarget);
    container.Controls.Add(cardTarget);
    cardY += cardTarget.Height + cardGap;

    // 导入按钮
    var btnImport = new Button
    {
        Text = "▶ 开始导入",
        Location = new Point(200, cardY),
        Size = new Size(120, 36)
    };
    StylePrimaryButton(btnImport);
    btnImport.Click += (s, e) => BtnO365Import_Click(cmbO365TargetFolder?.Text ?? "Inbox");
    container.Controls.Add(btnImport);
    cardY += btnImport.Height + cardGap;

    // 卡片4: 导入进度
    var cardProgress = CreateCardPanel("导入进度", 80);
    cardProgress.Location = new Point(0, cardY);
    SetupProgressCard(cardProgress);
    container.Controls.Add(cardProgress);

    return mainPanel;
}
```

- [ ] **Step 2: 添加 SetupAccountCard 方法**

```csharp
private void SetupAccountCard(Panel card)
{
    var content = card.Controls[1] as Panel; // 内容面板

    // 账户选择
    var lblAccount = new Label { Text = "目标账户:", Location = new Point(0, 5), AutoSize = true };
    cmbO365SavedAccounts = new ComboBox
    {
        Location = new Point(0, 25),
        Size = new Size(200, 26),
        DropDownStyle = ComboBoxStyle.DropDownList
    };
    cmbO365SavedAccounts.Items.Add("-- 选择账户 --");
    cmbO365SavedAccounts.SelectedIndex = 0;
    cmbO365SavedAccounts.SelectedIndexChanged += CmbO365SavedAccounts_SelectedIndexChanged;
    StyleComboBox(cmbO365SavedAccounts);

    var btnOtherAccount = new Button { Text = "使用其他账户", Location = new Point(210, 24), Size = new Size(100, 26) };
    StyleSecondaryButton(btnOtherAccount);

    // Client ID
    var lblClientId = new Label { Text = "Client ID:", Location = new Point(0, 55), AutoSize = true };
    txtO365ClientId = new TextBox { Location = new Point(75, 53), Size = new Size(180, 26) };
    StyleTextBox(txtO365ClientId);

    // 租户ID
    var lblTenantId = new Label { Text = "租户ID:", Location = new Point(270, 55), AutoSize = true };
    txtO365TenantId = new TextBox { Location = new Point(330, 53), Size = new Size(120, 26) };
    StyleTextBox(txtO365TenantId);

    // 登录按钮
    btnO365OAuthLogin = new Button
    {
        Text = "🔐 使用 Microsoft 登录",
        Location = new Point(460, 24),
        Size = new Size(140, 55)
    };
    StylePrimaryButton(btnO365OAuthLogin);
    btnO365OAuthLogin.Click += BtnO365OAuthLogin_Click;

    content.Controls.AddRange(new Control[] {
        lblAccount, cmbO365SavedAccounts, btnOtherAccount,
        lblClientId, txtO365ClientId, lblTenantId, txtO365TenantId,
        btnO365OAuthLogin
    });

    // 加载已保存的账户
    LoadSavedOAuthAccounts();
}
```

- [ ] **Step 3: 添加 SetupSourceCard 方法**

```csharp
private void SetupSourceCard(Panel card, bool isPstMode)
{
    var content = card.Controls[1] as Panel;

    var lblSourceType = new Label { Text = "邮件来源:", Location = new Point(0, 5), AutoSize = true };
    cmbO365SourceType = new ComboBox
    {
        Location = new Point(75, 3),
        Size = new Size(150, 26),
        Items = { "EML 文件夹", "PST 文件" },
        SelectedIndex = isPstMode ? 1 : 0
    };
    StyleComboBox(cmbO365SourceType);
    cmbO365SourceType.SelectedIndexChanged += CmbO365SourceType_SelectedIndexChanged;

    var lblPath = new Label { Text = "路径:", Location = new Point(0, 38), AutoSize = true };
    txtO365SourcePath = new TextBox { Location = new Point(40, 36), Size = new Size(360, 26) };
    StyleTextBox(txtO365SourcePath);

    btnO365SourceBrowse = new Button { Text = "浏览...", Location = new Point(410, 34), Size = new Size(80, 26) };
    StyleSecondaryButton(btnO365SourceBrowse);
    btnO365SourceBrowse.Click += BtnO365SourceBrowse_Click;

    content.Controls.AddRange(new Control[] {
        lblSourceType, cmbO365SourceType, lblPath, txtO365SourcePath, btnO365SourceBrowse
    });
}
```

- [ ] **Step 4: 添加 SetupTargetCard 方法**

```csharp
private void SetupTargetCard(Panel card)
{
    var content = card.Controls[1] as Panel;

    var lblFolder = new Label { Text = "目标文件夹:", Location = new Point(0, 5), AutoSize = true };
    cmbO365TargetFolder = new ComboBox
    {
        Location = new Point(85, 3),
        Size = new Size(150, 26),
        DropDownStyle = ComboBoxStyle.DropDown,
        Items = { "Inbox", "Sent Items", "Drafts", "Deleted Items" }
    };
    cmbO365TargetFolder.Text = "Inbox";
    StyleComboBox(cmbO365TargetFolder);

    content.Controls.AddRange(new Control[] { lblFolder, cmbO365TargetFolder });
}
```

- [ ] **Step 5: 添加 SetupProgressCard 方法**

```csharp
private void SetupProgressCard(Panel card)
{
    var content = card.Controls[1] as Panel;

    progressO365 = new ProgressBar
    {
        Location = new Point(0, 5),
        Size = new Size(content.Width, 6),
        Style = ProgressBarStyle.Continuous
    };

    lblO365Status = new Label
    {
        Location = new Point(0, 20),
        Size = new Size(content.Width, 20),
        ForeColor = Color.Gray,
        Text = "就绪"
    };

    content.Controls.AddRange(new Control[] { progressO365, lblO365Status });
}
```

- [ ] **Step 6: 编译验证**

Run: `dotnet build src/MailConverter/MailConverter.csproj`
Expected: 编译成功

---

## Task 3: 创建 PST 导入页面 CreatePstImportPanelV2()

**Files:**
- Modify: `src/MailConverter/MainForm.cs`

- [ ] **Step 1: 添加 CreatePstImportPanelV2 方法框架**

```csharp
/// <summary>
/// PST 文件导入页面（卡片式布局）
/// </summary>
private Panel CreatePstImportPanelV2()
{
    var mainPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(240, 240, 240) };

    var container = new Panel
    {
        Location = new Point(20, 20),
        Size = new Size(580, 620)
    };
    mainPanel.Controls.Add(container);

    int cardY = 0;
    const int cardGap = 15;

    // 卡片1: 账户配置
    var cardAccount = CreateCardPanel("账户配置", 100);
    cardAccount.Location = new Point(0, cardY);
    SetupAccountCard(cardAccount);
    container.Controls.Add(cardAccount);
    cardY += 115;

    // 卡片2: PST 文件配置
    var cardPst = CreateCardPanel("PST 文件配置", 110);
    cardPst.Location = new Point(0, cardY);
    SetupPstCard(cardPst);
    container.Controls.Add(cardPst);
    cardY += cardPst.Height + cardGap;

    // 卡片3: 目标配置
    var cardTarget = CreateCardPanel("目标配置", 80);
    cardTarget.Location = new Point(0, cardY);
    SetupPstTargetCard(cardTarget);
    container.Controls.Add(cardTarget);
    cardY += cardTarget.Height + cardGap;

    // 导入按钮
    var btnImport = new Button
    {
        Text = "▶ 开始导入",
        Location = new Point(200, cardY),
        Size = new Size(120, 36)
    };
    StylePrimaryButton(btnImport);
    btnImport.Click += (s, e) => BtnO365Import_Click(cmbO365TargetFolder?.Text ?? "Inbox");
    container.Controls.Add(btnImport);
    cardY += btnImport.Height + cardGap;

    // 卡片4: 导入进度
    var cardProgress = CreateCardPanel("导入进度", 80);
    cardProgress.Location = new Point(0, cardY);
    SetupProgressCard(cardProgress);
    container.Controls.Add(cardProgress);

    return mainPanel;
}
```

- [ ] **Step 2: 添加 SetupPstCard 方法**

```csharp
private void SetupPstCard(Panel card)
{
    var content = card.Controls[1] as Panel;

    // 导入方式
    var lblMethod = new Label { Text = "导入方式:", Location = new Point(0, 5), AutoSize = true };
    radO365PstExtract = new RadioButton
    {
        Text = "提取为 EML 后导入",
        Location = new Point(80, 5),
        AutoSize = true,
        Checked = true
    };
    radO365PstDirect = new RadioButton
    {
        Text = "EWS 直接上传",
        Location = new Point(220, 5),
        AutoSize = true
    };

    // 文件选择
    var lblPath = new Label { Text = "文件:", Location = new Point(0, 35), AutoSize = true };
    txtO365SourcePath = new TextBox { Location = new Point(40, 33), Size = new Size(360, 26) };
    StyleTextBox(txtO365SourcePath);

    btnO365SourceBrowse = new Button { Text = "浏览...", Location = new Point(410, 31), Size = new Size(80, 26) };
    StyleSecondaryButton(btnO365SourceBrowse);
    btnO365SourceBrowse.Click += BtnO365SourceBrowse_Click;

    // 提示
    var lblNote = new Label
    {
        Text = "⚠ 注意: EWS 直接上传需要 Exchange Server 支持",
        Location = new Point(0, 65),
        Size = new Size(400, 18),
        ForeColor = Color.OrangeRed,
        Font = new Font("Microsoft Sans Serif", 8F)
    };

    content.Controls.AddRange(new Control[] {
        lblMethod, radO365PstExtract, radO365PstDirect,
        lblPath, txtO365SourcePath, btnO365SourceBrowse, lblNote
    });
}
```

- [ ] **Step 3: 添加 SetupPstTargetCard 方法**

```csharp
private void SetupPstTargetCard(Panel card)
{
    var content = card.Controls[1] as Panel;

    var lblFolder = new Label { Text = "目标文件夹:", Location = new Point(0, 5), AutoSize = true };
    cmbO365TargetFolder = new ComboBox
    {
        Location = new Point(85, 3),
        Size = new Size(150, 26),
        DropDownStyle = ComboBoxStyle.DropDown,
        Items = { "Inbox", "Sent Items", "Drafts", "Deleted Items" }
    };
    cmbO365TargetFolder.Text = "Inbox";
    StyleComboBox(cmbO365TargetFolder);

    var lblThread = new Label { Text = "并发线程:", Location = new Point(260, 5), AutoSize = true };
    var numThread = new NumericUpDown
    {
        Location = new Point(335, 3),
        Size = new Size(50, 26),
        Minimum = 1,
        Maximum = 10,
        Value = 2
    };

    content.Controls.AddRange(new Control[] {
        lblFolder, cmbO365TargetFolder, lblThread, numThread
    });
}
```

- [ ] **Step 4: 编译验证**

Run: `dotnet build src/MailConverter/MailConverter.csproj`
Expected: 编译成功

---

## Task 4: 调整 Tab 创建逻辑

**Files:**
- Modify: `src/MailConverter/MainForm.cs` (约第1345-1360行)

- [ ] **Step 1: 定位并修改 Tab 创建代码**

找到 `_emlImportTab` 和 `_pstImportTab` 的创建代码，修改为：

```csharp
// EML 导入 Tab
_emlImportTab = new TabPage("EML 导入");
var emlImportPanel = CreateEmImportPanel();
emlImportPanel.Dock = DockStyle.Fill;
_emlImportTab.Controls.Add(emlImportPanel);

// PST 导入 Tab
_pstImportTab = new TabPage("PST 导入");
var pstImportPanel = CreatePstImportPanelV2();
pstImportPanel.Dock = DockStyle.Fill;
_pstImportTab.Controls.Add(pstImportPanel);
```

- [ ] **Step 2: 删除旧的 CreateOffice365Panel 调用**

删除 `CreateOffice365Panel()` 的调用，改用新的 `CreateEmImportPanel()` 和 `CreatePstImportPanelV2()`

- [ ] **Step 3: 编译验证**

Run: `dotnet build src/MailConverter/MailConverter.csproj`
Expected: 编译成功

---

## Task 5: 联系人/日历同步统一样式（可选，仅视觉统一）

**Files:**
- Modify: `src/MailConverter/MainForm.cs` (Contact/Calendar 相关方法)

- [ ] **Step 1: 检查现有方法结构**

这些方法保持不变，仅在需要时添加样式辅助方法调用

- [ ] **Step 2: 如需统一样式，添加样式调用**

在现有 `CreateSyncContactsCalendarPanel()` 等方法的末尾添加卡片容器样式

---

## Task 6: 功能验证

- [ ] **Step 1: 启动应用**

Run: `dotnet run --project src/MailConverter/MailConverter.csproj` 或直接运行 exe

- [ ] **Step 2: 检查 EML 导入页面**

进入"单用户同步（O365）" → "EML 导入" Tab，验证：
- 卡片式布局显示正常
- 账户配置、源文件配置、目标配置、导入进度四个卡片
- 按钮样式正确

- [ ] **Step 3: 检查 PST 导入页面**

进入"PST 导入" Tab，验证：
- 卡片式布局显示正常
- 账户配置、PST文件配置、目标配置、导入进度四个卡片
- 导入方式单选框可用

- [ ] **Step 4: 测试基本功能**

1. 选择账户或输入 Client ID/Tenant ID
2. 点击"使用 Microsoft 登录"
3. 选择 EML 文件夹路径
4. 点击"开始导入"验证流程

---

## 字段引用对照表

| 字段 | 类型 | 用途 | 保留位置 |
|------|------|------|----------|
| `cmbO365SavedAccounts` | ComboBox | 账户选择 | SetupAccountCard |
| `txtO365ClientId` | TextBox | Client ID | SetupAccountCard |
| `txtO365TenantId` | TextBox | 租户ID | SetupAccountCard |
| `btnO365OAuthLogin` | Button | 登录按钮 | SetupAccountCard |
| `cmbO365SourceType` | ComboBox | 邮件来源 | SetupSourceCard |
| `txtO365SourcePath` | TextBox | 文件路径 | SetupSourceCard/SetupPstCard |
| `btnO365SourceBrowse` | Button | 浏览按钮 | SetupSourceCard/SetupPstCard |
| `cmbO365TargetFolder` | ComboBox | 目标文件夹 | SetupTargetCard/SetupPstTargetCard |
| `btnO365Import` | Button | 导入按钮 | CreateEmImportPanel/CreatePstImportPanelV2 |
| `progressO365` | ProgressBar | 进度条 | SetupProgressCard |
| `lblO365Status` | Label | 状态文字 | SetupProgressCard |
| `radO365PstExtract` | RadioButton | PST提取方式 | SetupPstCard |
| `radO365PstDirect` | RadioButton | PST EWS方式 | SetupPstCard |

## 事件处理保留

| 事件 | 处理方法 | 说明 |
|------|----------|------|
| `cmbO365SavedAccounts.SelectedIndexChanged` | `CmbO365SavedAccounts_SelectedIndexChanged` | 账户切换 |
| `btnO365OAuthLogin.Click` | `BtnO365OAuthLogin_Click` | OAuth登录 |
| `cmbO365SourceType.SelectedIndexChanged` | `CmbO365SourceType_SelectedIndexChanged` | 来源类型切换 |
| `btnO365SourceBrowse.Click` | `BtnO365SourceBrowse_Click` | 浏览文件 |
| `btnO365Import.Click` | `BtnO365Import_Click` | 开始导入 |

---

## 实施检查清单

- [ ] Task 1: 样式辅助方法已添加
- [ ] Task 2: EML 导入页面已创建
- [ ] Task 3: PST 导入页面已创建
- [ ] Task 4: Tab 逻辑已调整
- [ ] Task 5: 联系人/日历样式（可选）
- [ ] Task 6: 功能验证通过