using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using MailConverter.Services.Calendars;

namespace MailConverter.Services.Calendars
{
    public class SyncCalendarControl : UserControl
    {
        private TextBox txtClientId;
        private TextBox txtTenantId;
        private TextBox txtEmail;
        private Button btnOAuthLogin;
        private Label lblCurrentEmail;
        private ComboBox cmbSourceType;
        private TextBox txtServerUrl;
        private TextBox txtUsername;
        private TextBox txtPassword;
        private TextBox txtSourceFile;
        private Button btnBrowse;
        private Button btnStartSync;
        private ProgressBar progressSync;
        private Label lblSyncStatus;
        private Label lblProgressPercent;
        private Label lblServerUrl;
        private Label lblUsername;
        private Label lblPassword;
        private Label lblSourceFile;
        private ComboBox cmbSyncAccounts;
        private ComboBox cmbCalDavAccounts;
        private Label lblCalDavAccount;
        // 两个独立的时区下拉框(显示同一选中值):
        //   cmbSourceTimeZone - "文件源时区" 在数据来源行
        //   cmbImportTimeZone - "导入时区" 在本地文件行
        private ComboBox cmbSourceTimeZone;
        private Label lblSourceTimeZone;
        private ComboBox cmbImportTimeZone;
        private Label lblImportTimeZone;
        private TableLayoutPanel tblSource;

        /// <summary>
        /// 由 MainForm 在创建时注入,用于调用 BtnO365OAuthLogin_Click 等方法
        /// (命名避开 UserControl 自身的只读 ParentForm 属性)
        /// </summary>
        public MainForm MainForm { get; set; }

        public SyncCalendarControl()
        {
            Dock = DockStyle.Fill;
            BuildUI();
        }

        private void BuildUI()
        {
            // 外部容器 - 仿 PST/EML 面板的 3 行 TableLayoutPanel
            var mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(15),
                RowStyles = {
                    new RowStyle(SizeType.Absolute, 180),  // 账户配置 (4 行)
                    new RowStyle(SizeType.Percent, 100),   // 同步源配置
                    new RowStyle(SizeType.Absolute, 100)   // 操作区
                }
            };

            // ========== 区块 1: 账户配置 (GroupBox) ==========
            var grpAccount = new GroupBox
            {
                Text = "账户配置",
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };

            var tblAccount = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 4,
                Padding = new Padding(5),
                RowStyles = {
                    new RowStyle(SizeType.Absolute, 30),  // Row 0: 已保存账户
                    new RowStyle(SizeType.Absolute, 30),  // Row 1: Client ID + 租户ID
                    new RowStyle(SizeType.Absolute, 30),  // Row 2: 邮箱
                    new RowStyle(SizeType.Absolute, 40)   // Row 3: 登录按钮 (独占一行)
                },
                ColumnStyles = {
                    new ColumnStyle(SizeType.AutoSize),
                    new ColumnStyle(SizeType.Percent, 50),
                    new ColumnStyle(SizeType.AutoSize),
                    new ColumnStyle(SizeType.Percent, 50)
                }
            };

            var lblSavedAccount = new Label { Text = "已保存账户:", AutoSize = true };
            cmbSyncAccounts = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Dock = DockStyle.Fill
            };
            cmbSyncAccounts.SelectedIndexChanged += CmbSyncAccounts_SelectedIndexChanged;

            var lblClientId = new Label { Text = "Client ID:", AutoSize = true };
            txtClientId = new TextBox { Dock = DockStyle.Fill };
            var lblTenantId = new Label { Text = "租户ID:", AutoSize = true };
            txtTenantId = new TextBox { Dock = DockStyle.Fill };

            var lblEmail = new Label { Text = "邮箱:", AutoSize = true };
            txtEmail = new TextBox { Dock = DockStyle.Fill };

            btnOAuthLogin = new Button
            {
                Text = "▶ Microsoft 登录",
                AutoSize = false,
                Size = new Size(160, 32),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold)
            };
            btnOAuthLogin.Click += BtnOAuthLogin_Click;

            // lblCurrentEmail 字段保留但隐藏 (供 BtnOAuthLogin_Click 等处理器引用)
            lblCurrentEmail = new Label
            {
                Text = "未登录",
                AutoSize = true,
                ForeColor = Color.Gray,
                Visible = false
            };

            // Row 0: 已保存账户 (跨 3 列)
            tblAccount.Controls.Add(lblSavedAccount, 0, 0);
            tblAccount.SetColumnSpan(cmbSyncAccounts, 3);
            tblAccount.Controls.Add(cmbSyncAccounts, 1, 0);

            // Row 1: Client ID + 租户ID
            tblAccount.Controls.Add(lblClientId, 0, 1);
            tblAccount.Controls.Add(txtClientId, 1, 1);
            tblAccount.Controls.Add(lblTenantId, 2, 1);
            tblAccount.Controls.Add(txtTenantId, 3, 1);

            // Row 2: 邮箱 (跨 3 列)
            tblAccount.Controls.Add(lblEmail, 0, 2);
            tblAccount.SetColumnSpan(txtEmail, 3);
            tblAccount.Controls.Add(txtEmail, 1, 2);

            // Row 3: 登录按钮 (右对齐,独占一行)
            tblAccount.Controls.Add(btnOAuthLogin, 3, 3);
            btnOAuthLogin.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            btnOAuthLogin.Dock = DockStyle.None;

            grpAccount.Controls.Add(tblAccount);

            // ========== 区块 2: 同步源配置 (GroupBox) ==========
            var grpSource = new GroupBox
            {
                Text = "同步源配置",
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };

            tblSource = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 5,
                RowCount = 6,
                Padding = new Padding(5),
                RowStyles = {
                    new RowStyle(SizeType.Absolute, 30),  // Row 0: 数据来源 + 文件源时区
                    new RowStyle(SizeType.Absolute, 30),  // Row 1: 本地文件 + 导入时区
                    new RowStyle(SizeType.Absolute, 32),  // Row 2: 浏览按钮 (only 本地文件)
                    new RowStyle(SizeType.Absolute, 30),  // Row 3: CalDAV/CardDAV 账户
                    new RowStyle(SizeType.Absolute, 30),  // Row 4: 服务器
                    new RowStyle(SizeType.Absolute, 36)   // Row 5: 用户名 + 密码
                },
                ColumnStyles = {
                    new ColumnStyle(SizeType.AutoSize),     // 0: label
                    new ColumnStyle(SizeType.Percent, 30),  // 1: control/combo
                    new ColumnStyle(SizeType.AutoSize),     // 2: label
                    new ColumnStyle(SizeType.Percent, 30),  // 3: control/combo
                    new ColumnStyle(SizeType.AutoSize)      // 4: button
                }
            };

            var lblSourceType = new Label { Text = "数据来源:", AutoSize = true };
            cmbSourceType = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Dock = DockStyle.Fill
            };
            cmbSourceType.Items.AddRange(new object[] { "CSV", "ICS", "CalDAV" });
            cmbSourceType.SelectedIndex = 0;
            cmbSourceType.SelectedIndexChanged += CmbSourceType_SelectedIndexChanged;

            // 文件源时区 (Row 0, 紧跟数据来源)
            lblSourceTimeZone = new Label { Text = "文件源时区:", AutoSize = true };
            cmbSourceTimeZone = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Dock = DockStyle.Fill
            };
            PopulateTimeZoneCombo(cmbSourceTimeZone, defaultValue: "UTC+8");
            cmbSourceTimeZone.SelectedIndexChanged += CmbSourceTimeZone_SelectedIndexChanged;

            // 本地文件相关控件
            lblSourceFile = new Label { Text = "本地文件:", AutoSize = true };
            txtSourceFile = new TextBox { Dock = DockStyle.Fill };
            btnBrowse = new Button
            {
                Text = "浏览...",
                Width = 80,
                BackColor = Color.FromArgb(245, 245, 245),
                FlatStyle = FlatStyle.Flat
            };
            btnBrowse.Click += BtnBrowse_Click;

            // 导入时区 (Row 1, 紧跟本地文件)
            lblImportTimeZone = new Label { Text = "导入时区:", AutoSize = true, Visible = false };
            cmbImportTimeZone = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Dock = DockStyle.Fill,
                Visible = false
            };
            PopulateTimeZoneCombo(cmbImportTimeZone, defaultValue: "UTC+8");
            cmbImportTimeZone.SelectedIndexChanged += CmbImportTimeZone_SelectedIndexChanged;

            // CalDAV 相关控件
            lblCalDavAccount = new Label { Text = "CalDAV账户:", AutoSize = true, Visible = false };
            cmbCalDavAccounts = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Dock = DockStyle.Fill,
                Visible = false
            };
            cmbCalDavAccounts.SelectedIndexChanged += CmbCalDavAccounts_SelectedIndexChanged;

            lblServerUrl = new Label { Text = "服务器:", AutoSize = true };
            txtServerUrl = new TextBox { Dock = DockStyle.Fill };

            lblUsername = new Label { Text = "用户名:", AutoSize = true };
            txtUsername = new TextBox { Dock = DockStyle.Fill };
            lblPassword = new Label { Text = "密码:", AutoSize = true };
            txtPassword = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = true };

            // Row 0: 数据来源 (col 0-1) + 文件源时区 (col 2-3)
            tblSource.Controls.Add(lblSourceType, 0, 0);
            tblSource.Controls.Add(cmbSourceType, 1, 0);
            tblSource.Controls.Add(lblSourceTimeZone, 2, 0);
            tblSource.Controls.Add(cmbSourceTimeZone, 3, 0);

            // Row 1: 本地文件 (col 0-1) + 导入时区 label (col 2) + 导入时区 cmb (col 3)
            // col 3 与 Row 0 的"文件源时区"下拉框同列,实现两个时区下拉框纵向对齐
            tblSource.Controls.Add(lblSourceFile, 0, 1);
            tblSource.Controls.Add(txtSourceFile, 1, 1);
            tblSource.Controls.Add(lblImportTimeZone, 2, 1);
            tblSource.Controls.Add(cmbImportTimeZone, 3, 1);

            // Row 2: 浏览按钮 (only 本地文件),col 1 右对齐
            btnBrowse.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            btnBrowse.Dock = DockStyle.None;
            tblSource.Controls.Add(btnBrowse, 1, 2);

            // Row 3: CalDAV 账户下拉框 (only CalDAV,跨 3 列)
            tblSource.Controls.Add(lblCalDavAccount, 0, 3);
            tblSource.SetColumnSpan(cmbCalDavAccounts, 3);
            tblSource.Controls.Add(cmbCalDavAccounts, 1, 3);

            // Row 4: 服务器 (only CalDAV)
            tblSource.Controls.Add(lblServerUrl, 0, 4);
            tblSource.SetColumnSpan(txtServerUrl, 3);
            tblSource.Controls.Add(txtServerUrl, 1, 4);

            // Row 5: 用户名 + 密码 (only CalDAV)
            tblSource.Controls.Add(lblUsername, 0, 5);
            tblSource.Controls.Add(txtUsername, 1, 5);
            tblSource.Controls.Add(lblPassword, 2, 5);
            tblSource.Controls.Add(txtPassword, 3, 5);

            grpSource.Controls.Add(tblSource);

            // ========== 区块 3: 操作区 (GroupBox) ==========
            var grpAction = new GroupBox
            {
                Text = "操作",
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };

            var tblAction = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 2,
                Padding = new Padding(5),
                RowStyles = {
                    new RowStyle(SizeType.Absolute, 36),
                    new RowStyle(SizeType.Absolute, 24)
                },
                ColumnStyles = {
                    new ColumnStyle(SizeType.AutoSize),
                    new ColumnStyle(SizeType.Percent, 100),
                    new ColumnStyle(SizeType.AutoSize)
                }
            };

            btnStartSync = new Button
            {
                Text = "▶ 全量同步",
                Size = new Size(130, 36),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold)
            };
            btnStartSync.Click += BtnStartSync_Click;

            progressSync = new ProgressBar
            {
                Dock = DockStyle.Fill,
                Height = 12,
                Style = ProgressBarStyle.Continuous
            };

            lblSyncStatus = new Label
            {
                Text = "就绪",
                AutoSize = true,
                ForeColor = Color.Gray
            };

            var lblProgressPercent = new Label
            {
                Text = "0%",
                AutoSize = true,
                ForeColor = Color.Gray
            };
            this.lblProgressPercent = lblProgressPercent;

            // Row 0: 全量同步按钮 + 进度条 + 百分比
            tblAction.Controls.Add(btnStartSync, 0, 0);
            tblAction.Controls.Add(progressSync, 1, 0);
            tblAction.Controls.Add(lblProgressPercent, 2, 0);

            // Row 1: 状态 (跨 3 列)
            tblAction.Controls.Add(lblSyncStatus, 0, 1);
            tblAction.SetColumnSpan(lblSyncStatus, 3);

            grpAction.Controls.Add(tblAction);

            // ========== 主布局 ==========
            mainLayout.Controls.Add(grpAccount, 0, 0);
            mainLayout.Controls.Add(grpSource, 0, 1);
            mainLayout.Controls.Add(grpAction, 0, 2);

            this.Controls.Add(mainLayout);
            this.Controls.Add(lblCurrentEmail);  // 隐藏占位,供事件处理器引用

            // 初始化同步源配置区的可见性 (事件绑定在 SelectedIndex=0 之后,需手动调用)
            UpdateControlsVisibility(cmbSourceType.SelectedIndex);
        }

        /// <summary>
        /// 由 MainForm 在初始化时调用,加载已保存的 OAuth 账户到下拉框
        /// </summary>
        public void LoadAccounts()
        {
            if (cmbSyncAccounts != null && MainForm != null)
            {
                MainForm.LoadSavedOAuthAccountsToComboBox(cmbSyncAccounts);
            }
        }

        /// <summary>
        /// 重新加载所有账户列表(OAuth + CalDAV)
        /// </summary>
        public void ReloadAllAccounts()
        {
            LoadAccounts();
            ReloadCalDavAccounts();
        }

        /// <summary>
        /// 填充时区下拉框(UTC-12 ~ UTC+12)
        /// </summary>
        private void PopulateTimeZoneCombo(ComboBox cmb, string defaultValue)
        {
            var timeZones = new List<KeyValuePair<string, string>>();
            var cityMap = new Dictionary<int, string>
            {
                { -12, "国际日期变更线西" }, { -11, "萨摩亚" }, { -10, "夏威夷" },
                { -9, "阿拉斯加" }, { -8, "洛杉矶" }, { -7, "丹佛" },
                { -6, "芝加哥" }, { -5, "纽约" }, { -4, "圣地亚哥" },
                { -3, "圣保罗" }, { -2, "大西洋中部" }, { -1, "亚速尔群岛" },
                { 0, "伦敦" }, { 1, "巴黎" }, { 2, "开罗" },
                { 3, "莫斯科" }, { 4, "迪拜" }, { 5, "伊斯兰堡" },
                { 6, "达卡" }, { 7, "曼谷" }, { 8, "北京" },
                { 9, "东京" }, { 10, "悉尼" }, { 11, "所罗门群岛" },
                { 12, "奥克兰" }
            };

            for (int offset = -12; offset <= 12; offset++)
            {
                string city = cityMap.ContainsKey(offset) ? cityMap[offset] : "";
                // 中文习惯: 东X区 / 西X区 / 零时区
                string cnName = offset == 0 ? "零时区" : (offset > 0 ? $"东{offset}区" : $"西{-offset}区");
                string display = $"{cnName} ({city})";
                string value = offset == 0 ? "UTC" : (offset > 0 ? $"UTC+{offset}" : $"UTC{offset}");
                timeZones.Add(new KeyValuePair<string, string>(value, display));
            }
            cmb.DataSource = timeZones;
            cmb.DisplayMember = "Value";
            cmb.ValueMember = "Key";
            if (timeZones.Any(kv => kv.Key == defaultValue))
            {
                cmb.SelectedValue = defaultValue;
            }
        }

        /// <summary>
        /// 获取当前选中的源时区,用于把 CSV/ICS/VCS 里的本地时间转 UTC 后再同步到 O365
        /// </summary>
        public TimeZoneInfo GetSelectedSourceTimeZone()
        {
            var key = cmbSourceTimeZone?.SelectedValue?.ToString() ?? "UTC+8";
            int offset = 0;
            if (!string.IsNullOrEmpty(key))
            {
                if (key == "UTC") offset = 0;
                else if (key.StartsWith("UTC") && (key.Contains("+") || key.Contains("-")))
                {
                    int.TryParse(key.Substring(3), out offset);
                }
            }
            return TimeZoneInfo.CreateCustomTimeZone(key, TimeSpan.FromHours(offset), key, key);
        }

        // 同步两个时区下拉框(任一变化时把另一个 SelectedValue 同步)
        private bool _suppressTimeZoneSync;

        private void CmbSourceTimeZone_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_suppressTimeZoneSync) return;
            _suppressTimeZoneSync = true;
            try
            {
                if (cmbImportTimeZone != null && cmbSourceTimeZone.SelectedValue != null)
                {
                    cmbImportTimeZone.SelectedValue = cmbSourceTimeZone.SelectedValue;
                }
            }
            finally { _suppressTimeZoneSync = false; }
        }

        private void CmbImportTimeZone_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_suppressTimeZoneSync) return;
            _suppressTimeZoneSync = true;
            try
            {
                if (cmbSourceTimeZone != null && cmbImportTimeZone.SelectedValue != null)
                {
                    cmbSourceTimeZone.SelectedValue = cmbImportTimeZone.SelectedValue;
                }
            }
            finally { _suppressTimeZoneSync = false; }
        }

        /// <summary>
        /// 从 SettingsService 重新加载 CalDAV/CardDAV 账户到下拉框
        /// (DAV 服务器通常同时支持 CardDAV 和 CalDAV 协议,所以共享同一份账户配置)
        /// </summary>
        public void ReloadCalDavAccounts()
        {
            if (cmbCalDavAccounts == null) return;
            var calDavAccounts = SettingsService.Load().CardDavAccounts;
            var previousSelection = cmbCalDavAccounts.SelectedItem?.ToString();
            cmbCalDavAccounts.Items.Clear();
            foreach (var acc in calDavAccounts)
            {
                cmbCalDavAccounts.Items.Add(acc.Name);
            }
            if (cmbCalDavAccounts.Items.Count == 0) return;

            int idx = string.IsNullOrEmpty(previousSelection)
                ? 0
                : cmbCalDavAccounts.Items.IndexOf(previousSelection);
            cmbCalDavAccounts.SelectedIndex = idx >= 0 ? idx : 0;

            var selected = calDavAccounts.Find(a => a.Name == cmbCalDavAccounts.SelectedItem?.ToString());
            if (selected != null && txtServerUrl != null)
            {
                txtServerUrl.Text = selected.ServerUrl;
            }
        }

        private void CmbSyncAccounts_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbSyncAccounts.SelectedIndex > 0)
            {
                var selectedName = cmbSyncAccounts.SelectedItem?.ToString();
                if (!string.IsNullOrEmpty(selectedName))
                {
                    var settings = ConfigService.LoadAll();
                    var account = settings.OAuthAccounts.Find(a => a.Name == selectedName);
                    if (account != null)
                    {
                        txtClientId.Text = account.ClientId;
                        txtTenantId.Text = account.TenantId;
                        txtEmail.Text = account.Email;
                    }
                }
            }
        }

        private void CmbCalDavAccounts_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbCalDavAccounts.SelectedIndex < 0) return;
            var calDavAccounts = SettingsService.Load().CardDavAccounts;
            if (cmbCalDavAccounts.SelectedIndex >= calDavAccounts.Count) return;
            var selected = calDavAccounts[cmbCalDavAccounts.SelectedIndex];
            if (selected != null && txtServerUrl != null)
            {
                txtServerUrl.Text = selected.ServerUrl;
            }
        }

        private void UpdateControlsVisibility(int sourceType)
        {
            if (sourceType == 0 || sourceType == 1)
            {
                // CSV / ICS 本地文件 - 隐藏 CalDAV 字段并折叠其行
                lblCalDavAccount.Visible = false;
                cmbCalDavAccounts.Visible = false;
                lblServerUrl.Visible = false;
                txtServerUrl.Visible = false;
                lblUsername.Visible = false;
                txtUsername.Visible = false;
                lblPassword.Visible = false;
                txtPassword.Visible = false;
                lblSourceFile.Visible = true;
                txtSourceFile.Visible = true;
                btnBrowse.Visible = true;
                // 导入时区 (本地文件行)
                lblImportTimeZone.Visible = true;
                cmbImportTimeZone.Visible = true;
                CollapseRows(localFileMode: true);
            }
            else
            {
                // CalDAV
                lblCalDavAccount.Visible = true;
                cmbCalDavAccounts.Visible = true;
                lblServerUrl.Visible = true;
                txtServerUrl.Visible = true;
                lblUsername.Visible = true;
                txtUsername.Visible = true;
                lblPassword.Visible = true;
                txtPassword.Visible = true;
                lblSourceFile.Visible = false;
                txtSourceFile.Visible = false;
                btnBrowse.Visible = false;
                // CalDAV 模式隐藏本地文件行内的导入时区(改由"文件源时区"承担)
                lblImportTimeZone.Visible = false;
                cmbImportTimeZone.Visible = false;
                // 切换到 CalDAV 模式时刷新账户下拉框
                ReloadCalDavAccounts();
                CollapseRows(localFileMode: false);
            }
        }

        /// <summary>
        /// 折叠/展开 tblSource 的各行:
        ///   本地文件模式: Row 1 (本地文件) + Row 2 (浏览)
        ///   CalDAV 模式:   Row 3 (账户) + Row 4 (服务器) + Row 5 (凭据)
        /// </summary>
        private void CollapseRows(bool localFileMode)
        {
            if (tblSource == null || tblSource.RowStyles.Count < 6) return;
            // Row 1: 本地文件 (30px) - 仅本地文件模式显示
            tblSource.RowStyles[1].Height = localFileMode ? 30 : 0;
            // Row 2: 浏览按钮 (32px) - 仅本地文件模式显示
            tblSource.RowStyles[2].Height = localFileMode ? 32 : 0;
            // Row 3: CalDAV 账户 (30px) - 仅 CalDAV 模式显示
            tblSource.RowStyles[3].Height = localFileMode ? 0 : 30;
            // Row 4: 服务器 (30px) - 仅 CalDAV 模式显示
            tblSource.RowStyles[4].Height = localFileMode ? 0 : 30;
            // Row 5: 用户名 + 密码 (36px) - 仅 CalDAV 模式显示
            tblSource.RowStyles[5].Height = localFileMode ? 0 : 36;
        }

        // === 事件处理 ===

        private void CmbSourceType_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateControlsVisibility(cmbSourceType.SelectedIndex);
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "选择日历文件";
                openFileDialog.Filter = "所有支持格式|*.csv;*.ics;*.vcs;*.msg|ICS文件|*.ics|VCS文件|*.vcs|CSV文件|*.csv|MSG文件|*.msg";
                openFileDialog.Multiselect = false;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtSourceFile.Text = openFileDialog.FileName;
                }
            }
        }

        private void BtnOAuthLogin_Click(object sender, EventArgs e)
        {
            if (MainForm == null)
            {
                lblSyncStatus.Text = "未设置 MainForm,无法登录";
                lblSyncStatus.ForeColor = Color.Red;
                return;
            }

            // 将本控件的字段值同步到 MainForm (MainForm 的 OAuth 流程从其私有字段读取)
            MainForm.SetO365TextFields(txtClientId.Text, txtTenantId.Text, txtEmail.Text);

            // 触发 MainForm 的 OAuth 登录
            MainForm.BtnO365OAuthLogin_Click(sender, e);

            // 轮询 MainForm 的 OAuth 状态
            var timer = new System.Windows.Forms.Timer { Interval = 3000 };
            timer.Tick += (ts, te) =>
            {
                timer.Stop();
                timer.Dispose();
                if (MainForm.IsO365OAuthConnected)
                {
                    lblCurrentEmail.Text = MainForm.O365OAuthEmail;
                    lblCurrentEmail.ForeColor = Color.Green;
                    lblSyncStatus.Text = "登录成功!";
                    lblSyncStatus.ForeColor = Color.Green;
                }
                else
                {
                    lblSyncStatus.Text = "登录失败,请重试";
                    lblSyncStatus.ForeColor = Color.Red;
                }
            };
            timer.Start();
        }

        private void BtnStartSync_Click(object sender, EventArgs e)
        {
            if (MainForm == null)
            {
                lblSyncStatus.Text = "未设置 MainForm,无法同步";
                lblSyncStatus.ForeColor = Color.Red;
                return;
            }

            var sourceType = cmbSourceType.SelectedIndex;

            // CalDAV 不需要选择源文件
            if (sourceType != 2)
            {
                if (string.IsNullOrWhiteSpace(txtSourceFile.Text))
                {
                    lblSyncStatus.Text = "请先选择源文件";
                    lblSyncStatus.ForeColor = Color.Red;
                    return;
                }
            }

            if (!MainForm.IsO365OAuthConnected || string.IsNullOrEmpty(MainForm.O365AccessToken))
            {
                lblSyncStatus.Text = "请先登录";
                lblSyncStatus.ForeColor = Color.Red;
                return;
            }

            // 确保 MainForm._office365Service 已连接
            var o365Service = MainForm.Office365Service;
            if (o365Service == null || !o365Service.IsOAuthConnected)
            {
                o365Service = new Office365ImportService();
                if (!o365Service.ConnectWithOAuth(MainForm.O365OAuthEmail, MainForm.O365AccessToken))
                {
                    lblSyncStatus.Text = "OAuth2 连接失败,请重新登录";
                    lblSyncStatus.ForeColor = Color.Red;
                    return;
                }
                MainForm.SetO365Service(o365Service);
            }

            var filePath = txtSourceFile.Text;
            var extension = string.IsNullOrEmpty(filePath) ? "" : Path.GetExtension(filePath).ToLower();

            // 禁用按钮
            btnStartSync.Enabled = false;

            // 进度回调:从后台线程调用,通过 Invoke 切到 UI 线程更新 lblSyncStatus + 进度条 + 百分比
            Action<int, int, string> progressReporter = (current, total, msg) =>
            {
                if (this.IsDisposed || !this.IsHandleCreated) return;
                try
                {
                    this.BeginInvoke(new Action(() =>
                    {
                        if (this.IsDisposed) return;
                        lblSyncStatus.Text = msg;
                        lblSyncStatus.ForeColor = Color.Blue;
                        // 计算百分比:避免除零;超过 100 截断
                        if (total > 0)
                        {
                            int percent = (int)((long)current * 100 / total);
                            if (percent < 0) percent = 0;
                            if (percent > 100) percent = 100;
                            // 必须先确认是 Continuous 风格(初始状态),否则 Value 赋值无效
                            if (progressSync.Style != ProgressBarStyle.Continuous)
                                progressSync.Style = ProgressBarStyle.Continuous;
                            // 进度条 Value 不能超过 Maximum - 1 (WinForms 限制)
                            if (progressSync.Maximum != 100) progressSync.Maximum = 100;
                            if (progressSync.Value != percent) progressSync.Value = percent;
                            lblProgressPercent.Text = $"{percent}%";
                        }
                    }));
                }
                catch { /* 控件已销毁时静默忽略 */ }
            };

            Task.Run(() =>
            {
                try
                {
                    this.Invoke(new Action(() =>
                    {
                        lblSyncStatus.Text = "正在同步...";
                        lblSyncStatus.ForeColor = Color.Blue;
                        // 切到 Continuous 并初始化进度条 / 百分比
                        progressSync.Style = ProgressBarStyle.Continuous;
                        progressSync.Minimum = 0;
                        progressSync.Maximum = 100;
                        progressSync.Value = 0;
                        lblProgressPercent.Text = "0%";
                    }));

                    string resultMessage = "";
                    var sourceTz = GetSelectedSourceTimeZone();

                    // 本地文件同步
                    if (sourceType == 0) // CSV
                    {
                        if (extension == ".csv")
                            resultMessage = MainForm.SyncCalendarFromCsv(filePath, sourceTz, progressReporter);
                        else
                            resultMessage = "CSV 模式下请选择 .csv 文件";
                    }
                    else if (sourceType == 1) // ICS
                    {
                        if (extension == ".ics")
                            resultMessage = MainForm.SyncCalendarFromIcs(filePath, sourceTz, progressReporter);
                        else if (extension == ".vcs")
                            resultMessage = MainForm.SyncCalendarFromVcs(filePath, sourceTz, progressReporter);
                        else
                            resultMessage = "ICS 模式下请选择 .ics / .vcs 文件";
                    }
                    else if (sourceType == 2) // CalDAV
                    {
                        // CalDAV 服务暂未实现
                        resultMessage = "CalDAV 日历同步功能暂未实现,目前仅支持本地文件 (CSV/ICS)";
                    }

                    // 完成后的UI更新
                    this.Invoke(new Action(() =>
                    {
                        progressSync.Style = ProgressBarStyle.Continuous;
                        if (resultMessage.Contains("完成"))
                        {
                            progressSync.Value = 100;
                            lblProgressPercent.Text = "100%";
                        }
                        else
                        {
                            progressSync.Value = 0;
                            lblProgressPercent.Text = "0%";
                        }
                        lblSyncStatus.Text = resultMessage;
                        lblSyncStatus.ForeColor = resultMessage.Contains("完成") ? Color.Green : Color.Red;
                        btnStartSync.Enabled = true;
                    }));
                }
                catch (Exception ex)
                {
                    this.Invoke(new Action(() =>
                    {
                        progressSync.Style = ProgressBarStyle.Continuous;
                        progressSync.Value = 0;
                        lblProgressPercent.Text = "0%";
                        lblSyncStatus.Text = $"同步失败: {ex.Message}";
                        lblSyncStatus.ForeColor = Color.Red;
                        btnStartSync.Enabled = true;
                    }));
                    Serilog.Log.Error(ex, "日历同步失败");
                }
            });
        }
    }
}
