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
        private Button btnDownloadTemplate;
        private Button btnSelectContacts;
        private Button btnIncrementalSync;
        private Button btnStartSync;
        private ProgressBar progressSync;
        private Label lblSyncStatus;
        private Label lblProgressPercent;
        private Label lblCardDavAccount;
        private Label lblServerUrl;
        private Label lblUsername;
        private Label lblPassword;
        private Label lblSourceFile;
        private Label lblCardDavProviderList;
        private ComboBox cmbSyncAccounts;
        private TableLayoutPanel tblSource;
        private ToolTip _wechatWorkTip;

        private List<string> _selectedContactUrls = new List<string>();

        /// <summary>
        /// 由 MainForm 在创建时注入,用于调用 BtnO365OAuthLogin_Click 等方法
        /// (命名避开 UserControl 自身的只读 ParentForm 属性)
        /// </summary>
        public MainForm MainForm { get; set; }

        public SyncContactsControl()
        {
            Dock = DockStyle.Fill;
            _wechatWorkTip = new ToolTip
            {
                AutoPopDelay = 10000,
                InitialDelay = 300,
                ReshowDelay = 200,
                IsBalloon = false
            };
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
                ColumnCount = 6,
                RowCount = 5,
                Padding = new Padding(5),
                RowStyles = {
                    new RowStyle(SizeType.Absolute, 30),  // Row 0: 数据来源
                    new RowStyle(SizeType.Absolute, 30),  // Row 1: CardDAV 提供商 / 本地文件
                    new RowStyle(SizeType.Absolute, 30),  // Row 2: 服务器
                    new RowStyle(SizeType.Absolute, 36),  // Row 3: 用户名 + 密码
                    new RowStyle(SizeType.Absolute, 40)   // Row 4: CardDAV 按钮
                },
                ColumnStyles = {
                    new ColumnStyle(SizeType.AutoSize),     // 0: 标签
                    new ColumnStyle(SizeType.Percent, 25),  // 1: 文本框左侧
                    new ColumnStyle(SizeType.AutoSize),     // 2: 文本框右侧 / 提示
                    new ColumnStyle(SizeType.AutoSize),     // 3: 下载模板按钮
                    new ColumnStyle(SizeType.Percent, 35),  // 4: 文本框 / 浏览按钮
                    new ColumnStyle(SizeType.Percent, 30)   // 5: CardDAV 提供商列表(右侧)
                }
            };

            var lblSourceType = new Label { Text = "数据来源:", AutoSize = true };
            cmbSourceType = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Dock = DockStyle.Fill
            };
            cmbSourceType.Items.AddRange(new object[] { "本地文件(CSV)", "本地文件(VCF)", "CardDAV", "企业微信API(内部)", "企业微信(客户联系)", "Exchange", "Office 365 额外租户" });
            cmbSourceType.SelectedIndex = 0;
            cmbSourceType.SelectedIndexChanged += CmbSourceType_SelectedIndexChanged;

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

            // 下载模板按钮 (本地文件模式可见)
            btnDownloadTemplate = new Button
            {
                Text = "下载模板",
                Width = 80,
                BackColor = Color.FromArgb(245, 245, 245),
                FlatStyle = FlatStyle.Flat,
                Visible = false
            };
            var templateMenu = new ContextMenuStrip();
            templateMenu.Items.Add("下载 CSV 模板", null, (s, e) => DownloadTemplate("csv"));
            templateMenu.Items.Add("下载 VCF 模板", null, (s, e) => DownloadTemplate("vcf"));
            btnDownloadTemplate.ContextMenuStrip = templateMenu;
            btnDownloadTemplate.Click += (s, e) => templateMenu.Show(btnDownloadTemplate, 0, btnDownloadTemplate.Height);

            // CardDAV / 企业微信相关控件
            lblCardDavAccount = new Label { Text = "提供商:", AutoSize = true, Visible = false };
            cmbCardDavAccounts = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Dock = DockStyle.Fill,
                Visible = false
            };
            cmbCardDavAccounts.SelectedIndexChanged += CmbCardDavAccounts_SelectedIndexChanged;
            lblCardDavTip = new Label
            {
                Text = "(QQ邮箱需使用授权码)",
                AutoSize = true,
                ForeColor = Color.Red,
                Visible = false
            };

            // 右侧 CardDAV 提供商参考列表(只读,带表情)
            lblCardDavProviderList = new Label
            {
                Text = "常见 CardDAV 服务器:\r\n" +
                       "🍎 iCloud\r\n" +
                       "📧 Gmail (已停用)\r\n" +
                       "🐧 QQ邮箱 (授权码)\r\n" +
                       "📮 Outlook / Microsoft 365\r\n" +
                       "📨 Yahoo\r\n" +
                       "💬 飞书",
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.TopLeft,
                ForeColor = Color.DarkSlateGray,
                BackColor = Color.FromArgb(245, 248, 250),
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(6, 4, 6, 4),
                Font = new Font("Segoe UI Emoji", 8F),
                Visible = false
            };

            lblServerUrl = new Label { Text = "服务器:", AutoSize = true };
            txtServerUrl = new TextBox { Dock = DockStyle.Fill };

            lblUsername = new Label { Text = "用户名:", AutoSize = true };
            txtUsername = new TextBox { Dock = DockStyle.Fill };
            txtUsername.TextChanged += TxtUsername_TextChanged;
            lblPassword = new Label { Text = "密码:", AutoSize = true };
            txtPassword = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = true };

            // CardDAV 专用按钮
            btnSelectContacts = new Button
            {
                Text = "选择联系人",
                AutoSize = true,
                Padding = new Padding(10, 5, 10, 5),
                BackColor = Color.FromArgb(255, 193, 7),
                ForeColor = Color.Black,
                FlatStyle = FlatStyle.Flat,
                Visible = false
            };
            btnSelectContacts.Click += BtnSelectContacts_Click;
            btnIncrementalSync = new Button
            {
                Text = "▶ 增量同步",
                AutoSize = true,
                Padding = new Padding(10, 5, 10, 5),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Visible = false
            };
            btnIncrementalSync.Click += BtnIncrementalSync_Click;

            // Row 0: 数据来源 (combo 跨 3 列)
            tblSource.Controls.Add(lblSourceType, 0, 0);
            tblSource.SetColumnSpan(cmbSourceType, 3);
            tblSource.Controls.Add(cmbSourceType, 1, 0);

            // 右侧:CardDAV 提供商参考列表(只读,跨 5 行,只占 Column 5)
            tblSource.Controls.Add(lblCardDavProviderList, 5, 0);
            tblSource.SetRowSpan(lblCardDavProviderList, 5);

            // Row 1: CardDAV 提供商 + 提示 (only CardDAV) + 本地文件 (only 本地文件) 互相隐藏
            tblSource.Controls.Add(lblCardDavAccount, 0, 1);
            tblSource.Controls.Add(cmbCardDavAccounts, 1, 1);
            tblSource.Controls.Add(lblCardDavTip, 2, 1);
            tblSource.SetColumnSpan(cmbCardDavAccounts, 2);

            tblSource.Controls.Add(lblSourceFile, 0, 1);
            tblSource.Controls.Add(txtSourceFile, 1, 1);
            tblSource.SetColumnSpan(txtSourceFile, 2);
            tblSource.Controls.Add(btnDownloadTemplate, 3, 1);
            tblSource.Controls.Add(btnBrowse, 4, 1);

            // Row 2: 服务器 (only CardDAV/WeChat)
            tblSource.Controls.Add(lblServerUrl, 0, 2);
            tblSource.SetColumnSpan(txtServerUrl, 3);
            tblSource.Controls.Add(txtServerUrl, 1, 2);

            // Row 3: 用户名 + 密码 (only CardDAV/WeChat)
            tblSource.Controls.Add(lblUsername, 0, 3);
            tblSource.Controls.Add(txtUsername, 1, 3);
            tblSource.Controls.Add(lblPassword, 2, 3);
            tblSource.Controls.Add(txtPassword, 3, 3);

            // Row 4: CardDAV 按钮 (only CardDAV)
            tblSource.Controls.Add(btnSelectContacts, 0, 4);
            tblSource.Controls.Add(btnIncrementalSync, 1, 4);

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
                    new RowStyle(SizeType.Absolute, 36),  // Row 0: 按钮 + 进度条 + 百分比
                    new RowStyle(SizeType.Absolute, 24)   // Row 1: 状态
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

            lblProgressPercent = new Label
            {
                Text = "0%",
                AutoSize = true,
                ForeColor = Color.Gray
            };

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
        /// 重新加载所有账户列表（OAuth + CardDAV），从磁盘读取最新数据
        /// </summary>
        public void ReloadAllAccounts()
        {
            LoadAccounts();
            ReloadCardDavAccounts();
        }

        public void ReloadCardDavAccounts()
        {
            if (cmbCardDavAccounts == null) return;
            var cardDavAccounts = SettingsService.Load().CardDavAccounts;
            var previousSelection = cmbCardDavAccounts.SelectedItem?.ToString();
            cmbCardDavAccounts.Items.Clear();
            foreach (var acc in cardDavAccounts)
            {
                cmbCardDavAccounts.Items.Add(acc.Name);
            }
            if (cmbCardDavAccounts.Items.Count == 0) return;

            int idx = string.IsNullOrEmpty(previousSelection)
                ? 0
                : cmbCardDavAccounts.Items.IndexOf(previousSelection);
            cmbCardDavAccounts.SelectedIndex = idx >= 0 ? idx : 0;

            var selected = cardDavAccounts.Find(a => a.Name == cmbCardDavAccounts.SelectedItem?.ToString());
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

        private void UpdateControlsVisibility(int sourceType)
        {
            if (sourceType == 0 || sourceType == 1)
            {
                // 本地文件 - 隐藏 CardDAV/WeChat 字段并折叠其行
                lblCardDavAccount.Visible = false;
                cmbCardDavAccounts.Visible = false;
                lblCardDavTip.Visible = false;
                lblCardDavProviderList.Visible = false;
                lblServerUrl.Visible = false;
                txtServerUrl.Visible = false;
                lblUsername.Visible = false;
                txtUsername.Visible = false;
                lblPassword.Visible = false;
                txtPassword.Visible = false;
                lblSourceFile.Visible = true;
                txtSourceFile.Visible = true;
                btnBrowse.Visible = true;
                btnDownloadTemplate.Visible = true;
                btnSelectContacts.Visible = false;
                btnIncrementalSync.Visible = false;
                CollapseRows(localFileMode: true);
            }
            else if (sourceType == 2)
            {
                // CardDAV
                var cardDavAccounts = SettingsService.Load().CardDavAccounts;
                cmbCardDavAccounts.Items.Clear();
                foreach (var acc in cardDavAccounts)
                {
                    cmbCardDavAccounts.Items.Add(acc.Name);
                }
                if (cardDavAccounts.Count > 0)
                {
                    cmbCardDavAccounts.SelectedIndex = 0;
                    txtServerUrl.Text = cardDavAccounts[0].ServerUrl;
                }
                lblCardDavAccount.Visible = true;
                cmbCardDavAccounts.Visible = true;
                lblCardDavTip.Visible = true;
                lblCardDavProviderList.Visible = true;
                lblServerUrl.Visible = true;
                txtServerUrl.Visible = true;
                lblUsername.Visible = true;
                txtUsername.Visible = true;
                lblPassword.Visible = true;
                txtPassword.Visible = true;
                btnSelectContacts.Visible = true;
                btnIncrementalSync.Visible = true;
                lblSourceFile.Visible = false;
                txtSourceFile.Visible = false;
                btnBrowse.Visible = false;
                btnDownloadTemplate.Visible = false;
                CollapseRows(localFileMode: false, showCardDavButtons: true);
            }
            else
            {
                // 企业微信API 等其他
                lblCardDavAccount.Visible = false;
                cmbCardDavAccounts.Visible = false;
                lblCardDavTip.Visible = false;
                lblCardDavProviderList.Visible = false;
                lblServerUrl.Visible = true;
                txtServerUrl.Visible = true;
                lblUsername.Visible = true;
                txtUsername.Visible = true;
                lblPassword.Visible = true;
                txtPassword.Visible = true;
                btnSelectContacts.Visible = false;
                btnIncrementalSync.Visible = false;
                lblSourceFile.Visible = false;
                txtSourceFile.Visible = false;
                btnBrowse.Visible = false;
                btnDownloadTemplate.Visible = false;
                CollapseRows(localFileMode: false, showCardDavButtons: false);
            }

            if (sourceType == 4)
            {
                // 企业微信 (客户联系 / 外部联系人) - 复用三输入框, 标签由 SetSourceFieldLabels 切到 API/CorpID/客户联系 Secret
                lblCardDavAccount.Visible = false;
                cmbCardDavAccounts.Visible = false;
                lblCardDavTip.Visible = false;
                lblCardDavProviderList.Visible = false;
                lblServerUrl.Visible = true;
                txtServerUrl.Visible = true;
                lblUsername.Visible = true;
                txtUsername.Visible = true;
                lblPassword.Visible = true;
                txtPassword.Visible = true;
                lblSourceFile.Visible = false;
                txtSourceFile.Visible = false;
                btnBrowse.Visible = false;
                btnDownloadTemplate.Visible = false;
                btnSelectContacts.Visible = false;
                btnIncrementalSync.Visible = false;
                SetSourceFieldLabels(4);
                CollapseRows(localFileMode: false, showCardDavButtons: false);
            }
            else if (sourceType == 5)
            {
                // Exchange - 复用 txtServerUrl/Username/Password, 标签由 SetSourceFieldLabels 切到 EWS URL/用户名/密码
                lblCardDavAccount.Visible = false;
                cmbCardDavAccounts.Visible = false;
                lblCardDavTip.Visible = false;
                lblCardDavProviderList.Visible = false;
                lblServerUrl.Visible = true;
                txtServerUrl.Visible = true;
                lblUsername.Visible = true;
                txtUsername.Visible = true;
                lblPassword.Visible = true;
                txtPassword.Visible = true;
                lblSourceFile.Visible = false;
                txtSourceFile.Visible = false;
                btnBrowse.Visible = false;
                btnDownloadTemplate.Visible = false;
                btnSelectContacts.Visible = false;
                btnIncrementalSync.Visible = false;
                SetSourceFieldLabels(5);
                CollapseRows(localFileMode: false, showCardDavButtons: false);
            }
            else if (sourceType == 6)
            {
                // Office 365 额外租户 - 复用三输入框, 标签由 SetSourceFieldLabels 切到源 Client ID/Tenant ID/Email
                lblCardDavAccount.Visible = false;
                cmbCardDavAccounts.Visible = false;
                lblCardDavTip.Visible = false;
                lblCardDavProviderList.Visible = false;
                lblServerUrl.Visible = true;
                txtServerUrl.Visible = true;
                lblUsername.Visible = true;
                txtUsername.Visible = true;
                lblPassword.Visible = true;
                txtPassword.Visible = true;
                lblSourceFile.Visible = false;
                txtSourceFile.Visible = false;
                btnBrowse.Visible = false;
                btnDownloadTemplate.Visible = false;
                btnSelectContacts.Visible = false;
                btnIncrementalSync.Visible = false;
                SetSourceFieldLabels(6);
                CollapseRows(localFileMode: false, showCardDavButtons: false);
            }
            else
            {
                // 其他本地文件 / CardDAV / 企业微信API 模式恢复默认标签 (以防从 idx 4/5/6 切回)
                SetSourceFieldLabels(sourceType);
            }
        }

        /// <summary>
        /// 折叠/展开 tblSource 的服务器行(2)、凭据行(3)、CardDAV 按钮行(4)
        /// </summary>
        private void CollapseRows(bool localFileMode, bool showCardDavButtons = false)
        {
            if (tblSource == null || tblSource.RowStyles.Count < 5) return;
            // Row 2: 服务器 (30px)
            tblSource.RowStyles[2].Height = localFileMode ? 0 : 30;
            // Row 3: 用户名 + 密码 (36px)
            tblSource.RowStyles[3].Height = localFileMode ? 0 : 36;
            // Row 4: CardDAV 按钮 (40px) - 仅 CardDAV 模式显示
            tblSource.RowStyles[4].Height = (!localFileMode && showCardDavButtons) ? 40 : 0;
        }

        /// <summary>
        /// 按数据来源切换 txtServerUrl/Username/Password 的标签和密码掩码
        /// 索引 4 = Exchange: EWS URL / 用户名 / 密码
        /// 索引 5 = Office 365 额外租户: 源 Client ID / 源 Tenant ID / 源 Email (Email 不掩码)
        /// 其他 = 恢复默认: 服务器 / 用户名 / 密码
        /// </summary>
        private void SetSourceFieldLabels(int sourceType)
        {
            if (sourceType == 3)
            {
                // 企业微信 API (内部通讯录)
                lblServerUrl.Text = "API 地址:";
                lblUsername.Text = "CorpID:";
                lblPassword.Text = "CorpSecret (自建应用):";
                txtPassword.UseSystemPasswordChar = true;
                _wechatWorkTip.SetToolTip(txtPassword,
                    "必须使用「自建应用」的 Secret, 不是「通讯录同步 Secret」。\n" +
                    "在企业微信管理后台 → 应用管理 → 自建应用 → Secret 栏获取。\n" +
                    "通讯录同步 Secret 只能写入或读 userid 列表, 读不到姓名/手机/邮箱等详情。");
            }
            else if (sourceType == 4)
            {
                // 企业微信 API (客户联系 / 外部联系人) - 待 Task C 实现
                lblServerUrl.Text = "API 地址:";
                lblUsername.Text = "CorpID:";
                lblPassword.Text = "CorpSecret (客户联系):";
                txtPassword.UseSystemPasswordChar = true;
                _wechatWorkTip.SetToolTip(txtPassword,
                    "必须使用「客户联系」应用的 Secret, 不是「通讯录同步 Secret」也不是「自建应用 Secret」。\n" +
                    "在企业微信管理后台 → 客户联系 → API → 客户联系 Secret 栏获取。");
            }
            else if (sourceType == 5)
            {
                // Exchange - 复用 txtServerUrl/Username/Password
                lblServerUrl.Text = "EWS URL:";
                lblUsername.Text = "用户名:";
                lblPassword.Text = "密码:";
                txtPassword.UseSystemPasswordChar = true;
            }
            else if (sourceType == 6)
            {
                // Office 365 额外租户
                lblServerUrl.Text = "源 Client ID:";
                lblUsername.Text = "源 Tenant ID:";
                lblPassword.Text = "源 Email:";
                // Email 不是密码, 关掉掩码以便用户能看清
                txtPassword.UseSystemPasswordChar = false;
            }
            else
            {
                lblServerUrl.Text = "服务器:";
                lblUsername.Text = "用户名:";
                lblPassword.Text = "密码:";
                txtPassword.UseSystemPasswordChar = true;
            }
        }

        // === 事件处理 ===

        private void CmbSourceType_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateControlsVisibility(cmbSourceType.SelectedIndex);

            // 自动填充服务器地址
            var email = txtEmail.Text;
            if (!string.IsNullOrEmpty(email) && email.Contains("@"))
            {
                switch (cmbSourceType.SelectedIndex)
                {
                    case 2: // CardDAV
                        var domain = email.Split('@')[1].ToLower();
                        if (domain.Contains("163.com") || domain.Contains("126.com") || domain.Contains("yeah.net"))
                            txtServerUrl.Text = "https://carddav.netease.com/carddav/addressbook/" + email.Split('@')[0];
                        else if (domain.Contains("gmail.com") || domain.Contains("googlemail.com"))
                            txtServerUrl.Text = "https://www.googleapis.com/carddav/v1/principals/" + email + "/";
                        else if (domain.Contains("outlook.com") || domain.Contains("hotmail.com") || domain.Contains("live.com"))
                            txtServerUrl.Text = "https://outlook.office.com/carddav/principals/" + Uri.EscapeDataString(email) + "/";
                        else
                            txtServerUrl.Text = "";
                        break;
                }
            }

            // 企业微信 API URL 是固定的 (与邮箱无关), 切到 idx 3/4 时若为空则自动填默认
            if ((cmbSourceType.SelectedIndex == 3 || cmbSourceType.SelectedIndex == 4)
                && string.IsNullOrWhiteSpace(txtServerUrl.Text))
            {
                txtServerUrl.Text = "https://qyapi.weixin.qq.com/cgi-bin/";
            }
        }

        private void TxtUsername_TextChanged(object sender, EventArgs e)
        {
            if (cmbSourceType.SelectedIndex == 2 && cmbCardDavAccounts.SelectedIndex >= 0)
            {
                var cardDavAccounts = SettingsService.Load().CardDavAccounts;
                if (cmbCardDavAccounts.SelectedIndex >= cardDavAccounts.Count) return;
                var selectedAccount = cardDavAccounts[cmbCardDavAccounts.SelectedIndex];
                var username = txtUsername.Text;

                if (selectedAccount.Provider == "Gmail" && !string.IsNullOrEmpty(username))
                {
                    txtServerUrl.Text = "https://www.googleapis.com/carddav/v1/principals/" + username + "/";
                }
            }
        }

        private void CmbCardDavAccounts_SelectedIndexChanged(object sender, EventArgs e)
        {
            var cardDavAccounts = SettingsService.Load().CardDavAccounts;
            if (cmbCardDavAccounts.SelectedIndex >= 0 && cmbCardDavAccounts.SelectedIndex < cardDavAccounts.Count)
            {
                var selectedAccount = cardDavAccounts[cmbCardDavAccounts.SelectedIndex];
                var username = txtUsername.Text;

                string serverUrl = selectedAccount.ServerUrl;
                if (selectedAccount.Provider == "Gmail" && !string.IsNullOrEmpty(username))
                {
                    serverUrl = "https://www.googleapis.com/carddav/v1/principals/" + username + "/";
                }
                txtServerUrl.Text = serverUrl;
            }
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                var sourceType = cmbSourceType.SelectedItem?.ToString() ?? "";
                openFileDialog.Title = "选择联系人文件";
                if (sourceType.Contains("CSV"))
                {
                    openFileDialog.Filter = "CSV文件|*.csv";
                    openFileDialog.Multiselect = false;
                }
                else if (sourceType.Contains("VCF"))
                {
                    openFileDialog.Filter = "VCF文件|*.vcf;*.vcard";
                    openFileDialog.Multiselect = true;
                }
                else
                {
                    openFileDialog.Filter = "所有支持格式|*.csv;*.vcf;*.vcard;*.msg|CSV文件|*.csv|VCF文件|*.vcf;*.vcard|MSG文件|*.msg";
                    openFileDialog.Multiselect = false;
                }
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (openFileDialog.FileNames.Length > 1)
                        txtSourceFile.Text = string.Join(";", openFileDialog.FileNames);
                    else
                        txtSourceFile.Text = openFileDialog.FileName;

                    var extension = Path.GetExtension(openFileDialog.FileName).ToLower();
                    if (extension == ".csv")
                    {
                        btnStartSync.PerformClick();
                    }
                }
            }
        }

        private void DownloadTemplate(string format)
        {
            using (var saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Title = "保存导入模板";
                if (format == "csv")
                {
                    saveFileDialog.Filter = "CSV文件|*.csv";
                    saveFileDialog.FileName = "contacts_template.csv";
                }
                else
                {
                    saveFileDialog.Filter = "VCF文件|*.vcf";
                    saveFileDialog.FileName = "contacts_template.vcf";
                }

                if (saveFileDialog.ShowDialog() != DialogResult.OK) return;

                try
                {
                    string content = format == "csv" ? GetCsvTemplate() : GetVcfTemplate();

                    // UTF-8 with BOM 保证 Excel 可正确识别中文表头
                    var encoding = new System.Text.UTF8Encoding(true);
                    File.WriteAllText(saveFileDialog.FileName, content, encoding);

                    lblSyncStatus.Text = $"模板已保存到: {saveFileDialog.FileName}";
                    lblSyncStatus.ForeColor = Color.Green;
                    Serilog.Log.Information("联系人导入模板已保存: {Path}", saveFileDialog.FileName);
                }
                catch (Exception ex)
                {
                    lblSyncStatus.Text = $"保存模板失败: {ex.Message}";
                    lblSyncStatus.ForeColor = Color.Red;
                    Serilog.Log.Error(ex, "保存联系人模板失败");
                }
            }
        }

        private string GetCsvTemplate()
        {
            // 列名与 SyncContactsFromCsv 中识别的列保持一致
            // 必填: 电子邮件地址 (Email) ;选填: 名/姓/电子邮件显示名称/电话/公司/职务
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("电子邮件地址,名,姓,电子邮件显示名称,电话,公司,职务");
            sb.AppendLine("zhangsan@example.com,三,张,张三,13800138000,示例公司,工程师");
            sb.AppendLine("lisi@example.com,四,李,李四,13900139000,示例公司,设计师");
            sb.AppendLine("wangwu@example.com,五,王,王五,13700137000,示例公司,产品经理");
            return sb.ToString();
        }

        private string GetVcfTemplate()
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("BEGIN:VCARD");
            sb.AppendLine("VERSION:3.0");
            sb.AppendLine("FN:张三");
            sb.AppendLine("N:张;三;;;");
            sb.AppendLine("EMAIL;TYPE=INTERNET:zhangsan@example.com");
            sb.AppendLine("TEL;TYPE=CELL:13800138000");
            sb.AppendLine("ORG:示例公司");
            sb.AppendLine("TITLE:工程师");
            sb.AppendLine("END:VCARD");
            sb.AppendLine("BEGIN:VCARD");
            sb.AppendLine("VERSION:3.0");
            sb.AppendLine("FN:李四");
            sb.AppendLine("N:李;四;;;");
            sb.AppendLine("EMAIL;TYPE=INTERNET:lisi@example.com");
            sb.AppendLine("TEL;TYPE=CELL:13900139000");
            sb.AppendLine("ORG:示例公司");
            sb.AppendLine("TITLE:设计师");
            sb.AppendLine("END:VCARD");
            return sb.ToString();
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

            // Exchange (5) / Office 365 额外租户 (6) 同步功能开发中, 先拦截
            if (sourceType == 5 || sourceType == 6)
            {
                lblSyncStatus.Text = (cmbSourceType.SelectedItem?.ToString() ?? "新源") + " 源同步功能开发中, 敬请期待";
                lblSyncStatus.ForeColor = Color.Orange;
                return;
            }

            // CardDAV/企业微信API(内部)/企业微信(客户联系) 不需要选择源文件
            if (sourceType != 2 && sourceType != 3 && sourceType != 4)
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

            // 进度回调 - 后台线程调用,通过 Invoke 更新 UI
            Action<int, int> progressCallback = (current, total) =>
            {
                this.Invoke(new Action(() =>
                {
                    if (total <= 0)
                    {
                        progressSync.Value = 0;
                        lblProgressPercent.Text = "0%";
                    }
                    else
                    {
                        int percent = (int)((double)current / total * 100);
                        progressSync.Value = Math.Min(percent, 100);
                        lblProgressPercent.Text = $"{percent}% ({current}/{total})";
                    }
                }));
            };

            // 日志回调 - 写入主窗体底部日志窗口 (AppendLogToMainWindow 内部已处理 Invoke)
            Action<string> logCallback = (message) =>
            {
                MainForm?.AppendLogToMainWindow(message);
            };

            Task.Run(() =>
            {
                try
                {
                    this.Invoke(new Action(() =>
                    {
                        lblSyncStatus.Text = "正在同步...";
                        lblSyncStatus.ForeColor = Color.Blue;
                        progressSync.Style = ProgressBarStyle.Continuous;
                        progressSync.Value = 0;
                        lblProgressPercent.Text = "0%";
                    }));

                    string resultMessage = "";

                    // 本地文件同步
                    if (sourceType == 0 || sourceType == 1)
                    {
                        if (extension == ".csv")
                        {
                            int totalCount = 0;
                            try
                            {
                                using (var reader = new StreamReader(filePath, System.Text.Encoding.UTF8))
                                using (var csv = new CsvHelper.CsvReader(reader, System.Globalization.CultureInfo.InvariantCulture))
                                {
                                    csv.Read();
                                    csv.ReadHeader();
                                    var records = csv.GetRecords<dynamic>().ToList();
                                    totalCount = records.Count;
                                }
                            }
                            catch { totalCount = 0; }

                            var confirmResult = System.Windows.Forms.DialogResult.Yes;
                            this.Invoke(new Action(() =>
                            {
                                confirmResult = System.Windows.Forms.MessageBox.Show(
                                    $"共读取到 {totalCount} 条联系人记录,是否开始同步到Office 365?",
                                    "确认同步",
                                    System.Windows.Forms.MessageBoxButtons.YesNo,
                                    System.Windows.Forms.MessageBoxIcon.Question);
                            }));

                            if (confirmResult == System.Windows.Forms.DialogResult.Yes)
                            {
                                this.Invoke(new Action(() =>
                                {
                                    lblSyncStatus.Text = "正在读取CSV文件...";
                                }));
                                resultMessage = MainForm.SyncContactsFromCsv(filePath, progressCallback, logCallback);
                            }
                            else
                            {
                                resultMessage = "已取消同步";
                            }
                        }
                        else if (extension == ".vcf" || extension == ".vcard")
                        {
                            if (filePath.Contains(";"))
                            {
                                var files = filePath.Split(';');
                                int totalSuccess = 0, totalSkip = 0, totalError = 0, totalCount = 0;
                                foreach (var f in files)
                                {
                                    if (File.Exists(f))
                                    {
                                        var result = MainForm.SyncContactsFromVcf(f, progressCallback, logCallback);
                                        var match = System.Text.RegularExpressions.Regex.Match(result, @"总计(\d+)条, 成功(\d+)条, 跳过(\d+)条, 失败(\d+)条");
                                        if (match.Success)
                                        {
                                            totalCount += int.Parse(match.Groups[1].Value);
                                            totalSuccess += int.Parse(match.Groups[2].Value);
                                            totalSkip += int.Parse(match.Groups[3].Value);
                                            totalError += int.Parse(match.Groups[4].Value);
                                        }
                                    }
                                }
                                resultMessage = $"多文件VCF同步完成: 总计{totalCount}条, 成功{totalSuccess}条, 跳过{totalSkip}条, 失败{totalError}条";
                            }
                            else
                            {
                                resultMessage = MainForm.SyncContactsFromVcf(filePath, progressCallback, logCallback);
                            }
                        }
                        else
                        {
                            resultMessage = "不支持的联系人文件格式";
                        }
                    }
                    else if (sourceType == 2) // CardDAV
                    {
                        this.Invoke(new Action(() =>
                        {
                            lblSyncStatus.Text = "正在连接CardDAV服务器...";
                        }));

                        var cardDavService = new CardDavService();
                        var serverUrl = txtServerUrl.Text;
                        var username = txtUsername.Text;
                        var password = txtPassword.Text;

                        Serilog.Log.Information("CardDAV同步参数: Url={Url}, User={User}", serverUrl, username);

                        if (string.IsNullOrWhiteSpace(serverUrl))
                        {
                            resultMessage = "请输入CardDAV服务器地址";
                        }
                        else if (!cardDavService.Connect(serverUrl, username, password))
                        {
                            resultMessage = "CardDAV连接失败,请检查服务器地址和凭据";
                        }
                        else
                        {
                            this.Invoke(new Action(() =>
                            {
                                lblSyncStatus.Text = "正在获取联系人列表...";
                            }));

                            var contacts = cardDavService.GetContactsAsync().Result;
                            if (contacts.Count == 0)
                            {
                                resultMessage = "未找到任何联系人";
                            }
                            else
                            {
                                var confirmResult = System.Windows.Forms.DialogResult.Yes;
                                this.Invoke(new Action(() =>
                                {
                                    confirmResult = System.Windows.Forms.MessageBox.Show(
                                        $"共找到 {contacts.Count} 个联系人,是否开始同步到Office 365?",
                                        "确认同步",
                                        System.Windows.Forms.MessageBoxButtons.YesNo,
                                        System.Windows.Forms.MessageBoxIcon.Question);
                                }));

                                if (confirmResult == System.Windows.Forms.DialogResult.Yes)
                                {
                                    int successCount = 0, skipCount = 0, errorCount = 0;
                                    this.Invoke(new Action(() =>
                                    {
                                        lblSyncStatus.Text = "正在同步联系人...";
                                    }));

                                    var svc = MainForm.Office365Service;
                                    int totalCount = contacts.Count;
                                    progressCallback(0, totalCount);
                                    int processed = 0;
                                    foreach (var contact in contacts)
                                    {
                                        try
                                        {
                                            var vcfContent = cardDavService.GetVCardContentAsync(contact.Url).Result;
                                            if (!string.IsNullOrEmpty(vcfContent))
                                            {
                                                var parsedContacts = cardDavService.ParseVCard(vcfContent);
                                                foreach (var pc in parsedContacts)
                                                {
                                                    if (!string.IsNullOrWhiteSpace(pc.Email))
                                                    {
                                                        if (svc.CreateContact(pc.Name, pc.Email, pc.Phone, null, null))
                                                        {
                                                            successCount++;
                                                            logCallback($"[成功] {pc.Name} <{pc.Email}>");
                                                        }
                                                        else
                                                        {
                                                            errorCount++;
                                                            logCallback($"[失败] {pc.Name} <{pc.Email}>");
                                                        }
                                                    }
                                                    else
                                                    {
                                                        skipCount++;
                                                        logCallback($"[跳过] 无邮箱: {pc.Name}");
                                                    }
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            errorCount++;
                                            Serilog.Log.Warning("同步联系人失败: {Error}", ex.Message);
                                            logCallback($"[异常] {ex.Message}");
                                        }
                                        finally
                                        {
                                            processed++;
                                            progressCallback(processed, totalCount);
                                        }
                                    }

                                    resultMessage = $"CardDAV同步完成: 成功{successCount}条, 跳过{skipCount}条, 失败{errorCount}条";
                                }
                                else
                                {
                                    resultMessage = "已取消同步";
                                }
                            }
                        }
                    }
                    else if (sourceType == 3) // 企业微信API (内部)
                    {
                        var corpId = txtUsername.Text.Trim();
                        var corpSecret = txtPassword.Text; // 不 Trim, secret 可能含特殊字符
                        var apiBase = string.IsNullOrWhiteSpace(txtServerUrl.Text) ? null : txtServerUrl.Text.Trim();

                        if (string.IsNullOrWhiteSpace(corpId) || string.IsNullOrWhiteSpace(corpSecret))
                        {
                            resultMessage = "请输入 CorpID 和 CorpSecret (自建应用)";
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
                    else if (sourceType == 4) // 企业微信 (客户联系 / 外部联系人)
                    {
                        var corpId = txtUsername.Text.Trim();
                        var corpSecret = txtPassword.Text;
                        var apiBase = string.IsNullOrWhiteSpace(txtServerUrl.Text) ? null : txtServerUrl.Text.Trim();

                        if (string.IsNullOrWhiteSpace(corpId) || string.IsNullOrWhiteSpace(corpSecret))
                        {
                            resultMessage = "请输入 CorpID 和 CorpSecret (客户联系)";
                        }
                        else
                        {
                            this.Invoke(new Action(() =>
                            {
                                lblSyncStatus.Text = "正在调用企业微信客户联系 API...";
                                lblSyncStatus.ForeColor = Color.Blue;
                            }));
                            resultMessage = MainForm.SyncContactsFromWeChatWorkExternal(
                                corpId, corpSecret, apiBase, progressCallback, logCallback);
                        }
                    }

                    // 完成后的UI更新
                    this.Invoke(new Action(() =>
                    {
                        progressSync.Style = ProgressBarStyle.Continuous;
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
                        lblSyncStatus.Text = $"同步失败: {ex.Message}";
                        lblSyncStatus.ForeColor = Color.Red;
                        btnStartSync.Enabled = true;
                    }));
                    Serilog.Log.Error(ex, "同步失败");
                }
            });
        }

        private void BtnSelectContacts_Click(object sender, EventArgs e)
        {
            var sourceType = cmbSourceType.SelectedIndex;
            if (sourceType != 2)
            {
                lblSyncStatus.Text = "选择联系人仅支持CardDAV";
                return;
            }

            var serverUrl = txtServerUrl.Text;
            var username = txtUsername.Text;
            var password = txtPassword.Text;

            if (string.IsNullOrWhiteSpace(serverUrl) || string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
            {
                lblSyncStatus.Text = "请填写完整的CardDAV同步参数";
                return;
            }

            btnSelectContacts.Enabled = false;
            lblSyncStatus.Text = "正在获取联系人列表...";
            progressSync.Style = ProgressBarStyle.Marquee;

            Task.Run(() =>
            {
                List<CardDavContact> contacts = null;
                string resultMessage;
                try
                {
                    var cardDavService = new CardDavService();
                    var connected = cardDavService.Connect(serverUrl, username, password);

                    if (!connected)
                    {
                        resultMessage = "CardDAV连接失败";
                    }
                    else
                    {
                        contacts = cardDavService.GetContactsWithETag().Result;
                        if (contacts == null || contacts.Count == 0)
                            resultMessage = "未获取到联系人";
                        else
                            resultMessage = $"获取到 {contacts.Count} 个联系人";
                    }
                }
                catch (Exception ex)
                {
                    resultMessage = $"获取联系人失败: {ex.Message}";
                    Serilog.Log.Error(ex, "获取联系人失败");
                    contacts = null;
                }

                this.Invoke(new Action(() =>
                {
                    progressSync.Style = ProgressBarStyle.Continuous;
                    btnSelectContacts.Enabled = true;

                    if (contacts != null && contacts.Count > 0)
                    {
                        var selectionForm = new ContactSelectionForm(contacts);
                        if (selectionForm.ShowDialog() == DialogResult.OK)
                        {
                            _selectedContactUrls = selectionForm.SelectedUrls;
                            lblSyncStatus.Text = $"已选择 {_selectedContactUrls.Count} 个联系人";
                            lblSyncStatus.ForeColor = Color.Green;
                        }
                        else
                        {
                            lblSyncStatus.Text = "已取消选择";
                        }
                    }
                    else
                    {
                        lblSyncStatus.Text = resultMessage;
                        lblSyncStatus.ForeColor = Color.Red;
                    }
                }));
            });
        }

        private void BtnIncrementalSync_Click(object sender, EventArgs e)
        {
            if (MainForm == null) return;
            if (!MainForm.IsO365OAuthConnected || string.IsNullOrEmpty(MainForm.O365AccessToken))
            {
                lblSyncStatus.Text = "请先连接Office 365";
                return;
            }

            var sourceType = cmbSourceType.SelectedIndex;
            if (sourceType != 2)
            {
                lblSyncStatus.Text = "增量同步仅支持CardDAV";
                return;
            }

            var serverUrl = txtServerUrl.Text;
            var username = txtUsername.Text;
            var password = txtPassword.Text;

            if (string.IsNullOrWhiteSpace(serverUrl) || string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
            {
                lblSyncStatus.Text = "请填写完整的CardDAV同步参数";
                return;
            }

            btnStartSync.Enabled = false;
            btnIncrementalSync.Enabled = false;
            lblSyncStatus.Text = "正在增量同步...";
            progressSync.Style = ProgressBarStyle.Marquee;

            // 加载本地 ETag
            var etagMapping = LoadEtags();

            Task.Run(() =>
            {
                string resultMessage;
                try
                {
                    var cardDavService = new CardDavService();
                    var connected = cardDavService.Connect(serverUrl, username, password);

                    if (!connected)
                    {
                        resultMessage = "CardDAV连接失败";
                    }
                    else
                    {
                        var contacts = cardDavService.GetContactsWithETag().Result;
                        if (contacts == null || contacts.Count == 0)
                        {
                            resultMessage = "未获取到联系人";
                        }
                        else
                        {
                            List<CardDavContact> contactsToSync;
                            if (_selectedContactUrls != null && _selectedContactUrls.Count > 0)
                            {
                                contactsToSync = contacts.Where(c => _selectedContactUrls.Contains(c.Url)).ToList();
                                Serilog.Log.Information("同步选中的 {Count} 个联系人", contactsToSync.Count);
                            }
                            else
                            {
                                contactsToSync = contacts;
                            }

                            int successCount = 0, updateCount = 0, skipCount = 0, newCount = 0;
                            var svc = MainForm.Office365Service;
                            foreach (var contact in contactsToSync)
                            {
                                try
                                {
                                    bool isNew = !etagMapping.ContainsKey(contact.Url);
                                    bool isChanged = !isNew && etagMapping[contact.Url] != contact.ETag;

                                    if (isNew || isChanged)
                                    {
                                        var vcfContent = cardDavService.GetVCardContentAsync(contact.Url).Result;
                                        if (!string.IsNullOrEmpty(vcfContent))
                                        {
                                            var parsedContacts = cardDavService.ParseVCard(vcfContent);
                                            foreach (var pc in parsedContacts)
                                            {
                                                var email = pc.Email;
                                                if (email != null && email.Contains("无邮箱"))
                                                    email = "";

                                                if (svc.CreateContact(pc.Name, email, pc.Phone, null, null))
                                                {
                                                    successCount++;
                                                    if (isNew) newCount++;
                                                    else updateCount++;
                                                }
                                            }
                                            etagMapping[contact.Url] = contact.ETag;
                                        }
                                    }
                                    else
                                    {
                                        skipCount++;
                                    }
                                }
                                catch { }
                            }

                            resultMessage = $"增量同步完成: 新增{newCount}条, 更新{updateCount}条, 跳过{skipCount}条";
                            SaveEtags(etagMapping);
                            Serilog.Log.Information(resultMessage);
                        }
                    }
                }
                catch (Exception ex)
                {
                    SaveEtags(etagMapping);
                    resultMessage = $"增量同步失败: {ex.Message}";
                    Serilog.Log.Error(ex, "增量同步失败");
                }

                this.Invoke(new Action(() =>
                {
                    progressSync.Style = ProgressBarStyle.Continuous;
                    lblSyncStatus.Text = resultMessage;
                    lblSyncStatus.ForeColor = resultMessage.Contains("完成") ? Color.Green : Color.Red;
                    btnStartSync.Enabled = true;
                    btnIncrementalSync.Enabled = true;
                }));
            });
        }

        // ETag helpers
        private string GetEtagsDir()
        {
            var targetEmail = MainForm?.O365OAuthEmail ?? "default";
            if (string.IsNullOrEmpty(targetEmail)) targetEmail = "default";
            var dirName = targetEmail.Replace("@", "_").Replace(".", "_");
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "carddav_etag", dirName);
        }

        private Dictionary<string, string> LoadEtags()
        {
            var dict = new Dictionary<string, string>();
            try
            {
                var etagFilePath = Path.Combine(GetEtagsDir(), "etag.txt");
                if (File.Exists(etagFilePath))
                {
                    foreach (var line in File.ReadAllLines(etagFilePath))
                    {
                        var parts = line.Split(new[] { '|' }, 2);
                        if (parts.Length == 2)
                            dict[parts[0]] = parts[1];
                    }
                    Serilog.Log.Information("已加载 {Count} 个ETag记录 from {Path}", dict.Count, etagFilePath);
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Warning(ex, "加载ETag文件失败");
            }
            return dict;
        }

        private void SaveEtags(Dictionary<string, string> dict)
        {
            try
            {
                var etagDir = GetEtagsDir();
                var etagFilePath = Path.Combine(etagDir, "etag.txt");
                if (!Directory.Exists(etagDir))
                    Directory.CreateDirectory(etagDir);
                var lines = dict.Select(kv => kv.Key + "|" + kv.Value).ToArray();
                File.WriteAllLines(etagFilePath, lines);
                Serilog.Log.Information("已保存 {Count} 个ETag记录到 {Path}", dict.Count, etagFilePath);
            }
            catch (Exception ex)
            {
                Serilog.Log.Warning(ex, "保存ETag文件失败");
            }
        }
    }
}
