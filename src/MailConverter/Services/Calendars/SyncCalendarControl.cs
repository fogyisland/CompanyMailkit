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
        private Label lblServerUrl;
        private Label lblUsername;
        private Label lblPassword;
        private Label lblSourceFile;
        private ComboBox cmbSyncAccounts;

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
                    new RowStyle(SizeType.Absolute, 145),  // 账户配置
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
                RowCount = 3,
                Padding = new Padding(5),
                RowStyles = {
                    new RowStyle(SizeType.Absolute, 30),
                    new RowStyle(SizeType.Absolute, 30),
                    new RowStyle(SizeType.Absolute, 36)
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
                Text = "▶ 使用 Microsoft 登录",
                AutoSize = true,
                Padding = new Padding(10, 5, 10, 5),
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

            // Row 2: 邮箱 + 登录按钮
            tblAccount.Controls.Add(lblEmail, 0, 2);
            tblAccount.Controls.Add(txtEmail, 1, 2);
            tblAccount.Controls.Add(btnOAuthLogin, 2, 2);
            tblAccount.SetColumnSpan(btnOAuthLogin, 2);

            grpAccount.Controls.Add(tblAccount);

            // ========== 区块 2: 同步源配置 (GroupBox) ==========
            var grpSource = new GroupBox
            {
                Text = "同步源配置",
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };

            var tblSource = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 4,
                Padding = new Padding(5),
                RowStyles = {
                    new RowStyle(SizeType.Absolute, 30),  // Row 0: 数据来源
                    new RowStyle(SizeType.Absolute, 30),  // Row 1: 本地文件
                    new RowStyle(SizeType.Absolute, 30),  // Row 2: 服务器
                    new RowStyle(SizeType.Absolute, 36)   // Row 3: 用户名 + 密码
                },
                ColumnStyles = {
                    new ColumnStyle(SizeType.AutoSize),
                    new ColumnStyle(SizeType.Percent, 40),
                    new ColumnStyle(SizeType.AutoSize),
                    new ColumnStyle(SizeType.Percent, 60)
                }
            };

            var lblSourceType = new Label { Text = "数据来源:", AutoSize = true };
            cmbSourceType = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Dock = DockStyle.Fill
            };
            cmbSourceType.Items.AddRange(new object[] { "本地文件(ICS)", "本地文件(VCS)", "CalDAV" });
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

            // CalDAV 相关控件
            lblServerUrl = new Label { Text = "服务器:", AutoSize = true };
            txtServerUrl = new TextBox { Dock = DockStyle.Fill };

            lblUsername = new Label { Text = "用户名:", AutoSize = true };
            txtUsername = new TextBox { Dock = DockStyle.Fill };
            lblPassword = new Label { Text = "密码:", AutoSize = true };
            txtPassword = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = true };

            // Row 0: 数据来源 (combo 跨 3 列)
            tblSource.Controls.Add(lblSourceType, 0, 0);
            tblSource.SetColumnSpan(cmbSourceType, 3);
            tblSource.Controls.Add(cmbSourceType, 1, 0);

            // Row 1: 本地文件 + 浏览 (only 本地文件)
            tblSource.Controls.Add(lblSourceFile, 0, 1);
            tblSource.Controls.Add(txtSourceFile, 1, 1);
            tblSource.SetColumnSpan(txtSourceFile, 2);
            tblSource.Controls.Add(btnBrowse, 3, 1);

            // Row 2: 服务器 (only CalDAV)
            tblSource.Controls.Add(lblServerUrl, 0, 2);
            tblSource.SetColumnSpan(txtServerUrl, 3);
            tblSource.Controls.Add(txtServerUrl, 1, 2);

            // Row 3: 用户名 + 密码 (only CalDAV)
            tblSource.Controls.Add(lblUsername, 0, 3);
            tblSource.Controls.Add(txtUsername, 1, 3);
            tblSource.Controls.Add(lblPassword, 2, 3);
            tblSource.Controls.Add(txtPassword, 3, 3);

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
                // 本地文件
                lblServerUrl.Visible = false;
                txtServerUrl.Visible = false;
                lblUsername.Visible = false;
                txtUsername.Visible = false;
                lblPassword.Visible = false;
                txtPassword.Visible = false;
                lblSourceFile.Visible = true;
                txtSourceFile.Visible = true;
                btnBrowse.Visible = true;
            }
            else
            {
                // CalDAV
                lblServerUrl.Visible = true;
                txtServerUrl.Visible = true;
                lblUsername.Visible = true;
                txtUsername.Visible = true;
                lblPassword.Visible = true;
                txtPassword.Visible = true;
                lblSourceFile.Visible = false;
                txtSourceFile.Visible = false;
                btnBrowse.Visible = false;
            }
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
                openFileDialog.Filter = "所有支持格式|*.ics;*.vcs;*.msg|ICS文件|*.ics|VCS文件|*.vcs|MSG文件|*.msg";
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

            Task.Run(() =>
            {
                try
                {
                    this.Invoke(new Action(() =>
                    {
                        lblSyncStatus.Text = "正在同步...";
                        lblSyncStatus.ForeColor = Color.Blue;
                        progressSync.Style = ProgressBarStyle.Marquee;
                    }));

                    string resultMessage = "";

                    // 本地文件同步
                    if (sourceType == 0 || sourceType == 1)
                    {
                        if (extension == ".ics")
                            resultMessage = MainForm.SyncCalendarFromIcs(filePath);
                        else if (extension == ".vcs")
                            resultMessage = MainForm.SyncCalendarFromVcs(filePath);
                        else
                            resultMessage = "不支持的日历文件格式";
                    }
                    else if (sourceType == 2) // CalDAV
                    {
                        // CalDAV 服务暂未实现
                        resultMessage = "CalDAV 日历同步功能暂未实现,目前仅支持本地文件 (ICS/VCS)";
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
                    Serilog.Log.Error(ex, "日历同步失败");
                }
            });
        }
    }
}
