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
            // 标题
            var lblTitle = new Label
            {
                Text = "个人同步日历到 Office 365",
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
            txtClientId = new TextBox { Location = new Point(110, 73), Size = new Size(160, 22), Name = "txtSyncCalendarClientId" };

            var lblTenantId = new Label { Text = "租户ID:", Location = new Point(280, 75), AutoSize = true };
            txtTenantId = new TextBox { Location = new Point(330, 73), Size = new Size(120, 22), Name = "txtSyncCalendarTenantId" };

            var lblEmail = new Label { Text = "邮箱:", Location = new Point(30, 100), AutoSize = true };
            txtEmail = new TextBox { Location = new Point(110, 98), Size = new Size(220, 22), Name = "txtSyncCalendarEmail" };

            btnOAuthLogin = new Button
            {
                Text = "登录",
                Location = new Point(350, 96),
                Size = new Size(80, 25),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                Name = "btnSyncCalendarLogin"
            };
            btnOAuthLogin.Click += BtnOAuthLogin_Click;

            lblCurrentEmail = new Label
            {
                Text = "未登录",
                Location = new Point(110, 125),
                AutoSize = true,
                ForeColor = Color.Gray,
                Name = "lblSyncCalendarCurrent"
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
                Name = "cmbSyncCalendarSourceType"
            };
            cmbSourceType.Items.AddRange(new object[] { "本地文件(ICS)", "本地文件(VCS)", "CalDAV" });
            cmbSourceType.SelectedIndex = 0;
            cmbSourceType.SelectedIndexChanged += CmbSourceType_SelectedIndexChanged;

            lblServerUrl = new Label { Text = "服务器:", Location = new Point(30, 240), AutoSize = true, Name = "lblSyncCalendarServer" };
            txtServerUrl = new TextBox { Location = new Point(110, 238), Size = new Size(320, 22), Name = "txtSyncCalendarServerUrl" };

            lblUsername = new Label { Text = "用户名:", Location = new Point(30, 268), AutoSize = true, Name = "lblSyncCalendarUsername" };
            txtUsername = new TextBox { Location = new Point(110, 266), Size = new Size(180, 22), Name = "txtSyncCalendarUsername" };
            lblPassword = new Label { Text = "密码:", Location = new Point(310, 268), AutoSize = true, Name = "lblSyncCalendarPassword" };
            txtPassword = new TextBox { Location = new Point(350, 266), Size = new Size(120, 22), UseSystemPasswordChar = true, Name = "txtSyncCalendarPassword" };

            lblSourceFile = new Label { Text = "本地文件:", Location = new Point(30, 240), AutoSize = true, Name = "lblSyncCalendarSourceFile" };
            txtSourceFile = new TextBox { Location = new Point(110, 238), Size = new Size(270, 22), Name = "txtSyncCalendarSourceFile" };
            btnBrowse = new Button
            {
                Text = "浏览...",
                Location = new Point(390, 236),
                Size = new Size(75, 25),
                Name = "btnSyncCalendarBrowse"
            };
            btnBrowse.Click += BtnBrowse_Click;

            // ===== 按钮 =====
            btnStartSync = new Button
            {
                Text = "全量同步",
                Location = new Point(220, 305),
                Size = new Size(100, 30),
                Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                Name = "btnSyncCalendarStart"
            };
            btnStartSync.Click += BtnStartSync_Click;

            // ===== 进度 =====
            progressSync = new ProgressBar
            {
                Location = new Point(20, 355),
                Size = new Size(570, 18),
                Style = ProgressBarStyle.Continuous,
                Name = "progressSyncCalendar"
            };
            lblSyncStatus = new Label
            {
                Location = new Point(20, 380),
                Size = new Size(570, 18),
                ForeColor = Color.Gray,
                Name = "lblSyncCalendarStatus"
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
            this.Controls.Add(lblServerUrl); this.Controls.Add(txtServerUrl);
            this.Controls.Add(lblUsername); this.Controls.Add(txtUsername);
            this.Controls.Add(lblPassword); this.Controls.Add(txtPassword);
            this.Controls.Add(lblSourceFile); this.Controls.Add(txtSourceFile);
            this.Controls.Add(btnBrowse);
            this.Controls.Add(btnStartSync);
            this.Controls.Add(progressSync);
            this.Controls.Add(lblSyncStatus);
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
