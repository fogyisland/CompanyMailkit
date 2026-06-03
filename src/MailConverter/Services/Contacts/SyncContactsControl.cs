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
                Text = "请注意:QQ邮箱需使用授权码",
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

        // 事件占位 - Task 12 集成到 MainForm 时填充实际逻辑
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
                lblSyncStatus.Text = "登录失败,请重试";
                lblSyncStatus.ForeColor = Color.Red;
            }
        }
    }
}
