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
        private Label lblCardDavAccount;
        private Label lblServerUrl;
        private Label lblUsername;
        private Label lblPassword;
        private Label lblSourceFile;
        private ComboBox cmbSyncAccounts;

        private List<string> _selectedContactUrls = new List<string>();

        /// <summary>
        /// 由 MainForm 在创建时注入,用于调用 BtnO365OAuthLogin_Click 等方法
        /// (命名避开 UserControl 自身的只读 ParentForm 属性)
        /// </summary>
        public MainForm MainForm { get; set; }

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

            // 已保存账户下拉框 (右侧)
            var lblSavedAccount = new Label
            {
                Text = "已保存账户:",
                Location = new Point(300, 47),
                AutoSize = true,
                Name = "lblSyncContactsSavedAccount"
            };
            cmbSyncAccounts = new ComboBox
            {
                Location = new Point(380, 44),
                Size = new Size(200, 22),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Name = "cmbSyncContactsSavedAccount"
            };
            cmbSyncAccounts.SelectedIndexChanged += CmbSyncAccounts_SelectedIndexChanged;

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
            btnOAuthLogin.Click += BtnOAuthLogin_Click;

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

            lblCardDavAccount = new Label { Text = "提供商:", Location = new Point(30, 213), AutoSize = true, Visible = false, Name = "lblSyncContactsCardDavAcc" };
            cmbCardDavAccounts = new ComboBox
            {
                Location = new Point(110, 211),
                Size = new Size(180, 22),
                Visible = false,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Name = "cmbSyncContactsCardDav"
            };
            cmbCardDavAccounts.SelectedIndexChanged += CmbCardDavAccounts_SelectedIndexChanged;
            lblCardDavTip = new Label
            {
                Text = "请注意:QQ邮箱需使用授权码",
                Location = new Point(295, 213),
                Size = new Size(200, 20),
                ForeColor = Color.Red,
                Visible = false,
                Name = "lblSyncContactsCardDavTip"
            };

            lblServerUrl = new Label { Text = "服务器:", Location = new Point(30, 240), AutoSize = true, Name = "lblSyncContactsServer" };
            txtServerUrl = new TextBox { Location = new Point(110, 238), Size = new Size(320, 22), Name = "txtSyncContactsServerUrl" };

            lblUsername = new Label { Text = "用户名:", Location = new Point(30, 268), AutoSize = true, Name = "lblSyncContactsUsername" };
            txtUsername = new TextBox { Location = new Point(110, 266), Size = new Size(180, 22), Name = "txtSyncContactsUsername" };
            txtUsername.TextChanged += TxtUsername_TextChanged;
            lblPassword = new Label { Text = "密码:", Location = new Point(310, 268), AutoSize = true, Name = "lblSyncContactsPassword" };
            txtPassword = new TextBox { Location = new Point(350, 266), Size = new Size(120, 22), UseSystemPasswordChar = true, Name = "txtSyncContactsPassword" };

            lblSourceFile = new Label { Text = "本地文件:", Location = new Point(30, 240), AutoSize = true, Name = "lblSyncContactsSourceFile" };
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
            btnSelectContacts.Click += BtnSelectContacts_Click;
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
            btnIncrementalSync.Click += BtnIncrementalSync_Click;
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
            this.Controls.Add(lblSavedAccount); this.Controls.Add(cmbSyncAccounts);
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
                lblCardDavAccount.Visible = false;
                cmbCardDavAccounts.Visible = false;
                lblCardDavTip.Visible = false;
                lblServerUrl.Visible = false;
                txtServerUrl.Visible = false;
                lblUsername.Visible = false;
                txtUsername.Visible = false;
                lblPassword.Visible = false;
                txtPassword.Visible = false;
                lblSourceFile.Visible = true;
                txtSourceFile.Visible = true;
                btnBrowse.Visible = true;
                btnSelectContacts.Visible = false;
                btnIncrementalSync.Visible = false;
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
            }
            else
            {
                // 企业微信API 等其他
                lblCardDavAccount.Visible = false;
                cmbCardDavAccounts.Visible = false;
                lblCardDavTip.Visible = false;
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
                    case 3: // 企业微信API
                        txtServerUrl.Text = "https://qyapi.weixin.qq.com/cgi-bin/";
                        break;
                }
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

            // CardDAV/企业微信API 不需要选择源文件
            if (sourceType != 2 && sourceType != 3)
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
                                resultMessage = MainForm.SyncContactsFromCsv(filePath);
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
                                        var result = MainForm.SyncContactsFromVcf(f);
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
                                resultMessage = MainForm.SyncContactsFromVcf(filePath);
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
                                                            successCount++;
                                                        else
                                                            errorCount++;
                                                    }
                                                    else
                                                    {
                                                        skipCount++;
                                                    }
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            errorCount++;
                                            Serilog.Log.Warning("同步联系人失败: {Error}", ex.Message);
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
                    else if (sourceType == 3) // 企业微信API - 占位
                    {
                        resultMessage = "企业微信API同步功能暂未实现";
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
