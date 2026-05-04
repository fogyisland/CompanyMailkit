using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MailConverter
{
    public partial class RegistrationForm : Form
    {
        private TextBox txtUserName;
        private TextBox txtUserEmail;
        private TextBox txtOrganization;
        private TextBox txtSerialNumber;
        private Label lblMacAddress;
        private Label lblSoftwareName;
        private Label lblSoftwareVersion;
        private Button btnRegister;
        private Button btnCancel;
        private CheckBox chkShowMac;
        private Label lblStatus;
        private Label lblRegStatus;
        private Label lblRemainingDays;
        private Button btnUnregister;

        private readonly RegistrationService _registrationService;
        private readonly string _softwareName = "xiaomingMailtoolkitCompany";
        private readonly string _softwareVersion = "1.0.0";
        private string _macAddress;
        private bool _isRegistered;

        public RegistrationForm()
        {
            _registrationService = new RegistrationService();
            var settings = SettingsService.Load();
            _isRegistered = settings.IsRegistered;

            InitializeCustomComponents();
            LoadMacAddress();

            if (_isRegistered)
            {
                ShowLicenseInfo(settings);
            }
        }

        private void InitializeCustomComponents()
        {
            this.Text = "软件使用注册";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ClientSize = new System.Drawing.Size(550, _isRegistered ? 420 : 400);
            this.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);

            if (_isRegistered)
            {
                InitializeLicensePanel();
            }
            else
            {
                InitializeRegistrationPanel();
            }
        }

        private void InitializeLicensePanel()
        {
            int y = 20;
            int labelWidth = 120;
            int leftMargin = 150;
            int contentWidth = 350;

            // 已注册标题
            var lblTitle = new Label
            {
                Text = "正式版授权信息",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(410, 30),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(0, 120, 215),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };
            this.Controls.Add(lblTitle);

            y += 50;

            // 授权状态
            var lblStatusTitle = new Label
            {
                Text = "授权状态:",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(labelWidth, 25),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };
            var lblStatusValue = new Label
            {
                Text = "试用版",
                Location = new System.Drawing.Point(leftMargin, y),
                Size = new System.Drawing.Size(100, 25),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.Orange
            };
            lblRemainingDays = new Label
            {
                Text = "",
                Location = new System.Drawing.Point(leftMargin + 100, y),
                Size = new System.Drawing.Size(200, 25),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(0, 120, 215)
            };
            this.Controls.Add(lblStatusTitle);
            this.Controls.Add(lblStatusValue);
            this.Controls.Add(lblRemainingDays);

            y += 35;

            // 用户姓名
            var lblUserTitle = new Label
            {
                Text = "用户姓名:",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(labelWidth, 25),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };
            var lblUserValue = new Label
            {
                Text = "",
                Location = new System.Drawing.Point(leftMargin, y),
                Size = new System.Drawing.Size(contentWidth, 25),
                Tag = "lblUserName"
            };
            this.Controls.Add(lblUserTitle);
            this.Controls.Add(lblUserValue);

            y += 35;

            // 用户邮箱
            var lblEmailTitle = new Label
            {
                Text = "用户邮箱:",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(labelWidth, 25),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };
            var lblEmailValue = new Label
            {
                Text = "",
                Location = new System.Drawing.Point(leftMargin, y),
                Size = new System.Drawing.Size(contentWidth, 25),
                Tag = "lblUserEmail"
            };
            this.Controls.Add(lblEmailTitle);
            this.Controls.Add(lblEmailValue);

            y += 35;

            // 组织/公司
            var lblOrgTitle = new Label
            {
                Text = "组织/公司:",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(labelWidth, 25),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };
            var lblOrgValue = new Label
            {
                Text = "",
                Location = new System.Drawing.Point(leftMargin, y),
                Size = new System.Drawing.Size(contentWidth, 25),
                Tag = "lblOrg"
            };
            this.Controls.Add(lblOrgTitle);
            this.Controls.Add(lblOrgValue);

            y += 35;

            // 注册日期（订阅版注册日期）
            var lblDateTitle = new Label
            {
                Text = "注册日期:",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(labelWidth, 25),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };
            var lblDateValue = new Label
            {
                Text = "",
                Location = new System.Drawing.Point(leftMargin, y),
                Size = new System.Drawing.Size(contentWidth, 25),
                Tag = "lblDate"
            };
            this.Controls.Add(lblDateTitle);
            this.Controls.Add(lblDateValue);

            y += 35;

            // 第一次运行日期（正式版激活日期）
            var lblFirstRunTitle = new Label
            {
                Text = "激活日期:",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(labelWidth, 25),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };
            var lblFirstRunValue = new Label
            {
                Text = "-",
                Location = new System.Drawing.Point(leftMargin, y),
                Size = new System.Drawing.Size(contentWidth, 25),
                Tag = "lblFirstRun"
            };
            this.Controls.Add(lblFirstRunTitle);
            this.Controls.Add(lblFirstRunValue);

            y += 35;

            // 到期日期
            var lblExpireTitle = new Label
            {
                Text = "到期日期:",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(labelWidth, 25),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };
            var lblExpireValue = new Label
            {
                Text = "",
                Location = new System.Drawing.Point(leftMargin, y),
                Size = new System.Drawing.Size(contentWidth, 25),
                Tag = "lblExpire"
            };
            this.Controls.Add(lblExpireTitle);
            this.Controls.Add(lblExpireValue);

            y += 35;

            // MAC地址
            var lblMacTitle = new Label
            {
                Text = "MAC地址:",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(labelWidth, 25),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };
            lblMacAddress = new Label
            {
                Text = "",
                Location = new System.Drawing.Point(leftMargin, y),
                Size = new System.Drawing.Size(contentWidth - 60, 25),
                BackColor = System.Drawing.Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(5, 5, 5, 5)
            };
            chkShowMac = new CheckBox
            {
                Text = "显示",
                Location = new System.Drawing.Point(leftMargin + contentWidth - 45, y),
                Size = new System.Drawing.Size(50, 25)
            };
            chkShowMac.CheckedChanged += (s, e) =>
            {
                var settings = SettingsService.Load();
                lblMacAddress.Text = chkShowMac.Checked ? settings.RegisteredMacAddress : MaskMac(settings.RegisteredMacAddress);
            };
            this.Controls.Add(lblMacTitle);
            this.Controls.Add(lblMacAddress);
            this.Controls.Add(chkShowMac);

            y += 45;

            // 正式版激活区域
            var lblActivateTitle = new Label
            {
                Text = "正式版激活:",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(labelWidth, 25),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };
            txtSerialNumber = new TextBox
            {
                Location = new System.Drawing.Point(leftMargin, y),
                Size = new System.Drawing.Size(220, 25)
            };
            var btnActivate = new Button
            {
                Text = "激活",
                Location = new System.Drawing.Point(leftMargin + 240, y),
                Size = new System.Drawing.Size(120, 35),
                BackColor = System.Drawing.Color.FromArgb(0, 120, 215),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat,
                Tag = "btnActivate"
            };
            btnActivate.Click += BtnActivate_Click;
            this.Controls.Add(lblActivateTitle);
            this.Controls.Add(txtSerialNumber);
            this.Controls.Add(btnActivate);

            y += 40;

            // 状态标签
            lblStatus = new Label
            {
                Text = "",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(510, 25),
                ForeColor = System.Drawing.Color.Gray,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };
            this.Controls.Add(lblStatus);

            y += 40;

            // 注销按钮
            btnUnregister = new Button
            {
                Text = "注销授权",
                Location = new System.Drawing.Point(140, y),
                Size = new System.Drawing.Size(120, 35),
                BackColor = System.Drawing.Color.FromArgb(220, 53, 69),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnUnregister.Click += BtnUnregister_Click;
            this.Controls.Add(btnUnregister);

            // 关闭按钮
            btnCancel = new Button
            {
                Text = "关闭",
                Location = new System.Drawing.Point(250, y),
                Size = new System.Drawing.Size(120, 35),
                FlatStyle = FlatStyle.Flat
            };
            btnCancel.Click += (s, e) => { this.Close(); };
            this.Controls.Add(btnCancel);
        }

        private void ShowLicenseInfo(AppSettings settings)
        {
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is Label lbl && lbl.Tag != null)
                {
                    switch (lbl.Tag.ToString())
                    {
                        case "lblUserName":
                            lbl.Text = settings.RegisteredUserName;
                            break;
                        case "lblUserEmail":
                            lbl.Text = settings.RegisteredUserEmail;
                            break;
                        case "lblOrg":
                            lbl.Text = settings.RegisteredOrganization ?? "-";
                            break;
                        case "lblDate":
                            lbl.Text = settings.RegisterDate?.ToString("yyyy-MM-dd") ?? "-";
                            break;
                        case "lblFirstRun":
                            lbl.Text = settings.FirstRunDate?.ToString("yyyy-MM-dd") ?? "-";
                            break;
                        case "lblExpire":
                            lbl.Text = settings.RegisterExpireDate ?? "-";
                            break;
                    }
                }
            }
            if (lblMacAddress != null && chkShowMac != null)
            {
                lblMacAddress.Text = MaskMac(settings.RegisteredMacAddress);
            }
            if (lblRemainingDays != null)
            {
                if (settings.RegisterRemainingDays.HasValue)
                {
                    lblRemainingDays.Text = $"剩余 {settings.RegisterRemainingDays.Value} 天";
                }
            }
            if (txtSerialNumber != null)
            {
                txtSerialNumber.Text = settings.RegisterSerialNumber ?? "";
            }
            // 如果已注册，禁用用户信息字段（仅允许修改序列号）
            if (settings.IsRegistered)
            {
                if (txtUserName != null) txtUserName.Enabled = false;
                if (txtUserEmail != null) txtUserEmail.Enabled = false;
                if (txtOrganization != null) txtOrganization.Enabled = false;
            }
            // 更新授权状态显示
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is Label lbl && lbl.Tag == null && lbl.Text == "试用版")
                {
                    lbl.Text = string.IsNullOrEmpty(settings.RegisterSerialNumber) ? "试用版" : "订阅版";
                    lbl.ForeColor = string.IsNullOrEmpty(settings.RegisterSerialNumber) ? System.Drawing.Color.Orange : System.Drawing.Color.Green;
                    break;
                }
            }
        }

        private void InitializeRegistrationPanel()
        {
            int y = 20;
            int labelWidth = 120;
            int inputWidth = 380;
            int leftMargin = 150;

            // 软件名称
            var lblNameTitle = new Label
            {
                Text = "软件名称:",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(labelWidth, 25),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };
            lblSoftwareName = new Label
            {
                Text = _softwareName,
                Location = new System.Drawing.Point(leftMargin, y),
                Size = new System.Drawing.Size(inputWidth, 25),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(0, 120, 215)
            };
            this.Controls.Add(lblNameTitle);
            this.Controls.Add(lblSoftwareName);

            y += 35;

            // 软件版本
            var lblVerTitle = new Label
            {
                Text = "软件版本:",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(labelWidth, 25),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };
            lblSoftwareVersion = new Label
            {
                Text = _softwareVersion,
                Location = new System.Drawing.Point(leftMargin, y),
                Size = new System.Drawing.Size(inputWidth, 25)
            };
            this.Controls.Add(lblVerTitle);
            this.Controls.Add(lblSoftwareVersion);

            y += 35;

            // MAC地址
            var lblMacTitle = new Label
            {
                Text = "MAC地址:",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(labelWidth, 25),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };
            lblMacAddress = new Label
            {
                Text = "检测中...",
                Location = new System.Drawing.Point(leftMargin, y),
                Size = new System.Drawing.Size(inputWidth - 70, 25),
                BackColor = System.Drawing.Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(5, 5, 5, 5)
            };
            chkShowMac = new CheckBox
            {
                Text = "显示",
                Location = new System.Drawing.Point(leftMargin + inputWidth - 60, y),
                Size = new System.Drawing.Size(50, 25)
            };
            chkShowMac.CheckedChanged += (s, e) =>
            {
                lblMacAddress.Text = chkShowMac.Checked ? _macAddress : MaskMac(_macAddress);
            };
            this.Controls.Add(lblMacTitle);
            this.Controls.Add(lblMacAddress);
            this.Controls.Add(chkShowMac);

            y += 35;

            // 用户姓名
            var lblUserTitle = new Label
            {
                Text = "用户姓名:",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(labelWidth, 25),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };
            txtUserName = new TextBox
            {
                Location = new System.Drawing.Point(leftMargin, y),
                Size = new System.Drawing.Size(inputWidth, 25)
            };
            this.Controls.Add(lblUserTitle);
            this.Controls.Add(txtUserName);

            y += 35;

            // 用户邮箱
            var lblEmailTitle = new Label
            {
                Text = "用户邮箱:",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(labelWidth, 25),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };
            txtUserEmail = new TextBox
            {
                Location = new System.Drawing.Point(leftMargin, y),
                Size = new System.Drawing.Size(inputWidth, 25)
            };
            this.Controls.Add(lblEmailTitle);
            this.Controls.Add(txtUserEmail);

            y += 35;

            // 组织/公司
            var lblOrgTitle = new Label
            {
                Text = "组织/公司:",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(labelWidth, 25),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };
            txtOrganization = new TextBox
            {
                Location = new System.Drawing.Point(leftMargin, y),
                Size = new System.Drawing.Size(inputWidth, 25)
            };
            this.Controls.Add(lblOrgTitle);
            this.Controls.Add(txtOrganization);

            y += 40;

            // 状态标签
            lblRegStatus = new Label
            {
                Text = "请填写注册信息",
                Location = new System.Drawing.Point(20, y),
                Size = new System.Drawing.Size(510, 25),
                ForeColor = System.Drawing.Color.Gray,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };
            this.Controls.Add(lblRegStatus);

            y += 35;

            // 按钮
            btnRegister = new Button
            {
                Text = "注册",
                Location = new System.Drawing.Point(160, y),
                Size = new System.Drawing.Size(120, 35),
                BackColor = System.Drawing.Color.FromArgb(0, 120, 215),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnRegister.Click += BtnRegister_Click;

            btnCancel = new Button
            {
                Text = "取消",
                Location = new System.Drawing.Point(270, y),
                Size = new System.Drawing.Size(120, 35),
                FlatStyle = FlatStyle.Flat
            };
            btnCancel.Click += (s, e) => { this.Close(); };

            this.Controls.Add(btnRegister);
            this.Controls.Add(btnCancel);
        }

        private void LoadMacAddress()
        {
            Task.Run(() =>
            {
                _macAddress = _registrationService.GetPhysicalMacAddress();
                this.Invoke(new Action(() =>
                {
                    if (string.IsNullOrEmpty(_macAddress))
                    {
                        if (lblMacAddress != null)
                        {
                            lblMacAddress.Text = "未检测到物理网卡";
                            lblMacAddress.ForeColor = System.Drawing.Color.Red;
                        }
                    }
                    else
                    {
                        if (lblMacAddress != null)
                        {
                            lblMacAddress.Text = MaskMac(_macAddress);
                        }
                    }
                }));
            });
        }

        private string MaskMac(string mac)
        {
            if (string.IsNullOrEmpty(mac) || mac.Length < 8)
                return mac;
            return mac.Substring(0, 2) + "-**-**-**-**-" + mac.Substring(mac.Length - 2);
        }

        private async void BtnRegister_Click(object sender, EventArgs e)
        {
            // 验证必填字段
            if (string.IsNullOrWhiteSpace(txtUserName.Text))
            {
                lblRegStatus.Text = "请输入用户姓名";
                lblRegStatus.ForeColor = System.Drawing.Color.Red;
                txtUserName.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(txtUserEmail.Text))
            {
                lblRegStatus.Text = "请输入用户邮箱";
                lblRegStatus.ForeColor = System.Drawing.Color.Red;
                txtUserEmail.Focus();
                return;
            }

            if (!IsValidEmail(txtUserEmail.Text))
            {
                lblRegStatus.Text = "请输入有效的邮箱地址";
                lblRegStatus.ForeColor = System.Drawing.Color.Red;
                txtUserEmail.Focus();
                return;
            }

            if (string.IsNullOrEmpty(_macAddress))
            {
                lblRegStatus.Text = "未检测到物理网卡，无法注册";
                lblRegStatus.ForeColor = System.Drawing.Color.Red;
                return;
            }

            btnRegister.Enabled = false;
            lblRegStatus.Text = "正在提交注册信息...";
            lblRegStatus.ForeColor = System.Drawing.Color.Blue;
            Application.DoEvents();

            try
            {
                var result = await _registrationService.RegisterAsync(
                    _softwareName,
                    _softwareVersion,
                    txtUserName.Text.Trim(),
                    txtUserEmail.Text.Trim(),
                    txtOrganization.Text.Trim(),
                    _macAddress
                );

                if (result.Success)
                {
                    // 保存注册信息
                    var settings = SettingsService.Load();
                    settings.IsRegistered = true;
                    settings.RegisteredUserName = txtUserName.Text.Trim();
                    settings.RegisteredUserEmail = txtUserEmail.Text.Trim();
                    settings.RegisteredOrganization = txtOrganization.Text.Trim();
                    settings.RegisteredMacAddress = _macAddress;
                    settings.RegisterSerialNumber = "";
                    settings.RegisterDate = DateTime.Now;
                    settings.RegisterExpireDate = result.ExpireDate;

                    // 计算剩余天数
                    if (DateTime.TryParse(result.ExpireDate, out var expireDate))
                    {
                        settings.RegisterRemainingDays = (int)Math.Max(0, (expireDate - DateTime.Now).TotalDays);
                    }
                    else
                    {
                        settings.RegisterRemainingDays = result.RemainingDays;
                    }
                    ConfigService.SaveAll(settings);
                    RegistryService.SaveRegistration(settings);

                    lblRegStatus.Text = $"注册成功！剩余 {settings.RegisterRemainingDays} 天";
                    lblRegStatus.ForeColor = System.Drawing.Color.Green;

                    // 禁用用户信息字段（仅允许修改序列号）
                    if (txtUserName != null) txtUserName.Enabled = false;
                    if (txtUserEmail != null) txtUserEmail.Enabled = false;
                    if (txtOrganization != null) txtOrganization.Enabled = false;

                    await Task.Delay(2000);
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    lblRegStatus.Text = result.Message;
                    lblRegStatus.ForeColor = System.Drawing.Color.Red;
                    btnRegister.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Error(ex, "注册异常");
                lblRegStatus.Text = "注册异常: " + ex.Message;
                lblRegStatus.ForeColor = System.Drawing.Color.Red;
                btnRegister.Enabled = true;
            }
        }

        private async void BtnActivate_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtSerialNumber.Text))
            {
                lblStatus.Text = "请输入序列号";
                lblStatus.ForeColor = System.Drawing.Color.Red;
                txtSerialNumber.Focus();
                return;
            }

            if (string.IsNullOrEmpty(_macAddress))
            {
                lblStatus.Text = "未检测到物理网卡，无法激活";
                lblStatus.ForeColor = System.Drawing.Color.Red;
                return;
            }

            var btnActivate = sender as Button;
            if (btnActivate != null) btnActivate.Enabled = false;
            lblStatus.Text = "正在激活正式版...";
            lblStatus.ForeColor = System.Drawing.Color.Blue;
            Application.DoEvents();

            var activationCode = txtSerialNumber.Text.Trim();
            var macAddress = _macAddress;
            var settings = SettingsService.Load();
            var installDate = DateTime.Now.ToString("yyyy-MM-dd");

            Program.RegistrationLogger.Information("【正式版激活】开始激活流程");
            Program.RegistrationLogger.Information("【正式版激活】授权码: {ActivationCode}, MAC: {MacAddress}, 用户: {UserName}, 邮箱: {UserEmail}, 安装日期: {InstallDate}",
                activationCode, macAddress, settings.RegisteredUserName, settings.RegisteredUserEmail, installDate);

            try
            {
                var result = await _registrationService.ActivateByCodeAsync(
                    activationCode,
                    macAddress,
                    settings.RegisteredUserName,
                    settings.RegisteredUserEmail,
                    installDate
                );

                if (result.Success)
                {
                    Program.RegistrationLogger.Information("【正式版激活】API返回成功: 消息={Message}, 剩余天数={RemainingDays}, 到期日期={ExpireDate}, 注册日期={InstallDate}, 总天数={TotalDays}",
                        result.Message, result.RemainingDays, result.ExpireDate, result.InstallDate, result.TotalDays);

                    settings.RegisterSerialNumber = activationCode;

                    // 计算新到期日期
                    DateTime currentExpireDate;
                    bool hasCurrentSerial = !string.IsNullOrEmpty(settings.RegisterSerialNumber);
                    bool isExpired = !DateTime.TryParse(settings.RegisterExpireDate, out currentExpireDate) || currentExpireDate <= DateTime.Now;

                    if (!isExpired && result.TotalDays.HasValue)
                    {
                        // 未过期激活或续期：从当前到期日期 + 总天数
                        var newExpireDate = currentExpireDate.AddDays(result.TotalDays.Value);
                        settings.RegisterExpireDate = newExpireDate.ToString("yyyy-MM-dd");
                        settings.RegisterRemainingDays = (int)Math.Max(0, (newExpireDate - DateTime.Now).TotalDays);
                        Program.RegistrationLogger.Information("【正式版激活】{Case}: 原到期日期={OldExpireDate}, +{TotalDays}天, 新到期日期={NewExpireDate}",
                            hasCurrentSerial ? "续期" : "首次激活", currentExpireDate.ToString("yyyy-MM-dd"), result.TotalDays.Value, settings.RegisterExpireDate);
                    }
                    else
                    {
                        // 已过期激活：激活日期为今天，到期日期 = 今天 + 总天数
                        var activateDate = DateTime.Now;
                        if (result.TotalDays.HasValue)
                        {
                            var newExpireDate = activateDate.AddDays(result.TotalDays.Value);
                            settings.RegisterExpireDate = newExpireDate.ToString("yyyy-MM-dd");
                            settings.RegisterRemainingDays = result.TotalDays.Value;
                            settings.RegisterDate = activateDate;
                            Program.RegistrationLogger.Information("【正式版激活】已过期重新激活: 激活日期={ActivateDate}, +{TotalDays}天, 新到期日期={NewExpireDate}",
                                activateDate.ToString("yyyy-MM-dd"), result.TotalDays.Value, settings.RegisterExpireDate);
                        }
                        else
                        {
                            settings.RegisterExpireDate = result.ExpireDate;
                            settings.RegisterRemainingDays = result.RemainingDays;
                        }
                    }

                    if (!settings.FirstRunDate.HasValue)
                    {
                        settings.FirstRunDate = DateTime.Now;
                    }
                    ConfigService.SaveAll(settings);
                    RegistryService.SaveRegistration(settings);

                    Program.RegistrationLogger.Information("【正式版激活】激活成功！剩余 {RemainingDays} 天，到期日期: {ExpireDate}", settings.RegisterRemainingDays, settings.RegisterExpireDate);

                    lblStatus.Text = $"激活成功！剩余 {settings.RegisterRemainingDays} 天";
                    lblStatus.ForeColor = System.Drawing.Color.Green;

                    // 更新状态显示
                    foreach (Control ctrl in this.Controls)
                    {
                        if (ctrl is Label lbl && lbl.Tag == null)
                        {
                            if (lbl.Text == "试用版")
                            {
                                lbl.Text = "订阅版";
                                lbl.ForeColor = System.Drawing.Color.Green;
                                break;
                            }
                        }
                        if (ctrl is Label lbl2 && lbl2.Tag != null && lbl2.Tag.ToString() == "lblFirstRun")
                        {
                            lbl2.Text = DateTime.Now.ToString("yyyy-MM-dd");
                        }
                    }

                    await Task.Delay(2000);
                }
                else
                {
                    Program.RegistrationLogger.Warning("【正式版激活】API返回失败: {ErrorMessage}", result.Message);
                    lblStatus.Text = result.Message;
                    lblStatus.ForeColor = System.Drawing.Color.Red;
                    this.Refresh();
                    Application.DoEvents();
                    if (btnActivate != null) btnActivate.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                Program.RegistrationLogger.Error(ex, "【正式版激活】激活过程发生异常: ActivationCode={ActivationCode}, MacAddress={MacAddress}", activationCode, macAddress);
                lblStatus.Text = "激活异常: " + ex.Message;
                lblStatus.ForeColor = System.Drawing.Color.Red;
                if (btnActivate != null) btnActivate.Enabled = true;
            }
        }

        private void BtnUnregister_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show(
                "确定要注销当前授权吗？\n注销后将需要重新注册才能使用软件。",
                "确认注销",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                var settings = SettingsService.Load();
                settings.IsRegistered = false;
                settings.RegisteredUserName = "";
                settings.RegisteredUserEmail = "";
                settings.RegisteredOrganization = "";
                settings.RegisteredMacAddress = "";
                settings.RegisterSerialNumber = "";
                settings.RegisterDate = null;
                settings.RegisterRemainingDays = null;
                settings.RegisterExpireDate = null;
                settings.FirstRunDate = null;
                ConfigService.SaveAll(settings);
                RegistryService.ClearRegistration();

                MessageBox.Show("已注销授权。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
        }

        private bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }
    }
}
