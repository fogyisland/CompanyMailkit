using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MailConverter
{
    public class ActivationForm : Form
    {
        private Label lblMacAddress;
        private Label lblStatus;
        private Label lblRemainingDays;
        private Label lblUserName;
        private Label lblUserEmail;
        private Label lblOrg;
        private Label lblDate;
        private Label lblFirstRun;
        private Label lblExpire;
        private Label lblStatusValue;
        private TextBox txtSerialNumber;
        private Button btnCancel;
        private CheckBox chkShowMac;

        private readonly RegistrationService _registrationService;

        public ActivationForm()
        {
            _registrationService = new RegistrationService();
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.Text = "软件激活";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ClientSize = new Size(550, 480);
            this.BackColor = Color.FromArgb(240, 240, 240);

            int y = 20;
            int labelWidth = 120;
            int leftMargin = 150;
            int contentWidth = 350;

            // 标题
            var lblTitle = new Label
            {
                Text = "激活软件",
                Location = new Point(20, y),
                Size = new Size(410, 30),
                Font = new Font("Microsoft Sans Serif", 12F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 215),
                TextAlign = ContentAlignment.MiddleCenter
            };
            this.Controls.Add(lblTitle);

            y += 50;

            // 授权状态
            var lblStatusTitle = new Label
            {
                Text = "授权状态:",
                Location = new Point(20, y),
                Size = new Size(labelWidth, 25),
                TextAlign = ContentAlignment.MiddleRight
            };
            lblStatusValue = new Label
            {
                Text = "试用版",
                Location = new Point(leftMargin, y),
                Size = new Size(100, 25),
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold),
                ForeColor = Color.Orange
            };
            lblRemainingDays = new Label
            {
                Text = "",
                Location = new Point(leftMargin + 100, y),
                Size = new Size(200, 25),
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 215)
            };
            this.Controls.Add(lblStatusTitle);
            this.Controls.Add(lblStatusValue);
            this.Controls.Add(lblRemainingDays);

            y += 35;

            // 用户姓名
            var lblUserTitle = new Label
            {
                Text = "用户姓名:",
                Location = new Point(20, y),
                Size = new Size(labelWidth, 25),
                TextAlign = ContentAlignment.MiddleRight
            };
            lblUserName = new Label
            {
                Text = "",
                Location = new Point(leftMargin, y),
                Size = new Size(contentWidth, 25)
            };
            this.Controls.Add(lblUserTitle);
            this.Controls.Add(lblUserName);

            y += 35;

            // 用户邮箱
            var lblEmailTitle = new Label
            {
                Text = "用户邮箱:",
                Location = new Point(20, y),
                Size = new Size(labelWidth, 25),
                TextAlign = ContentAlignment.MiddleRight
            };
            lblUserEmail = new Label
            {
                Text = "",
                Location = new Point(leftMargin, y),
                Size = new Size(contentWidth, 25)
            };
            this.Controls.Add(lblEmailTitle);
            this.Controls.Add(lblUserEmail);

            y += 35;

            // 组织/公司
            var lblOrgTitle = new Label
            {
                Text = "组织/公司:",
                Location = new Point(20, y),
                Size = new Size(labelWidth, 25),
                TextAlign = ContentAlignment.MiddleRight
            };
            lblOrg = new Label
            {
                Text = "",
                Location = new Point(leftMargin, y),
                Size = new Size(contentWidth, 25)
            };
            this.Controls.Add(lblOrgTitle);
            this.Controls.Add(lblOrg);

            y += 35;

            // 注册日期
            var lblDateTitle = new Label
            {
                Text = "注册日期:",
                Location = new Point(20, y),
                Size = new Size(labelWidth, 25),
                TextAlign = ContentAlignment.MiddleRight
            };
            lblDate = new Label
            {
                Text = "",
                Location = new Point(leftMargin, y),
                Size = new Size(contentWidth, 25)
            };
            this.Controls.Add(lblDateTitle);
            this.Controls.Add(lblDate);

            y += 35;

            // 激活日期
            var lblFirstRunTitle = new Label
            {
                Text = "激活日期:",
                Location = new Point(20, y),
                Size = new Size(labelWidth, 25),
                TextAlign = ContentAlignment.MiddleRight
            };
            lblFirstRun = new Label
            {
                Text = "-",
                Location = new Point(leftMargin, y),
                Size = new Size(contentWidth, 25)
            };
            this.Controls.Add(lblFirstRunTitle);
            this.Controls.Add(lblFirstRun);

            y += 35;

            // 到期日期
            var lblExpireTitle = new Label
            {
                Text = "到期日期:",
                Location = new Point(20, y),
                Size = new Size(labelWidth, 25),
                TextAlign = ContentAlignment.MiddleRight
            };
            lblExpire = new Label
            {
                Text = "",
                Location = new Point(leftMargin, y),
                Size = new Size(contentWidth, 25)
            };
            this.Controls.Add(lblExpireTitle);
            this.Controls.Add(lblExpire);

            y += 35;

            // MAC地址
            var lblMacTitle = new Label
            {
                Text = "MAC地址:",
                Location = new Point(20, y),
                Size = new Size(labelWidth, 25),
                TextAlign = ContentAlignment.MiddleRight
            };
            lblMacAddress = new Label
            {
                Text = "",
                Location = new Point(leftMargin, y),
                Size = new Size(contentWidth - 60, 25),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(5, 5, 5, 5)
            };
            chkShowMac = new CheckBox
            {
                Text = "显示",
                Location = new Point(leftMargin + contentWidth - 45, y),
                Size = new Size(50, 25)
            };
            chkShowMac.CheckedChanged += (s, e) =>
            {
                var settings = ConfigService.LoadAll();
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
                Location = new Point(20, y),
                Size = new Size(labelWidth, 25),
                TextAlign = ContentAlignment.MiddleRight
            };
            txtSerialNumber = new TextBox
            {
                Location = new Point(leftMargin, y),
                Size = new Size(250, 25)
            };
            this.Controls.Add(lblActivateTitle);
            this.Controls.Add(txtSerialNumber);

            y += 40;

            // 状态标签
            lblStatus = new Label
            {
                Text = "",
                Location = new Point(20, y),
                Size = new Size(510, 25),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleCenter
            };
            this.Controls.Add(lblStatus);

            y += 40;

            // 激活授权按钮
            var btnActivate = new Button
            {
                Text = "激活授权",
                Location = new Point(140, y),
                Size = new Size(120, 35),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Tag = "btnActivate"
            };
            btnActivate.Click += BtnActivate_Click;
            this.Controls.Add(btnActivate);

            // 关闭按钮
            btnCancel = new Button
            {
                Text = "关闭",
                Location = new Point(250, y),
                Size = new Size(120, 35),
                FlatStyle = FlatStyle.Flat
            };
            btnCancel.Click += (s, e) => { this.Close(); };
            this.Controls.Add(btnCancel);

            // 加载用户信息
            LoadUserInfo();
        }

        private void LoadUserInfo()
        {
            var settings = ConfigService.LoadAll();

            lblUserName.Text = settings.RegisteredUserName;
            lblUserEmail.Text = settings.RegisteredUserEmail;
            lblOrg.Text = settings.RegisteredOrganization ?? "-";
            lblDate.Text = settings.RegisterDate?.ToString("yyyy-MM-dd") ?? "-";
            lblFirstRun.Text = settings.FirstRunDate?.ToString("yyyy-MM-dd") ?? "-";
            lblExpire.Text = settings.RegisterExpireDate ?? "-";

            lblMacAddress.Text = MaskMac(settings.RegisteredMacAddress);

            if (settings.RegisterRemainingDays.HasValue)
            {
                lblRemainingDays.Text = $"剩余 {settings.RegisterRemainingDays.Value} 天";
            }

            txtSerialNumber.Text = settings.RegisterSerialNumber ?? "";

            // 更新授权状态显示
            lblStatusValue.Text = string.IsNullOrEmpty(settings.RegisterSerialNumber) ? "试用版" : "订阅版";
            lblStatusValue.ForeColor = string.IsNullOrEmpty(settings.RegisterSerialNumber) ? Color.Orange : Color.Green;
        }

        private async void BtnActivate_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtSerialNumber.Text))
            {
                lblStatus.Text = "请输入授权码";
                lblStatus.ForeColor = Color.Red;
                return;
            }

            var activationCode = txtSerialNumber.Text.Trim();
            btnCancel.Enabled = false;

            // 找到激活按钮并禁用
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is Button btn && btn.Tag?.ToString() == "btnActivate")
                {
                    btn.Enabled = false;
                    break;
                }
            }

            lblStatus.Text = "正在验证授权码...";
            lblStatus.ForeColor = Color.Gray;

            try
            {
                var settings = ConfigService.LoadAll();
                var macAddress = settings.RegisteredMacAddress ?? "";

                var result = await _registrationService.ActivateByCodeAsync(
                    activationCode,
                    macAddress,
                    settings.RegisteredUserName,
                    settings.RegisteredUserEmail,
                    settings.RegisterDate?.ToString("yyyy-MM-dd") ?? DateTime.Now.ToString("yyyy-MM-dd")
                );

                if (result.Success)
                {
                    // 激活成功，更新本地信息
                    settings.IsRegistered = true;
                    settings.RegisterSerialNumber = activationCode;
                    settings.RegisterExpireDate = result.ExpireDate;
                    settings.RegisterRemainingDays = result.RemainingDays;

                    if (DateTime.TryParse(result.InstallDate, out var installDate))
                        settings.RegisterDate = installDate;

                    ConfigService.SaveAll(settings);
                    RegistryService.SaveRegistration(settings);
                    RegistryService.SaveActivation(settings);

                    lblStatus.Text = $"激活成功！剩余 {result.RemainingDays} 天";
                    lblStatus.ForeColor = Color.Green;

                    // 更新授权状态显示
                    foreach (Control ctrl in this.Controls)
                    {
                        if (ctrl is Label lbl && lbl.Tag == null && (lbl.Text.StartsWith("试用版") || lbl.Text.StartsWith("订阅版")))
                        {
                            lbl.Text = "订阅版";
                            lbl.ForeColor = Color.Green;
                            break;
                        }
                    }

                    // 更新 MainForm 的显示
                    if (Application.OpenForms["MainForm"] is MainForm mainForm)
                    {
                        mainForm.UpdateRegistrationStatus();
                    }

                    await Task.Delay(1500);
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    lblStatus.Text = result.Message;
                    lblStatus.ForeColor = Color.Red;

                    // 重新启用按钮
                    foreach (Control ctrl in this.Controls)
                    {
                        if (ctrl is Button btn && btn.Tag?.ToString() == "btnActivate")
                        {
                            btn.Enabled = true;
                            break;
                        }
                    }
                    btnCancel.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Error(ex, "激活异常");
                lblStatus.Text = "激活异常: " + ex.Message;
                lblStatus.ForeColor = Color.Red;

                // 重新启用按钮
                foreach (Control ctrl in this.Controls)
                {
                    if (ctrl is Button btn && btn.Tag?.ToString() == "btnActivate")
                    {
                        btn.Enabled = true;
                        break;
                    }
                }
                btnCancel.Enabled = true;
            }
        }

        private string MaskMac(string mac)
        {
            if (string.IsNullOrEmpty(mac)) return "";
            if (mac.Length < 12) return mac;
            return mac.Substring(0, 8) + "-****-****";
        }
    }
}
