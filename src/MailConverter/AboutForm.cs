using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace MailConverter
{
    public class AboutForm : Form
    {
        private Panel mainPanel;
        private PictureBox picLogo;
        private Label lblProductName;
        private Label lblVersion;
        private Label lblCopyright;
        private Label lblAllRights;
        private TextBox txtFeatures;
        private Button btnOK;
        private LinkLabel lnkPrivacy;
        private LinkLabel lnkLicense;
        private Panel separator1;
        private Panel separator2;

        public AboutForm()
        {
            InitializeComponent();
            LoadVersionInfo();
        }

        private void InitializeComponent()
        {
            this.Size = new Size(480, 380);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Text = "关于 小铭邮件百宝箱";
            this.BackColor = Color.FromArgb(245, 245, 245);

            // Main container with shadow effect
            mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(24)
            };

            // Product logo area
            picLogo = new PictureBox
            {
                Size = new Size(100, 100),
                Location = new Point(24, 24),
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.None
            };

            try
            {
                var logoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "guanyu.png");
                if (File.Exists(logoPath))
                    picLogo.Image = Image.FromFile(logoPath);
            }
            catch { }

            // If no image loaded, draw a placeholder envelope icon
            if (picLogo.Image == null)
            {
                var bmp = new Bitmap(100, 100);
                using (var g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.Clear(Color.FromArgb(0, 120, 212));
                    using (var pen = new Pen(Color.White, 3))
                    {
                        g.DrawRectangle(pen, 10, 25, 80, 55);
                        g.DrawLine(pen, 10, 25, 50, 52);
                        g.DrawLine(pen, 90, 25, 50, 52);
                    }
                    using (var font = new Font("Segoe UI", 18, FontStyle.Bold))
                    using (var brush = new SolidBrush(Color.White))
                    {
                        g.DrawString("M", font, brush, 35, 30);
                    }
                }
                picLogo.Image = bmp;
            }

            // Product info area
            lblProductName = new Label
            {
                Text = "小铭邮件百宝箱",
                Font = new Font("Segoe UI", 20, FontStyle.Bold),
                ForeColor = Color.FromArgb(30, 30, 30),
                Location = new Point(140, 28),
                AutoSize = true
            };

            lblVersion = new Label
            {
                Text = "版本 1.2.0",
                Font = new Font("Segoe UI", 10),
                ForeColor = Color.FromArgb(100, 100, 100),
                Location = new Point(140, 60),
                AutoSize = true
            };

            separator1 = new Panel
            {
                Location = new Point(24, 140),
                Size = new Size(432, 1),
                BackColor = Color.FromArgb(200, 200, 200)
            };

            // Features section
            lblCopyright = new Label
            {
                Text = "版权和归属",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.FromArgb(30, 30, 30),
                Location = new Point(24, 156),
                AutoSize = true
            };

            txtFeatures = new TextBox
            {
                Location = new Point(24, 184),
                Size = new Size(432, 80),
                Multiline = true,
                ReadOnly = true,
                BackColor = Color.FromArgb(250, 250, 250),
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.FromArgb(60, 60, 60),
                ScrollBars = ScrollBars.Vertical
            };
            txtFeatures.Text = @"© 2026 小铭科技 版权所有。保留所有权利。
小铭邮件百宝箱是用于邮件转换、Office 365同步等企业级应用工具。";

            separator2 = new Panel
            {
                Location = new Point(24, 276),
                Size = new Size(432, 1),
                BackColor = Color.FromArgb(200, 200, 200)
            };

            // Legal links
            lnkPrivacy = new LinkLabel
            {
                Text = "隐私声明",
                Location = new Point(24, 292),
                AutoSize = true,
                LinkColor = Color.FromArgb(0, 120, 212),
                ActiveLinkColor = Color.FromArgb(0, 100, 180),
                Font = new Font("Segoe UI", 9)
            };
            lnkPrivacy.LinkClicked += (s, e) => System.Diagnostics.Process.Start("https://www.booming.one/privacy");

            lnkLicense = new LinkLabel
            {
                Text = "  |  许可协议",
                Location = new Point(110, 292),
                AutoSize = true,
                LinkColor = Color.FromArgb(0, 120, 212),
                ActiveLinkColor = Color.FromArgb(0, 100, 180),
                Font = new Font("Segoe UI", 9)
            };
            lnkLicense.LinkClicked += (s, e) => System.Diagnostics.Process.Start("https://www.booming.one/license");

            lblAllRights = new Label
            {
                Text = "保留一切权利",
                Location = new Point(24, 314),
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.FromArgb(130, 130, 130),
                AutoSize = true
            };

            // OK Button - Microsoft style
            btnOK = new Button
            {
                Text = "确定",
                Size = new Size(90, 28),
                Location = new Point(366, 310),
                FlatStyle = FlatStyle.Standard,
                BackColor = Color.FromArgb(242, 242, 242),
                ForeColor = Color.FromArgb(30, 30, 30),
                Font = new Font("Segoe UI", 9),
                DialogResult = DialogResult.OK
            };
            btnOK.Click += (s, e) => this.Close();

            // Add controls to main panel first
            mainPanel.Controls.Add(picLogo);
            mainPanel.Controls.Add(lblProductName);
            mainPanel.Controls.Add(lblVersion);
            mainPanel.Controls.Add(separator1);
            mainPanel.Controls.Add(lblCopyright);
            mainPanel.Controls.Add(txtFeatures);
            mainPanel.Controls.Add(separator2);
            mainPanel.Controls.Add(lnkPrivacy);
            mainPanel.Controls.Add(lnkLicense);
            mainPanel.Controls.Add(lblAllRights);
            mainPanel.Controls.Add(btnOK);

            this.Controls.Add(mainPanel);
        }

        private void LoadVersionInfo()
        {
            try
            {
                var version = Assembly.GetExecutingAssembly().GetName().Version;
                lblVersion.Text = $"版本 {version.Major}.{version.Minor}.{version.Build}";
            }
            catch { }
        }
    }
}