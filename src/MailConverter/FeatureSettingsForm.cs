using System;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;

namespace MailConverter
{
    public class FeatureSettingsForm : Form
    {
        private FeatureSettings _settings;
        private CheckBox _chkToPst, _chkExtract, _chkSingleUserSync, _chkBatchSync, _chkO365Toolkit, _chkOnPremiseToolkit;
        private CheckBox _chkToPstEml, _chkToPstOst, _chkToPstImap, _chkToPstMultiImap;
        private CheckBox _chkExtractImap, _chkExtractFiles;
        private CheckBox _chkSingleUserSyncEmlImport, _chkSingleUserSyncContacts;
        private CheckBox _chkBatchSyncLogin, _chkBatchSyncPstMail, _chkBatchSyncPstContacts, _chkBatchSyncPstCalendar;
        private CheckBox _chkBatchSyncCsvContacts, _chkBatchSyncVcfContacts, _chkBatchSyncCsvCalendar, _chkBatchSyncPurview;
        private CheckBox _chkO365ToolkitLogin, _chkO365ToolkitAccount, _chkO365ToolkitGroup;
        private CheckBox _chkO365ToolkitMobile, _chkO365ToolkitTraffic, _chkO365ToolkitMigration;
        private CheckBox _chkO365ToolkitWhois, _chkO365ToolkitDns, _chkO365ToolkitMailSearch;
        private Button _btnOk, _btnSelectAll, _btnDeselectAll;
        private Panel _contentPanel;
        private Label _lblFontSettings;
        private ComboBox _cmbLogFontName;
        private NumericUpDown _numLogFontSize;
        private ComboBox _cmbStatusFontName;
        private NumericUpDown _numStatusFontSize;
        private bool _isInitializing = false;

        // Exchange On-Premise 管理员设置
        private TextBox _txtOnPremiseAdminEmail;
        private TextBox _txtOnPremisePassword;
        private TextBox _txtOnPremiseEwsUrl;
        private TextBox _txtOnPremiseDomain;
        private Label _lblOnPremiseSettings;

        /// <summary>
        /// 设置变更时触发的事件（实时更新界面）
        /// </summary>
        public event Action<FeatureSettings> SettingsChanged;

        public FeatureSettingsForm(FeatureSettings currentSettings)
        {
            _settings = currentSettings.Clone();
            InitializeComponent();
            LoadSettings();
        }

        private void InitializeComponent()
        {
            this.Text = "自定义功能与设置";
            this.Size = new Size(520, 600);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;

            // 使用 TabControl
            var tabControl = new TabControl
            {
                Location = new Point(10, 10),
                Size = new Size(485, 500)
            };

            // Tab 1: 功能模块
            var tabFeatures = new TabPage("功能模块");
            tabFeatures.Padding = new Padding(10);

            _btnSelectAll = new Button { Text = "全选", Location = new Point(10, 10), Size = new Size(75, 25) };
            _btnDeselectAll = new Button { Text = "全取消", Location = new Point(90, 10), Size = new Size(75, 25) };
            _btnSelectAll.Click += (s, e) => SetAll(true);
            _btnDeselectAll.Click += (s, e) => SetAll(false);

            _contentPanel = new Panel
            {
                Location = new Point(10, 45),
                Size = new Size(455, 400),
                AutoScroll = true
            };

            tabFeatures.Controls.Add(_btnSelectAll);
            tabFeatures.Controls.Add(_btnDeselectAll);
            tabFeatures.Controls.Add(_contentPanel);

            int y = 10;

            // 转换为PST
            _chkToPst = CreateMainCheckBox("转换为PST", 0, y);
            _chkToPst.CheckedChanged += (s, e) => SetSubFeaturesEnabled("ToPst", _chkToPst.Checked);
            y += 30;

            y = AddSubFeatures(_contentPanel, y, new (CheckBox, string)[]
            {
                (_chkToPstEml = CreateCheckBox("  EML转PST", 20, y), "ToPst_Eml"),
                (_chkToPstOst = CreateCheckBox("  OST转PST", 20, y += 25), "ToPst_Ost"),
                (_chkToPstImap = CreateCheckBox("  IMAP收件转PST", 20, y += 25), "ToPst_Imap"),
                (_chkToPstMultiImap = CreateCheckBox("  IMAP多线程转PST", 20, y += 25), "ToPst_MultiImap")
            });
            y += 10;

            // 邮件提取
            _chkExtract = CreateMainCheckBox("邮件提取", 0, y);
            _chkExtract.CheckedChanged += (s, e) => SetSubFeaturesEnabled("Extract", _chkExtract.Checked);
            y += 30;

            y = AddSubFeatures(_contentPanel, y, new (CheckBox, string)[]
            {
                (_chkExtractImap = CreateCheckBox("  IMAP邮件提取EML(MSG)", 20, y), "Extract_Imap"),
                (_chkExtractFiles = CreateCheckBox("  提取邮件", 20, y += 25), "Extract_Files")
            });
            y += 10;

            // 单用户同步O365
            _chkSingleUserSync = CreateMainCheckBox("单用户同步O365", 0, y);
            _chkSingleUserSync.CheckedChanged += (s, e) => SetSubFeaturesEnabled("SingleUserSync", _chkSingleUserSync.Checked);
            y += 30;

            y = AddSubFeatures(_contentPanel, y, new (CheckBox, string)[]
            {
                (_chkSingleUserSyncEmlImport = CreateCheckBox("  EML导入", 20, y), "SingleUserSync_EmlImport"),
                (_chkSingleUserSyncContacts = CreateCheckBox("  同步联系人/日历", 20, y += 25), "SingleUserSync_Contacts")
            });
            y += 10;

            // 用户批量同步到O365
            _chkBatchSync = CreateMainCheckBox("用户批量同步到O365", 0, y);
            _chkBatchSync.CheckedChanged += (s, e) => SetSubFeaturesEnabled("BatchSync", _chkBatchSync.Checked);
            y += 30;

            y = AddSubFeatures(_contentPanel, y, new (CheckBox, string)[]
            {
                (_chkBatchSyncLogin = CreateCheckBox("  登录", 20, y), "BatchSync_Login"),
                (_chkBatchSyncPstMail = CreateCheckBox("  PST同步邮件", 20, y += 25), "BatchSync_PstMail"),
                (_chkBatchSyncPstContacts = CreateCheckBox("  PST同步联系人", 20, y += 25), "BatchSync_PstContacts"),
                (_chkBatchSyncPstCalendar = CreateCheckBox("  PST同步日历", 20, y += 25), "BatchSync_PstCalendar"),
                (_chkBatchSyncCsvContacts = CreateCheckBox("  CSV同步联系人", 20, y += 25), "BatchSync_CsvContacts"),
                (_chkBatchSyncVcfContacts = CreateCheckBox("  VCF文件夹联系人同步", 20, y += 25), "BatchSync_VcfContacts"),
                (_chkBatchSyncCsvCalendar = CreateCheckBox("  文件夹日历同步", 20, y += 25), "BatchSync_CsvCalendar"),
                (_chkBatchSyncPurview = CreateCheckBox("  purView方案批量同步", 20, y += 25), "BatchSync_Purview")
            });
            y += 10;

            // Exchange Online 百宝箱
            _chkO365Toolkit = CreateMainCheckBox("Exchange Online 百宝箱", 0, y);
            _chkO365Toolkit.CheckedChanged += (s, e) => SetSubFeaturesEnabled("O365Toolkit", _chkO365Toolkit.Checked);
            y += 30;

            y = AddSubFeatures(_contentPanel, y, new (CheckBox, string)[]
            {
                (_chkO365ToolkitLogin = CreateCheckBox("  登录", 20, y), "O365Toolkit_Login"),
                (_chkO365ToolkitAccount = CreateCheckBox("  账户管理", 20, y += 25), "O365Toolkit_Account"),
                (_chkO365ToolkitGroup = CreateCheckBox("  组管理", 20, y += 25), "O365Toolkit_Group"),
                (_chkO365ToolkitMobile = CreateCheckBox("  移动设备管理", 20, y += 25), "O365Toolkit_Mobile"),
                (_chkO365ToolkitTraffic = CreateCheckBox("  邮件流量搜索", 20, y += 25), "O365Toolkit_Traffic"),
                (_chkO365ToolkitMigration = CreateCheckBox("  邮件迁移", 20, y += 25), "O365Toolkit_Migration"),
                (_chkO365ToolkitWhois = CreateCheckBox("  WHOIS查询", 20, y += 25), "O365Toolkit_Whois"),
                (_chkO365ToolkitDns = CreateCheckBox("  DNS查询", 20, y += 25), "O365Toolkit_Dns"),
                (_chkO365ToolkitMailSearch = CreateCheckBox("  邮件搜索导出", 20, y += 25), "O365Toolkit_MailSearch")
            });
            y += 10;

            // Exchange On-Premise 百宝箱
            _chkOnPremiseToolkit = CreateMainCheckBox("Exchange On-Premise 百宝箱", 0, y);

            // Tab 2: Exchange On-Premise 设置
            var tabOnPremise = new TabPage("Exchange On-Premise");
            tabOnPremise.Padding = new Padding(15);

            var lblOnPremiseTitle = new Label
            {
                Text = "管理员默认凭据",
                Location = new Point(15, 15),
                Size = new Size(200, 25),
                Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold)
            };
            tabOnPremise.Controls.Add(lblOnPremiseTitle);

            int y2 = 50;

            // 管理员邮箱
            var lblEmail = new Label { Text = "管理员邮箱:", Location = new Point(15, y2), Size = new Size(100, 25), TextAlign = ContentAlignment.MiddleRight };
            _txtOnPremiseAdminEmail = new TextBox { Location = new Point(120, y2), Size = new Size(300, 25) };
            tabOnPremise.Controls.Add(lblEmail);
            tabOnPremise.Controls.Add(_txtOnPremiseAdminEmail);
            y2 += 35;

            // 密码
            var lblPassword = new Label { Text = "密码:", Location = new Point(15, y2), Size = new Size(100, 25), TextAlign = ContentAlignment.MiddleRight };
            _txtOnPremisePassword = new TextBox { Location = new Point(120, y2), Size = new Size(300, 25), UseSystemPasswordChar = true };
            tabOnPremise.Controls.Add(lblPassword);
            tabOnPremise.Controls.Add(_txtOnPremisePassword);
            y2 += 35;

            // EWS地址
            var lblEwsUrl = new Label { Text = "EWS地址:", Location = new Point(15, y2), Size = new Size(100, 25), TextAlign = ContentAlignment.MiddleRight };
            _txtOnPremiseEwsUrl = new TextBox { Location = new Point(120, y2), Size = new Size(300, 25) };
            tabOnPremise.Controls.Add(lblEwsUrl);
            tabOnPremise.Controls.Add(_txtOnPremiseEwsUrl);
            y2 += 35;

            // 域
            var lblDomain = new Label { Text = "域:", Location = new Point(15, y2), Size = new Size(100, 25), TextAlign = ContentAlignment.MiddleRight };
            _txtOnPremiseDomain = new TextBox { Location = new Point(120, y2), Size = new Size(150, 25) };
            tabOnPremise.Controls.Add(lblDomain);
            tabOnPremise.Controls.Add(_txtOnPremiseDomain);

            // Tab 3: 界面字体
            var tabFont = new TabPage("界面字体");
            tabFont.Padding = new Padding(15);

            var lblFontTitle = new Label
            {
                Text = "字体设置",
                Location = new Point(15, 15),
                Size = new Size(200, 25),
                Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold)
            };
            tabFont.Controls.Add(lblFontTitle);

            int y3 = 50;

            // 日志窗口字体
            var lblLogFont = new Label { Text = "日志窗口字体:", Location = new Point(15, y3), Size = new Size(100, 25), TextAlign = ContentAlignment.MiddleRight };
            _cmbLogFontName = new ComboBox { Location = new Point(120, y3), Size = new Size(150, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbLogFontName.Items.AddRange(new[] { "Consolas", "Courier New", "Fixedsys", "Lucida Console", "宋体", "微软雅黑" });
            _cmbLogFontName.SelectedIndex = 0;
            _numLogFontSize = new NumericUpDown { Location = new Point(280, y3), Size = new Size(50, 25), Minimum = 8, Maximum = 20, Value = 10 };
            var lblLogPt = new Label { Text = "pt", Location = new Point(333, y3), Size = new Size(30, 25) };
            tabFont.Controls.Add(lblLogFont);
            tabFont.Controls.Add(_cmbLogFontName);
            tabFont.Controls.Add(_numLogFontSize);
            tabFont.Controls.Add(lblLogPt);
            y3 += 35;

            // 状态栏字体
            var lblStatusFont = new Label { Text = "状态栏字体:", Location = new Point(15, y3), Size = new Size(100, 25), TextAlign = ContentAlignment.MiddleRight };
            _cmbStatusFontName = new ComboBox { Location = new Point(120, y3), Size = new Size(150, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbStatusFontName.Items.AddRange(new[] { "Microsoft Sans Serif", "Segoe UI", "宋体", "微软雅黑", "Tahoma" });
            _cmbStatusFontName.SelectedIndex = 0;
            _numStatusFontSize = new NumericUpDown { Location = new Point(280, y3), Size = new Size(50, 25), Minimum = 8, Maximum = 16, Value = 10 };
            var lblStatusPt = new Label { Text = "pt", Location = new Point(333, y3), Size = new Size(30, 25) };
            tabFont.Controls.Add(lblStatusFont);
            tabFont.Controls.Add(_cmbStatusFontName);
            tabFont.Controls.Add(_numStatusFontSize);
            tabFont.Controls.Add(lblStatusPt);

            tabControl.TabPages.Add(tabFeatures);
            tabControl.TabPages.Add(tabOnPremise);
            tabControl.TabPages.Add(tabFont);

            // 按钮
            _btnOk = new Button { Text = "确定", Location = new Point(380, 520), Size = new Size(90, 30) };
            _btnOk.Click += BtnOk_Click;

            this.Controls.Add(tabControl);
            this.Controls.Add(_btnOk);

            // 为所有复选框附加实时变更事件
            AttachCheckboxEvents();
        }

        /// <summary>
        /// 为所有复选框附加实时变更事件，实时通知设置变化
        /// </summary>
        private void AttachCheckboxEvents()
        {
            foreach (Control ctrl in _contentPanel.Controls)
            {
                if (ctrl is CheckBox chk)
                {
                    chk.CheckedChanged += (s, e) =>
                    {
                        if (_isInitializing) return;
                        ReadCurrentSettings();
                        SettingsChanged?.Invoke(_settings);
                    };
                }
            }
        }

        /// <summary>
        /// 从UI读取当前设置
        /// </summary>
        private void ReadCurrentSettings()
        {
            _settings.Feature_ToPst = _chkToPst.Checked;
            _settings.Feature_ToPst_Eml = _chkToPstEml.Checked;
            _settings.Feature_ToPst_Ost = _chkToPstOst.Checked;
            _settings.Feature_ToPst_Imap = _chkToPstImap.Checked;
            _settings.Feature_ToPst_MultiImap = _chkToPstMultiImap.Checked;

            _settings.Feature_Extract = _chkExtract.Checked;
            _settings.Feature_Extract_Imap = _chkExtractImap.Checked;
            _settings.Feature_Extract_Files = _chkExtractFiles.Checked;

            _settings.Feature_SingleUserSync = _chkSingleUserSync.Checked;
            _settings.Feature_SingleUserSync_EmlImport = _chkSingleUserSyncEmlImport.Checked;
            _settings.Feature_SingleUserSync_Contacts = _chkSingleUserSyncContacts.Checked;

            _settings.Feature_BatchSync = _chkBatchSync.Checked;
            _settings.Feature_BatchSync_Login = _chkBatchSyncLogin.Checked;
            _settings.Feature_BatchSync_PstMail = _chkBatchSyncPstMail.Checked;
            _settings.Feature_BatchSync_PstContacts = _chkBatchSyncPstContacts.Checked;
            _settings.Feature_BatchSync_PstCalendar = _chkBatchSyncPstCalendar.Checked;
            _settings.Feature_BatchSync_CsvContacts = _chkBatchSyncCsvContacts.Checked;
            _settings.Feature_BatchSync_VcfContacts = _chkBatchSyncVcfContacts.Checked;
            _settings.Feature_BatchSync_CsvCalendar = _chkBatchSyncCsvCalendar.Checked;
            _settings.Feature_BatchSync_Purview = _chkBatchSyncPurview.Checked;

            _settings.Feature_O365Toolkit = _chkO365Toolkit.Checked;
            _settings.Feature_O365Toolkit_Login = _chkO365ToolkitLogin.Checked;
            _settings.Feature_O365Toolkit_Account = _chkO365ToolkitAccount.Checked;
            _settings.Feature_O365Toolkit_Group = _chkO365ToolkitGroup.Checked;
            _settings.Feature_O365Toolkit_Mobile = _chkO365ToolkitMobile.Checked;
            _settings.Feature_O365Toolkit_Traffic = _chkO365ToolkitTraffic.Checked;
            _settings.Feature_O365Toolkit_Migration = _chkO365ToolkitMigration.Checked;
            _settings.Feature_O365Toolkit_Whois = _chkO365ToolkitWhois.Checked;
            _settings.Feature_O365Toolkit_Dns = _chkO365ToolkitDns.Checked;
            _settings.Feature_O365Toolkit_MailSearch = _chkO365ToolkitMailSearch.Checked;

            _settings.Feature_OnPremiseToolkit = _chkOnPremiseToolkit.Checked;

            // 读取字体设置
            _settings.LogFontName = _cmbLogFontName.SelectedItem?.ToString() ?? "Consolas";
            _settings.LogFontSize = (float)_numLogFontSize.Value;
            _settings.StatusFontName = _cmbStatusFontName.SelectedItem?.ToString() ?? "Microsoft Sans Serif";
            _settings.StatusFontSize = (float)_numStatusFontSize.Value;
        }

        private CheckBox CreateMainCheckBox(string text, int x, int y)
        {
            var chk = new CheckBox
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(440, 25),
                Font = new Font("Microsoft YaHei UI", 9, FontStyle.Bold)
            };
            _contentPanel.Controls.Add(chk);
            return chk;
        }

        private CheckBox CreateCheckBox(string text, int x, int y)
        {
            var chk = new CheckBox
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(420, 22)
            };
            _contentPanel.Controls.Add(chk);
            return chk;
        }

        private int AddSubFeatures(Panel panel, int y, (CheckBox chk, string key)[] features)
        {
            foreach (var (chk, _) in features)
            {
                panel.Controls.Add(chk);
            }
            return y + 25 * features.Length;
        }

        private void SetSubFeaturesEnabled(string category, bool enabled)
        {
            foreach (Control ctrl in _contentPanel.Controls)
            {
                if (ctrl is CheckBox chk && chk.Text.StartsWith("  "))
                {
                    var key = GetSubFeatureKey(chk);
                    if (key.StartsWith(category))
                    {
                        chk.Enabled = enabled;
                        if (!enabled) chk.Checked = false;
                    }
                }
            }
        }

        private string GetSubFeatureKey(CheckBox chk)
        {
            if (chk == _chkToPstEml) return "ToPst_Eml";
            if (chk == _chkToPstOst) return "ToPst_Ost";
            if (chk == _chkToPstImap) return "ToPst_Imap";
            if (chk == _chkToPstMultiImap) return "ToPst_MultiImap";
            if (chk == _chkExtractImap) return "Extract_Imap";
            if (chk == _chkExtractFiles) return "Extract_Files";
            if (chk == _chkSingleUserSyncEmlImport) return "SingleUserSync_EmlImport";
            if (chk == _chkSingleUserSyncContacts) return "SingleUserSync_Contacts";
            if (chk == _chkBatchSyncLogin) return "BatchSync_Login";
            if (chk == _chkBatchSyncPstMail) return "BatchSync_PstMail";
            if (chk == _chkBatchSyncPstContacts) return "BatchSync_PstContacts";
            if (chk == _chkBatchSyncPstCalendar) return "BatchSync_PstCalendar";
            if (chk == _chkBatchSyncCsvContacts) return "BatchSync_CsvContacts";
            if (chk == _chkBatchSyncVcfContacts) return "BatchSync_VcfContacts";
            if (chk == _chkBatchSyncCsvCalendar) return "BatchSync_CsvCalendar";
            if (chk == _chkBatchSyncPurview) return "BatchSync_Purview";
            if (chk == _chkO365ToolkitLogin) return "O365Toolkit_Login";
            if (chk == _chkO365ToolkitAccount) return "O365Toolkit_Account";
            if (chk == _chkO365ToolkitGroup) return "O365Toolkit_Group";
            if (chk == _chkO365ToolkitMobile) return "O365Toolkit_Mobile";
            if (chk == _chkO365ToolkitTraffic) return "O365Toolkit_Traffic";
            if (chk == _chkO365ToolkitMigration) return "O365Toolkit_Migration";
            if (chk == _chkO365ToolkitWhois) return "O365Toolkit_Whois";
            if (chk == _chkO365ToolkitDns) return "O365Toolkit_Dns";
            if (chk == _chkO365ToolkitMailSearch) return "O365Toolkit_MailSearch";
            return "";
        }

        private void LoadSettings()
        {
            // 临时设置一个标志，让 ReadCurrentSettings 不要在初始化期间覆盖 _settings
            _isInitializing = true;

            _chkToPst.Checked = _settings.Feature_ToPst;
            _chkToPstEml.Checked = _settings.Feature_ToPst_Eml;
            _chkToPstOst.Checked = _settings.Feature_ToPst_Ost;
            _chkToPstImap.Checked = _settings.Feature_ToPst_Imap;
            _chkToPstMultiImap.Checked = _settings.Feature_ToPst_MultiImap;

            _chkExtract.Checked = _settings.Feature_Extract;
            _chkExtractImap.Checked = _settings.Feature_Extract_Imap;
            _chkExtractFiles.Checked = _settings.Feature_Extract_Files;

            _chkSingleUserSync.Checked = _settings.Feature_SingleUserSync;
            _chkSingleUserSyncEmlImport.Checked = _settings.Feature_SingleUserSync_EmlImport;
            _chkSingleUserSyncContacts.Checked = _settings.Feature_SingleUserSync_Contacts;

            _chkBatchSync.Checked = _settings.Feature_BatchSync;
            _chkBatchSyncLogin.Checked = _settings.Feature_BatchSync_Login;
            _chkBatchSyncPstMail.Checked = _settings.Feature_BatchSync_PstMail;
            _chkBatchSyncPstContacts.Checked = _settings.Feature_BatchSync_PstContacts;
            _chkBatchSyncPstCalendar.Checked = _settings.Feature_BatchSync_PstCalendar;
            _chkBatchSyncCsvContacts.Checked = _settings.Feature_BatchSync_CsvContacts;
            _chkBatchSyncVcfContacts.Checked = _settings.Feature_BatchSync_VcfContacts;
            _chkBatchSyncCsvCalendar.Checked = _settings.Feature_BatchSync_CsvCalendar;
            _chkBatchSyncPurview.Checked = _settings.Feature_BatchSync_Purview;

            _chkO365Toolkit.Checked = _settings.Feature_O365Toolkit;
            _chkO365ToolkitLogin.Checked = _settings.Feature_O365Toolkit_Login;
            _chkO365ToolkitAccount.Checked = _settings.Feature_O365Toolkit_Account;
            _chkO365ToolkitGroup.Checked = _settings.Feature_O365Toolkit_Group;
            _chkO365ToolkitMobile.Checked = _settings.Feature_O365Toolkit_Mobile;
            _chkO365ToolkitTraffic.Checked = _settings.Feature_O365Toolkit_Traffic;
            _chkO365ToolkitMigration.Checked = _settings.Feature_O365Toolkit_Migration;
            _chkO365ToolkitWhois.Checked = _settings.Feature_O365Toolkit_Whois;
            _chkO365ToolkitDns.Checked = _settings.Feature_O365Toolkit_Dns;
            _chkO365ToolkitMailSearch.Checked = _settings.Feature_O365Toolkit_MailSearch;

            _chkOnPremiseToolkit.Checked = _settings.Feature_OnPremiseToolkit;

            // 加载字体设置
            SetComboBoxByText(_cmbLogFontName, _settings.LogFontName);
            _numLogFontSize.Value = (decimal)Math.Max((double)_numLogFontSize.Minimum, Math.Min((double)_numLogFontSize.Maximum, _settings.LogFontSize));
            SetComboBoxByText(_cmbStatusFontName, _settings.StatusFontName);
            _numStatusFontSize.Value = (decimal)Math.Max((double)_numStatusFontSize.Minimum, Math.Min((double)_numStatusFontSize.Maximum, _settings.StatusFontSize));

            // 加载 Exchange On-Premise 设置
            _txtOnPremiseAdminEmail.Text = _settings.OnPremise_AdminEmail;
            _txtOnPremisePassword.Text = _settings.OnPremise_Password;
            _txtOnPremiseEwsUrl.Text = _settings.OnPremise_EwsUrl;
            _txtOnPremiseDomain.Text = _settings.OnPremise_Domain;

            // 初始化完成，允许 ReadCurrentSettings 正常工作
            _isInitializing = false;

            UpdateSubFeaturesEnabled();
        }

        private void UpdateSubFeaturesEnabled()
        {
            SetSubFeaturesEnabled("ToPst", _chkToPst.Checked);
            SetSubFeaturesEnabled("Extract", _chkExtract.Checked);
            SetSubFeaturesEnabled("SingleUserSync", _chkSingleUserSync.Checked);
            SetSubFeaturesEnabled("BatchSync", _chkBatchSync.Checked);
            SetSubFeaturesEnabled("O365Toolkit", _chkO365Toolkit.Checked);
        }

        private void SetComboBoxByText(ComboBox cmb, string text)
        {
            for (int i = 0; i < cmb.Items.Count; i++)
            {
                if (cmb.Items[i].ToString() == text)
                {
                    cmb.SelectedIndex = i;
                    return;
                }
            }
            cmb.SelectedIndex = 0;
        }

        private void SetAll(bool value)
        {
            _chkToPst.Checked = value;
            _chkToPstEml.Checked = value;
            _chkToPstOst.Checked = value;
            _chkToPstImap.Checked = value;
            _chkToPstMultiImap.Checked = value;

            _chkExtract.Checked = value;
            _chkExtractImap.Checked = value;
            _chkExtractFiles.Checked = value;

            _chkSingleUserSync.Checked = value;
            _chkSingleUserSyncEmlImport.Checked = value;
            _chkSingleUserSyncContacts.Checked = value;

            _chkBatchSync.Checked = value;
            _chkBatchSyncLogin.Checked = value;
            _chkBatchSyncPstMail.Checked = value;
            _chkBatchSyncPstContacts.Checked = value;
            _chkBatchSyncPstCalendar.Checked = value;
            _chkBatchSyncCsvContacts.Checked = value;
            _chkBatchSyncVcfContacts.Checked = value;
            _chkBatchSyncCsvCalendar.Checked = value;
            _chkBatchSyncPurview.Checked = value;

            _chkO365Toolkit.Checked = value;
            _chkO365ToolkitLogin.Checked = value;
            _chkO365ToolkitAccount.Checked = value;
            _chkO365ToolkitGroup.Checked = value;
            _chkO365ToolkitMobile.Checked = value;
            _chkO365ToolkitTraffic.Checked = value;
            _chkO365ToolkitMigration.Checked = value;
            _chkO365ToolkitWhois.Checked = value;
            _chkO365ToolkitDns.Checked = value;
            _chkO365ToolkitMailSearch.Checked = value;

            _chkOnPremiseToolkit.Checked = value;
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            _settings.Feature_ToPst = _chkToPst.Checked;
            _settings.Feature_ToPst_Eml = _chkToPstEml.Checked;
            _settings.Feature_ToPst_Ost = _chkToPstOst.Checked;
            _settings.Feature_ToPst_Imap = _chkToPstImap.Checked;
            _settings.Feature_ToPst_MultiImap = _chkToPstMultiImap.Checked;

            _settings.Feature_Extract = _chkExtract.Checked;
            _settings.Feature_Extract_Imap = _chkExtractImap.Checked;
            _settings.Feature_Extract_Files = _chkExtractFiles.Checked;

            _settings.Feature_SingleUserSync = _chkSingleUserSync.Checked;
            _settings.Feature_SingleUserSync_EmlImport = _chkSingleUserSyncEmlImport.Checked;
            _settings.Feature_SingleUserSync_Contacts = _chkSingleUserSyncContacts.Checked;

            _settings.Feature_BatchSync = _chkBatchSync.Checked;
            _settings.Feature_BatchSync_Login = _chkBatchSyncLogin.Checked;
            _settings.Feature_BatchSync_PstMail = _chkBatchSyncPstMail.Checked;
            _settings.Feature_BatchSync_PstContacts = _chkBatchSyncPstContacts.Checked;
            _settings.Feature_BatchSync_PstCalendar = _chkBatchSyncPstCalendar.Checked;
            _settings.Feature_BatchSync_CsvContacts = _chkBatchSyncCsvContacts.Checked;
            _settings.Feature_BatchSync_VcfContacts = _chkBatchSyncVcfContacts.Checked;
            _settings.Feature_BatchSync_CsvCalendar = _chkBatchSyncCsvCalendar.Checked;
            _settings.Feature_BatchSync_Purview = _chkBatchSyncPurview.Checked;

            _settings.Feature_O365Toolkit = _chkO365Toolkit.Checked;
            _settings.Feature_O365Toolkit_Login = _chkO365ToolkitLogin.Checked;
            _settings.Feature_O365Toolkit_Account = _chkO365ToolkitAccount.Checked;
            _settings.Feature_O365Toolkit_Group = _chkO365ToolkitGroup.Checked;
            _settings.Feature_O365Toolkit_Mobile = _chkO365ToolkitMobile.Checked;
            _settings.Feature_O365Toolkit_Traffic = _chkO365ToolkitTraffic.Checked;
            _settings.Feature_O365Toolkit_Migration = _chkO365ToolkitMigration.Checked;
            _settings.Feature_O365Toolkit_Whois = _chkO365ToolkitWhois.Checked;
            _settings.Feature_O365Toolkit_Dns = _chkO365ToolkitDns.Checked;

            _settings.Feature_OnPremiseToolkit = _chkOnPremiseToolkit.Checked;

            // 保存字体设置
            _settings.LogFontName = _cmbLogFontName.SelectedItem?.ToString() ?? "Consolas";
            _settings.LogFontSize = (float)_numLogFontSize.Value;
            _settings.StatusFontName = _cmbStatusFontName.SelectedItem?.ToString() ?? "Microsoft Sans Serif";
            _settings.StatusFontSize = (float)_numStatusFontSize.Value;

            // 保存 Exchange On-Premise 设置
            _settings.OnPremise_AdminEmail = _txtOnPremiseAdminEmail.Text;
            _settings.OnPremise_Password = _txtOnPremisePassword.Text;
            _settings.OnPremise_EwsUrl = _txtOnPremiseEwsUrl.Text;
            _settings.OnPremise_Domain = _txtOnPremiseDomain.Text;

            FeatureSettingsService.Save(_settings);
            this.Tag = _settings;
            this.DialogResult = DialogResult.OK;
        }

        public static FeatureSettings ShowDialog(FeatureSettings currentSettings)
        {
            using (var form = new FeatureSettingsForm(currentSettings))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    return form.Tag as FeatureSettings;
                }
                return null;
            }
        }
    }
}
