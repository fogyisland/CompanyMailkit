using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace MailConverter
{
    /// <summary>
    /// Exchange On-Premise 配置账户
    /// </summary>
    public class OnPremiseAccount
    {
        public string Name { get; set; } = "";
        public string AdminEmail { get; set; } = "";
        public string Password { get; set; } = "";
        public string EwsUrl { get; set; } = "";
        public string Domain { get; set; } = "";
    }

    /// <summary>
    /// 功能设置服务 - 控制各功能模块的显示/隐藏
    /// </summary>
    public class FeatureSettingsService
    {
        private static readonly string SettingsDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config");
        private static readonly string FeatureSettingsFile = Path.Combine(SettingsDir, "features.inf");

        // 默认所有功能开启
        private static FeatureSettings _defaultSettings;

        static FeatureSettingsService()
        {
            Directory.CreateDirectory(SettingsDir);
            _defaultSettings = new FeatureSettings();
        }

        /// <summary>
        /// 加载功能设置
        /// </summary>
        public static FeatureSettings Load()
        {
            var settings = _defaultSettings.Clone();

            if (!File.Exists(FeatureSettingsFile))
                return settings;

            try
            {
                var lines = File.ReadAllLines(FeatureSettingsFile);
                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line) || !line.Contains("="))
                        continue;

                    var parts = line.Split(new[] { '=' }, 2);
                    var key = parts[0].Trim();
                    var value = parts[1].Trim().ToLower();

                    bool enabled = value == "true" || value == "1" || value == "yes";

                    switch (key)
                    {
                        case "Feature_ToPst": settings.Feature_ToPst = enabled; break;
                        case "Feature_ToPst_Eml": settings.Feature_ToPst_Eml = enabled; break;
                        case "Feature_ToPst_Ost": settings.Feature_ToPst_Ost = enabled; break;
                        case "Feature_ToPst_Imap": settings.Feature_ToPst_Imap = enabled; break;
                        case "Feature_ToPst_MultiImap": settings.Feature_ToPst_MultiImap = enabled; break;
                        case "Feature_Extract": settings.Feature_Extract = enabled; break;
                        case "Feature_Extract_Imap": settings.Feature_Extract_Imap = enabled; break;
                        case "Feature_Extract_Files": settings.Feature_Extract_Files = enabled; break;
                        case "Feature_SingleUserSync": settings.Feature_SingleUserSync = enabled; break;
                        case "Feature_SingleUserSync_EmlImport": settings.Feature_SingleUserSync_EmlImport = enabled; break;
                        case "Feature_SingleUserSync_Contacts": settings.Feature_SingleUserSync_Contacts = enabled; break;
                        case "Feature_BatchSync": settings.Feature_BatchSync = enabled; break;
                        case "Feature_BatchSync_Login": settings.Feature_BatchSync_Login = enabled; break;
                        case "Feature_BatchSync_PstMail": settings.Feature_BatchSync_PstMail = enabled; break;
                        case "Feature_BatchSync_PstContacts": settings.Feature_BatchSync_PstContacts = enabled; break;
                        case "Feature_BatchSync_PstCalendar": settings.Feature_BatchSync_PstCalendar = enabled; break;
                        case "Feature_BatchSync_CsvContacts": settings.Feature_BatchSync_CsvContacts = enabled; break;
                        case "Feature_BatchSync_VcfContacts": settings.Feature_BatchSync_VcfContacts = enabled; break;
                        case "Feature_BatchSync_CsvCalendar": settings.Feature_BatchSync_CsvCalendar = enabled; break;
                        case "Feature_BatchSync_Purview": settings.Feature_BatchSync_Purview = enabled; break;
                        case "Feature_O365Toolkit": settings.Feature_O365Toolkit = enabled; break;
                        case "Feature_O365Toolkit_Login": settings.Feature_O365Toolkit_Login = enabled; break;
                        case "Feature_O365Toolkit_Account": settings.Feature_O365Toolkit_Account = enabled; break;
                        case "Feature_O365Toolkit_Group": settings.Feature_O365Toolkit_Group = enabled; break;
                        case "Feature_O365Toolkit_Mobile": settings.Feature_O365Toolkit_Mobile = enabled; break;
                        case "Feature_O365Toolkit_Traffic": settings.Feature_O365Toolkit_Traffic = enabled; break;
                        case "Feature_O365Toolkit_Migration": settings.Feature_O365Toolkit_Migration = enabled; break;
                        case "Feature_O365Toolkit_Whois": settings.Feature_O365Toolkit_Whois = enabled; break;
                        case "Feature_O365Toolkit_Dns": settings.Feature_O365Toolkit_Dns = enabled; break;
                        case "Feature_O365Toolkit_MailSearch": settings.Feature_O365Toolkit_MailSearch = enabled; break;
                        case "Feature_OnPremiseToolkit": settings.Feature_OnPremiseToolkit = enabled; break;
                        case "Feature_Preferences": settings.Feature_Preferences = enabled; break;
                        case "LogFontName": settings.LogFontName = parts[1].Trim(); break;
                        case "LogFontSize":
                            if (float.TryParse(parts[1].Trim(), out float logSize))
                                settings.LogFontSize = logSize;
                            break;
                        case "StatusFontName": settings.StatusFontName = parts[1].Trim(); break;
                        case "StatusFontSize":
                            if (float.TryParse(parts[1].Trim(), out float statusSize))
                                settings.StatusFontSize = statusSize;
                            break;
                        case "OnPremise_AdminEmail": settings.OnPremise_AdminEmail = parts[1].Trim(); break;
                        case "OnPremise_Password": settings.OnPremise_Password = parts[1].Trim(); break;
                        case "OnPremise_EwsUrl": settings.OnPremise_EwsUrl = parts[1].Trim(); break;
                        case "OnPremise_Domain": settings.OnPremise_Domain = parts[1].Trim(); break;
                        case "OnPremiseAccounts":
                            try {
                                settings.OnPremiseAccounts = Newtonsoft.Json.JsonConvert.DeserializeObject<List<OnPremiseAccount>>(parts[1].Trim()) ?? new List<OnPremiseAccount>();
                            } catch { }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Warning(ex, "加载功能设置失败，使用默认设置");
            }

            return settings;
        }

        /// <summary>
        /// 保存功能设置
        /// </summary>
        public static void Save(FeatureSettings settings)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine($"Feature_ToPst={settings.Feature_ToPst}");
                sb.AppendLine($"Feature_ToPst_Eml={settings.Feature_ToPst_Eml}");
                sb.AppendLine($"Feature_ToPst_Ost={settings.Feature_ToPst_Ost}");
                sb.AppendLine($"Feature_ToPst_Imap={settings.Feature_ToPst_Imap}");
                sb.AppendLine($"Feature_ToPst_MultiImap={settings.Feature_ToPst_MultiImap}");
                sb.AppendLine($"Feature_Extract={settings.Feature_Extract}");
                sb.AppendLine($"Feature_Extract_Imap={settings.Feature_Extract_Imap}");
                sb.AppendLine($"Feature_Extract_Files={settings.Feature_Extract_Files}");
                sb.AppendLine($"Feature_SingleUserSync={settings.Feature_SingleUserSync}");
                sb.AppendLine($"Feature_SingleUserSync_EmlImport={settings.Feature_SingleUserSync_EmlImport}");
                sb.AppendLine($"Feature_SingleUserSync_Contacts={settings.Feature_SingleUserSync_Contacts}");
                sb.AppendLine($"Feature_BatchSync={settings.Feature_BatchSync}");
                sb.AppendLine($"Feature_BatchSync_Login={settings.Feature_BatchSync_Login}");
                sb.AppendLine($"Feature_BatchSync_PstMail={settings.Feature_BatchSync_PstMail}");
                sb.AppendLine($"Feature_BatchSync_PstContacts={settings.Feature_BatchSync_PstContacts}");
                sb.AppendLine($"Feature_BatchSync_PstCalendar={settings.Feature_BatchSync_PstCalendar}");
                sb.AppendLine($"Feature_BatchSync_CsvContacts={settings.Feature_BatchSync_CsvContacts}");
                sb.AppendLine($"Feature_BatchSync_VcfContacts={settings.Feature_BatchSync_VcfContacts}");
                sb.AppendLine($"Feature_BatchSync_CsvCalendar={settings.Feature_BatchSync_CsvCalendar}");
                sb.AppendLine($"Feature_BatchSync_Purview={settings.Feature_BatchSync_Purview}");
                sb.AppendLine($"Feature_O365Toolkit={settings.Feature_O365Toolkit}");
                sb.AppendLine($"Feature_O365Toolkit_Login={settings.Feature_O365Toolkit_Login}");
                sb.AppendLine($"Feature_O365Toolkit_Account={settings.Feature_O365Toolkit_Account}");
                sb.AppendLine($"Feature_O365Toolkit_Group={settings.Feature_O365Toolkit_Group}");
                sb.AppendLine($"Feature_O365Toolkit_Mobile={settings.Feature_O365Toolkit_Mobile}");
                sb.AppendLine($"Feature_O365Toolkit_Traffic={settings.Feature_O365Toolkit_Traffic}");
                sb.AppendLine($"Feature_O365Toolkit_Migration={settings.Feature_O365Toolkit_Migration}");
                sb.AppendLine($"Feature_O365Toolkit_Whois={settings.Feature_O365Toolkit_Whois}");
                sb.AppendLine($"Feature_O365Toolkit_Dns={settings.Feature_O365Toolkit_Dns}");
                sb.AppendLine($"Feature_O365Toolkit_MailSearch={settings.Feature_O365Toolkit_MailSearch}");
                sb.AppendLine($"Feature_OnPremiseToolkit={settings.Feature_OnPremiseToolkit}");
                sb.AppendLine($"Feature_Preferences={settings.Feature_Preferences}");
                sb.AppendLine($"LogFontName={settings.LogFontName}");
                sb.AppendLine($"LogFontSize={settings.LogFontSize}");
                sb.AppendLine($"StatusFontName={settings.StatusFontName}");
                sb.AppendLine($"StatusFontSize={settings.StatusFontSize}");
                sb.AppendLine($"OnPremise_AdminEmail={settings.OnPremise_AdminEmail}");
                sb.AppendLine($"OnPremise_Password={settings.OnPremise_Password}");
                sb.AppendLine($"OnPremise_EwsUrl={settings.OnPremise_EwsUrl}");
                sb.AppendLine($"OnPremise_Domain={settings.OnPremise_Domain}");
                sb.AppendLine($"OnPremiseAccounts={Newtonsoft.Json.JsonConvert.SerializeObject(settings.OnPremiseAccounts)}");

                File.WriteAllText(FeatureSettingsFile, sb.ToString());
            }
            catch (Exception ex)
            {
                Serilog.Log.Error(ex, "保存功能设置失败");
            }
        }
    }

    /// <summary>
    /// 功能设置类
    /// </summary>
    public class FeatureSettings
    {
        // 主功能开关
        public bool Feature_ToPst { get; set; } = true;
        public bool Feature_Extract { get; set; } = true;
        public bool Feature_SingleUserSync { get; set; } = true;
        public bool Feature_BatchSync { get; set; } = true;
        public bool Feature_O365Toolkit { get; set; } = true;
        public bool Feature_OnPremiseToolkit { get; set; } = true;
        public bool Feature_Preferences { get; set; } = true;

        // 转换为PST子功能
        public bool Feature_ToPst_Eml { get; set; } = true;
        public bool Feature_ToPst_Ost { get; set; } = true;
        public bool Feature_ToPst_Imap { get; set; } = true;
        public bool Feature_ToPst_MultiImap { get; set; } = true;

        // 邮件提取子功能
        public bool Feature_Extract_Imap { get; set; } = true;
        public bool Feature_Extract_Files { get; set; } = true;

        // 单用户同步O365子功能
        public bool Feature_SingleUserSync_EmlImport { get; set; } = true;
        public bool Feature_SingleUserSync_Contacts { get; set; } = true;

        // 批量同步O365子功能
        public bool Feature_BatchSync_Login { get; set; } = true;
        public bool Feature_BatchSync_PstMail { get; set; } = true;
        public bool Feature_BatchSync_PstContacts { get; set; } = true;
        public bool Feature_BatchSync_PstCalendar { get; set; } = true;
        public bool Feature_BatchSync_CsvContacts { get; set; } = true;
        public bool Feature_BatchSync_VcfContacts { get; set; } = true;
        public bool Feature_BatchSync_CsvCalendar { get; set; } = true;
        public bool Feature_BatchSync_Purview { get; set; } = true;

        // Exchange Online 百宝箱子功能
        public bool Feature_O365Toolkit_Login { get; set; } = true;
        public bool Feature_O365Toolkit_Account { get; set; } = true;
        public bool Feature_O365Toolkit_Group { get; set; } = true;
        public bool Feature_O365Toolkit_Mobile { get; set; } = true;
        public bool Feature_O365Toolkit_Traffic { get; set; } = true;
        public bool Feature_O365Toolkit_Migration { get; set; } = true;
        public bool Feature_O365Toolkit_Whois { get; set; } = true;
        public bool Feature_O365Toolkit_Dns { get; set; } = true;
        public bool Feature_O365Toolkit_MailSearch { get; set; } = true;

        // 界面字体设置
        public string LogFontName { get; set; } = "Consolas";
        public float LogFontSize { get; set; } = 10F;
        public string StatusFontName { get; set; } = "Microsoft Sans Serif";
        public float StatusFontSize { get; set; } = 10F;

        // Exchange On-Premise 管理员默认设置（兼容旧版本）
        public string OnPremise_AdminEmail { get; set; } = "";
        public string OnPremise_Password { get; set; } = "";
        public string OnPremise_EwsUrl { get; set; } = "";
        public string OnPremise_Domain { get; set; } = "";

        // Exchange On-Premise 多账户列表
        public List<OnPremiseAccount> OnPremiseAccounts { get; set; } = new List<OnPremiseAccount>();

        public FeatureSettings Clone()
        {
            return new FeatureSettings
            {
                Feature_ToPst = this.Feature_ToPst,
                Feature_ToPst_Eml = this.Feature_ToPst_Eml,
                Feature_ToPst_Ost = this.Feature_ToPst_Ost,
                Feature_ToPst_Imap = this.Feature_ToPst_Imap,
                Feature_ToPst_MultiImap = this.Feature_ToPst_MultiImap,
                Feature_Extract = this.Feature_Extract,
                Feature_Extract_Imap = this.Feature_Extract_Imap,
                Feature_Extract_Files = this.Feature_Extract_Files,
                Feature_SingleUserSync = this.Feature_SingleUserSync,
                Feature_SingleUserSync_EmlImport = this.Feature_SingleUserSync_EmlImport,
                Feature_SingleUserSync_Contacts = this.Feature_SingleUserSync_Contacts,
                Feature_BatchSync = this.Feature_BatchSync,
                Feature_BatchSync_Login = this.Feature_BatchSync_Login,
                Feature_BatchSync_PstMail = this.Feature_BatchSync_PstMail,
                Feature_BatchSync_PstContacts = this.Feature_BatchSync_PstContacts,
                Feature_BatchSync_PstCalendar = this.Feature_BatchSync_PstCalendar,
                Feature_BatchSync_CsvContacts = this.Feature_BatchSync_CsvContacts,
                Feature_BatchSync_VcfContacts = this.Feature_BatchSync_VcfContacts,
                Feature_BatchSync_CsvCalendar = this.Feature_BatchSync_CsvCalendar,
                Feature_BatchSync_Purview = this.Feature_BatchSync_Purview,
                Feature_O365Toolkit = this.Feature_O365Toolkit,
                Feature_O365Toolkit_Login = this.Feature_O365Toolkit_Login,
                Feature_O365Toolkit_Account = this.Feature_O365Toolkit_Account,
                Feature_O365Toolkit_Group = this.Feature_O365Toolkit_Group,
                Feature_O365Toolkit_Mobile = this.Feature_O365Toolkit_Mobile,
                Feature_O365Toolkit_Traffic = this.Feature_O365Toolkit_Traffic,
                Feature_O365Toolkit_Migration = this.Feature_O365Toolkit_Migration,
                Feature_O365Toolkit_Whois = this.Feature_O365Toolkit_Whois,
                Feature_O365Toolkit_Dns = this.Feature_O365Toolkit_Dns,
                Feature_O365Toolkit_MailSearch = this.Feature_O365Toolkit_MailSearch,
                Feature_OnPremiseToolkit = this.Feature_OnPremiseToolkit,
                Feature_Preferences = this.Feature_Preferences,
                LogFontName = this.LogFontName,
                LogFontSize = this.LogFontSize,
                StatusFontName = this.StatusFontName,
                StatusFontSize = this.StatusFontSize,
                OnPremise_AdminEmail = this.OnPremise_AdminEmail,
                OnPremise_Password = this.OnPremise_Password,
                OnPremise_EwsUrl = this.OnPremise_EwsUrl,
                OnPremise_Domain = this.OnPremise_Domain,
                OnPremiseAccounts = new List<OnPremiseAccount>(this.OnPremiseAccounts)
            };
        }
    }
}
