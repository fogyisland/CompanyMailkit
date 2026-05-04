using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;

namespace MailConverter
{
    /// <summary>
    /// 配置文件服务 - 使用独立目录和INF文件存储各类配置
    /// 目录结构:
    /// Config/
    ///   oauth/        - OAuth账户(*.inf)
    ///   imap/         - IMAP账户(*.inf)
    ///   carddav/      - CardDAV账户(*.inf)
    ///   pst/          - PST账户(*.inf)
    ///   preferences.inf  - 应用首选项
    ///   registration.inf - 注册信息
    /// </summary>
    public class ConfigService
    {
        private static readonly string ConfigDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config");
        private static readonly string OAuthDir = Path.Combine(ConfigDir, "oauth");
        private static readonly string ImapDir = Path.Combine(ConfigDir, "imap");
        private static readonly string CardDavDir = Path.Combine(ConfigDir, "carddav");
        private static readonly string PstDir = Path.Combine(ConfigDir, "pst");
        private static readonly string PreferencesFile = Path.Combine(ConfigDir, "preferences.inf");
        private static readonly string RegistrationFile = Path.Combine(ConfigDir, "registration.inf");

        static ConfigService()
        {
            // 确保目录存在
            Directory.CreateDirectory(ConfigDir);
            Directory.CreateDirectory(OAuthDir);
            Directory.CreateDirectory(ImapDir);
            Directory.CreateDirectory(CardDavDir);
            Directory.CreateDirectory(PstDir);
        }

        #region OAuth账户

        public static List<OAuthAccount> LoadOAuthAccounts()
        {
            var accounts = new List<OAuthAccount>();
            if (!Directory.Exists(OAuthDir)) return accounts;

            foreach (var file in Directory.GetFiles(OAuthDir, "*.inf"))
            {
                var dict = LoadInfDict(file);
                if (dict == null || !dict.ContainsKey("Name")) continue;

                var account = new OAuthAccount
                {
                    Name = dict["Name"],
                    ClientId = dict.ContainsKey("ClientId") ? dict["ClientId"] : "",
                    TenantId = dict.ContainsKey("TenantId") ? dict["TenantId"] : "",
                    Email = dict.ContainsKey("Email") ? dict["Email"] : ""
                };
                accounts.Add(account);
            }
            return accounts;
        }

        public static void SaveOAuthAccount(OAuthAccount account)
        {
            var file = Path.Combine(OAuthDir, SanitizeFileName(account.Name) + ".inf");
            SaveInfFile(file, account.Name, new Dictionary<string, string>
            {
                { "ClientId", account.ClientId },
                { "TenantId", account.TenantId },
                { "Email", account.Email }
            });
        }

        public static void DeleteOAuthAccount(string name)
        {
            var file = Path.Combine(OAuthDir, SanitizeFileName(name) + ".inf");
            if (File.Exists(file)) File.Delete(file);
        }

        #endregion

        #region IMAP账户

        public static List<ImapAccountSetting> LoadImapAccounts()
        {
            var accounts = new List<ImapAccountSetting>();
            if (!Directory.Exists(ImapDir)) return accounts;

            foreach (var file in Directory.GetFiles(ImapDir, "*.inf"))
            {
                var dict = LoadInfDict(file);
                if (dict == null || !dict.ContainsKey("Name")) continue;

                int.TryParse(dict.ContainsKey("Port") ? dict["Port"] : "993", out int port);
                bool.TryParse(dict.ContainsKey("UseSsl") ? dict["UseSsl"] : "true", out bool useSsl);

                var account = new ImapAccountSetting
                {
                    Name = dict["Name"],
                    Host = dict.ContainsKey("Host") ? dict["Host"] : "",
                    Port = port > 0 ? port : 993,
                    UseSsl = useSsl,
                    Email = dict.ContainsKey("Email") ? dict["Email"] : "",
                    Password = dict.ContainsKey("Password") ? dict["Password"] : ""
                };
                accounts.Add(account);
            }
            return accounts;
        }

        public static void SaveImapAccount(ImapAccountSetting account)
        {
            var file = Path.Combine(ImapDir, SanitizeFileName(account.Name) + ".inf");
            SaveInfFile(file, account.Name, new Dictionary<string, string>
            {
                { "Host", account.Host },
                { "Port", account.Port.ToString() },
                { "UseSsl", account.UseSsl.ToString() },
                { "Email", account.Email },
                { "Password", account.Password }
            });
        }

        public static void DeleteImapAccount(string name)
        {
            var file = Path.Combine(ImapDir, SanitizeFileName(name) + ".inf");
            if (File.Exists(file)) File.Delete(file);
        }

        #endregion

        #region CardDAV账户

        public static List<CardDavAccount> LoadCardDavAccounts()
        {
            var accounts = new List<CardDavAccount>();
            if (!Directory.Exists(CardDavDir)) return accounts;

            foreach (var file in Directory.GetFiles(CardDavDir, "*.inf"))
            {
                var dict = LoadInfDict(file);
                if (dict == null || !dict.ContainsKey("Name")) continue;

                var account = new CardDavAccount
                {
                    Name = dict["Name"],
                    Provider = dict.ContainsKey("Provider") ? dict["Provider"] : "",
                    ServerUrl = dict.ContainsKey("ServerUrl") ? dict["ServerUrl"] : ""
                };
                accounts.Add(account);
            }
            return accounts;
        }

        public static void SaveCardDavAccount(CardDavAccount account)
        {
            var file = Path.Combine(CardDavDir, SanitizeFileName(account.Name) + ".inf");
            SaveInfFile(file, account.Name, new Dictionary<string, string>
            {
                { "Provider", account.Provider },
                { "ServerUrl", account.ServerUrl }
            });
        }

        public static void DeleteCardDavAccount(string name)
        {
            var file = Path.Combine(CardDavDir, SanitizeFileName(name) + ".inf");
            if (File.Exists(file)) File.Delete(file);
        }

        #endregion

        #region PST账户

        public static List<PstAccount> LoadPstAccounts()
        {
            var accounts = new List<PstAccount>();
            if (!Directory.Exists(PstDir)) return accounts;

            foreach (var file in Directory.GetFiles(PstDir, "*.inf"))
            {
                var dict = LoadInfDict(file);
                if (dict == null || !dict.ContainsKey("Name")) continue;

                var account = new PstAccount
                {
                    Name = dict["Name"],
                    TenantId = dict.ContainsKey("TenantId") ? dict["TenantId"] : "",
                    ClientId = dict.ContainsKey("ClientId") ? dict["ClientId"] : "",
                    ClientSecret = dict.ContainsKey("ClientSecret") ? dict["ClientSecret"] : "",
                    AccountName = dict.ContainsKey("AccountName") ? dict["AccountName"] : ""
                };
                accounts.Add(account);
            }
            return accounts;
        }

        public static void SavePstAccount(PstAccount account)
        {
            var file = Path.Combine(PstDir, SanitizeFileName(account.Name) + ".inf");
            SaveInfFile(file, account.Name, new Dictionary<string, string>
            {
                { "TenantId", account.TenantId },
                { "ClientId", account.ClientId },
                { "ClientSecret", account.ClientSecret },
                { "AccountName", account.AccountName }
            });
        }

        public static void DeletePstAccount(string name)
        {
            var file = Path.Combine(PstDir, SanitizeFileName(name) + ".inf");
            if (File.Exists(file)) File.Delete(file);
        }

        #endregion

        #region 首选项

        public static AppSettings LoadPreferences()
        {
            var settings = new AppSettings();
            if (!File.Exists(PreferencesFile)) return settings;

            var lines = File.ReadAllLines(PreferencesFile);
            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line) || !line.Contains("="))
                    continue;

                var parts = line.Split(new[] { '=' }, 2);
                var key = parts[0].Trim().ToLower();
                var value = parts[1].Trim();

                switch (key)
                {
                    case "lastusedemail": settings.LastUsedEmail = value; break;
                    case "lastsourcepath": settings.LastSourcePath = value; break;
                    case "lasttargetfolder": settings.LastTargetFolder = value; break;
                    case "psttenantid": settings.PstTenantId = value; break;
                    case "pstclientid": settings.PstClientId = value; break;
                    case "pstclientsecret": settings.PstClientSecret = value; break;
                    case "pstaccountname": settings.PstAccountName = value; break;
                    case "purviewlogpath": settings.PurviewLogPath = value; break;
                    case "purviewoutputpath": settings.PurviewOutputPath = value; break;
                }
            }
            return settings;
        }

        public static void SavePreferences(AppSettings settings)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"LastUsedEmail={settings.LastUsedEmail}");
            sb.AppendLine($"LastSourcePath={settings.LastSourcePath}");
            sb.AppendLine($"LastTargetFolder={settings.LastTargetFolder}");
            sb.AppendLine($"PstTenantId={settings.PstTenantId}");
            sb.AppendLine($"PstClientId={settings.PstClientId}");
            sb.AppendLine($"PstClientSecret={settings.PstClientSecret}");
            sb.AppendLine($"PstAccountName={settings.PstAccountName}");
            sb.AppendLine($"PurviewLogPath={settings.PurviewLogPath}");
            sb.AppendLine($"PurviewOutputPath={settings.PurviewOutputPath}");
            File.WriteAllText(PreferencesFile, sb.ToString());
        }

        #endregion

        #region 注册信息

        public static void SaveRegistration(AppSettings settings)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"IsRegistered={settings.IsRegistered}");
            sb.AppendLine($"RegisteredUserName={settings.RegisteredUserName}");
            sb.AppendLine($"RegisteredUserEmail={settings.RegisteredUserEmail}");
            sb.AppendLine($"RegisteredOrganization={settings.RegisteredOrganization}");
            sb.AppendLine($"RegisteredMacAddress={settings.RegisteredMacAddress}");
            if (settings.RegisterDate.HasValue)
                sb.AppendLine($"RegisterDate={settings.RegisterDate.Value:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"RegisterSerialNumber={settings.RegisterSerialNumber}");
            sb.AppendLine($"RegisterRemainingDays={settings.RegisterRemainingDays}");
            sb.AppendLine($"RegisterExpireDate={settings.RegisterExpireDate}");
            if (settings.FirstRunDate.HasValue)
                sb.AppendLine($"FirstRunDate={settings.FirstRunDate.Value:yyyy-MM-dd}");
            File.WriteAllText(RegistrationFile, sb.ToString());
        }

        public static AppSettings LoadRegistration(AppSettings settings)
        {
            if (!File.Exists(RegistrationFile)) return settings;

            var lines = File.ReadAllLines(RegistrationFile);
            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line) || !line.Contains("="))
                    continue;

                var parts = line.Split(new[] { '=' }, 2);
                var key = parts[0].Trim().ToLower();
                var value = parts[1].Trim();

                switch (key)
                {
                    case "isregistered": settings.IsRegistered = value.ToLower() == "true"; break;
                    case "registeredusername": settings.RegisteredUserName = value; break;
                    case "registereduseremail": settings.RegisteredUserEmail = value; break;
                    case "registeredorganization": settings.RegisteredOrganization = value; break;
                    case "registeredmacaddress": settings.RegisteredMacAddress = value; break;
                    case "registerdate":
                        if (DateTime.TryParse(value, out DateTime regDate))
                            settings.RegisterDate = regDate;
                        break;
                    case "registerserialnumber": settings.RegisterSerialNumber = value; break;
                    case "registerremainingdays":
                        if (int.TryParse(value, out int rdays))
                            settings.RegisterRemainingDays = rdays;
                        break;
                    case "registerexpiredate": settings.RegisterExpireDate = value; break;
                    case "firstrundate":
                        if (DateTime.TryParse(value, out DateTime frd))
                            settings.FirstRunDate = frd;
                        break;
                }
            }
            return settings;
        }

        #endregion

        #region 完整加载/保存

        public static AppSettings LoadAll()
        {
            // 检查是否需要迁移旧数据
            string oldSettingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.txt");
            bool needsMigration = false;

            if (File.Exists(oldSettingsPath))
            {
                // 检查旧文件是否有注册信息
                var oldLines = File.ReadAllLines(oldSettingsPath);
                bool oldIsRegistered = false;
                foreach (var line in oldLines)
                {
                    if (line.StartsWith("IsRegistered=", StringComparison.OrdinalIgnoreCase))
                    {
                        oldIsRegistered = line.Substring(13).Trim().Equals("True", StringComparison.OrdinalIgnoreCase);
                        break;
                    }
                }

                // 如果旧文件已注册，但新文件未注册或不存在，则需要迁移
                if (oldIsRegistered)
                {
                    if (!File.Exists(RegistrationFile))
                    {
                        needsMigration = true;
                    }
                    else
                    {
                        // 检查新文件是否有效注册
                        var newLines = File.ReadAllLines(RegistrationFile);
                        bool newIsRegistered = false;
                        foreach (var line in newLines)
                        {
                            if (line.StartsWith("IsRegistered=", StringComparison.OrdinalIgnoreCase))
                            {
                                newIsRegistered = line.Substring(13).Trim().Equals("True", StringComparison.OrdinalIgnoreCase);
                                break;
                            }
                        }
                        if (!newIsRegistered)
                            needsMigration = true;
                    }
                }
            }

            if (needsMigration)
            {
                MigrateFromOldSettings();
            }

            var settings = LoadPreferences();
            settings = LoadRegistration(settings);
            settings.OAuthAccounts = LoadOAuthAccounts();
            settings.ImapAccounts = LoadImapAccounts();
            settings.CardDavAccounts = LoadCardDavAccounts();
            settings.PstAccounts = LoadPstAccounts();
            return settings;
        }

        public static void SaveAll(AppSettings settings)
        {
            SavePreferences(settings);
            SaveRegistration(settings);

            // 保存所有账户类型
            foreach (var acc in settings.OAuthAccounts)
                SaveOAuthAccount(acc);
            foreach (var acc in settings.ImapAccounts)
                SaveImapAccount(acc);
            foreach (var acc in settings.CardDavAccounts)
                SaveCardDavAccount(acc);
            foreach (var acc in settings.PstAccounts)
                SavePstAccount(acc);
        }

        #endregion

        #region 辅助方法

        private static Dictionary<string, string> LoadInfDict(string file)
        {
            try
            {
                var dict = new Dictionary<string, string>();
                var lines = File.ReadAllLines(file);

                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    if (line.StartsWith("[") && line.EndsWith("]"))
                    {
                        dict["Name"] = line.Substring(1, line.Length - 2).Trim();
                    }
                    else if (line.Contains("="))
                    {
                        var parts = line.Split(new[] { '=' }, 2);
                        var key = parts[0].Trim();
                        var value = parts[1].Trim();
                        dict[key] = value;
                    }
                }

                return dict;
            }
            catch (Exception ex)
            {
                Serilog.Log.Warning(ex, "加载INF文件失败: {File}", file);
                return null;
            }
        }

        private static void SaveInfFile(string file, string name, Dictionary<string, string> values)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"[{name}]");
            foreach (var kvp in values)
            {
                sb.AppendLine($"{kvp.Key}={kvp.Value}");
            }
            File.WriteAllText(file, sb.ToString());
        }

        private static string SanitizeFileName(string name)
        {
            var invalid = Path.GetInvalidFileNameChars();
            foreach (var c in invalid)
            {
                name = name.Replace(c, '_');
            }
            return name;
        }

        #endregion

        #region 迁移旧数据

        /// <summary>
        /// 从旧的settings.txt迁移到新的Config目录结构
        /// </summary>
        public static void MigrateFromOldSettings()
        {
            var oldPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.txt");
            if (!File.Exists(oldPath))
                return;

            try
            {
                var settings = SettingsService.Load();
                SaveAll(settings);
                Serilog.Log.Information("配置迁移完成，旧文件已保留");
            }
            catch (Exception ex)
            {
                Serilog.Log.Error(ex, "配置迁移失败");
            }
        }

        #endregion

        #region 获取配置目录信息

        public static Dictionary<string, object> GetConfigInfo()
        {
            var info = new Dictionary<string, object>();

            info["ConfigDir"] = ConfigDir;
            info["OAuthCount"] = Directory.Exists(OAuthDir) ? Directory.GetFiles(OAuthDir, "*.inf").Length : 0;
            info["ImapCount"] = Directory.Exists(ImapDir) ? Directory.GetFiles(ImapDir, "*.inf").Length : 0;
            info["CardDavCount"] = Directory.Exists(CardDavDir) ? Directory.GetFiles(CardDavDir, "*.inf").Length : 0;
            info["PstCount"] = Directory.Exists(PstDir) ? Directory.GetFiles(PstDir, "*.inf").Length : 0;
            info["HasPreferences"] = File.Exists(PreferencesFile);
            info["HasRegistration"] = File.Exists(RegistrationFile);

            return info;
        }

        #endregion
    }
}
