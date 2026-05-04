using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.Win32;

namespace MailConverter
{
    public class AppSettings
    {
        public List<OAuthAccount> OAuthAccounts { get; set; } = new List<OAuthAccount>();
        public List<ImapAccountSetting> ImapAccounts { get; set; } = new List<ImapAccountSetting>();
        public List<CardDavAccount> CardDavAccounts { get; set; } = new List<CardDavAccount>();
        public List<PstAccount> PstAccounts { get; set; } = new List<PstAccount>();
        public string LastUsedEmail { get; set; } = "";
        public string LastSourcePath { get; set; } = "";
        public string LastTargetFolder { get; set; } = "Inbox";
        public string PstTenantId { get; set; } = "";
        public string PstClientId { get; set; } = "";
        public string PstClientSecret { get; set; } = "";
        public string PstAccountName { get; set; } = "";
        public string PurviewLogPath { get; set; } = "";
        public string PurviewOutputPath { get; set; } = "";
        // 注册信息
        public bool IsRegistered { get; set; } = false;
        public string RegisteredUserName { get; set; } = "";
        public string RegisteredUserEmail { get; set; } = "";
        public string RegisteredOrganization { get; set; } = "";
        public string RegisteredMacAddress { get; set; } = "";
        public DateTime? RegisterDate { get; set; }
        public string RegisterSerialNumber { get; set; } = "";
        public int? RegisterRemainingDays { get; set; }
        public string RegisterExpireDate { get; set; } = "";
        public DateTime? FirstRunDate { get; set; }
    }

    public class PstAccount
    {
        public string Name { get; set; } = "";
        public string TenantId { get; set; } = "";
        public string ClientId { get; set; } = "";
        public string ClientSecret { get; set; } = "";
        public string AccountName { get; set; } = "";
    }

    public class CardDavAccount
    {
        public string Name { get; set; } = "";
        public string Provider { get; set; } = "";
        public string ServerUrl { get; set; } = "";
    }

    public class OAuthAccount
    {
        public string Name { get; set; } = "";
        public string ClientId { get; set; } = "";
        public string TenantId { get; set; } = "";
        public string Email { get; set; } = "";
        public string AccessToken { get; set; } = "";
    }

    public class ImapAccountSetting
    {
        public string Name { get; set; } = "";
        public string Host { get; set; } = "";
        public int Port { get; set; } = 993;
        public bool UseSsl { get; set; } = true;
        public string Email { get; set; } = "";
        public string Password { get; set; } = "";
    }

    public class SettingsService
    {
        private const string RegistryKeyPath = @"SOFTWARE\MailConverter\Settings";
        private static readonly string SettingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.txt");

        private static RegistryKey OpenSettingsKey(bool writable = false)
        {
            if (writable)
                return Registry.CurrentUser.CreateSubKey(RegistryKeyPath);
            return Registry.CurrentUser.OpenSubKey(RegistryKeyPath);
        }

        public static AppSettings Load()
        {
            var settings = new AppSettings();
            try
            {
                // 优先从注册表加载
                using (var key = OpenSettingsKey())
                {
                    if (key != null)
                    {
                        var json = key.GetValue("Data") as string;
                        if (!string.IsNullOrEmpty(json))
                        {
                            var loaded = Newtonsoft.Json.JsonConvert.DeserializeObject<AppSettings>(json);
                            if (loaded != null)
                                return loaded;
                        }
                    }
                }

                // 注册表没有则从本地文件加载
                if (File.Exists(SettingsPath))
                {
                    var json = File.ReadAllText(SettingsPath);
                    if (!string.IsNullOrEmpty(json))
                    {
                        var loaded = Newtonsoft.Json.JsonConvert.DeserializeObject<AppSettings>(json);
                        if (loaded != null)
                            return loaded;
                    }
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Warning(ex, "加载设置失败");
            }
            return settings;
        }

        public static void Save(AppSettings settings)
        {
            try
            {
                var json = Newtonsoft.Json.JsonConvert.SerializeObject(settings);

                // 保存到注册表
                using (var key = OpenSettingsKey(true))
                {
                    if (key != null)
                    {
                        key.SetValue("Data", json);
                    }
                }

                // 同时保存到本地文件作为备份
                File.WriteAllText(SettingsPath, json);
            }
            catch (Exception ex)
            {
                Serilog.Log.Error(ex, "保存设置失败");
            }
        }

        public static void AddOrUpdateOAuthAccount(string name, string clientId, string tenantId, string email)
        {
            var settings = Load();
            var existing = settings.OAuthAccounts.FindIndex(a => a.Name == name);
            var account = new OAuthAccount { Name = name, ClientId = clientId, TenantId = tenantId, Email = email };

            if (existing >= 0) settings.OAuthAccounts[existing] = account;
            else settings.OAuthAccounts.Add(account);

            Save(settings);
        }

        public static void RemoveOAuthAccount(string name)
        {
            var settings = Load();
            settings.OAuthAccounts.RemoveAll(a => a.Name == name);
            Save(settings);
        }

        public static void AddOrUpdateImapAccount(string name, string host, int port, bool useSsl, string email, string password)
        {
            var settings = Load();
            var existing = settings.ImapAccounts.FindIndex(a => a.Name == name);
            var account = new ImapAccountSetting { Name = name, Host = host, Port = port, UseSsl = useSsl, Email = email, Password = password };

            if (existing >= 0) settings.ImapAccounts[existing] = account;
            else settings.ImapAccounts.Add(account);

            Save(settings);
        }

        public static void RemoveImapAccount(string name)
        {
            var settings = Load();
            settings.ImapAccounts.RemoveAll(a => a.Name == name);
            Save(settings);
        }

        public static void AddOrUpdateCardDavAccount(string name, string provider, string serverUrl)
        {
            var settings = Load();
            var existing = settings.CardDavAccounts.FindIndex(a => a.Name == name);
            var account = new CardDavAccount { Name = name, Provider = provider, ServerUrl = serverUrl };

            if (existing >= 0) settings.CardDavAccounts[existing] = account;
            else settings.CardDavAccounts.Add(account);

            Save(settings);
        }

        public static void RemoveCardDavAccount(string name)
        {
            var settings = Load();
            settings.CardDavAccounts.RemoveAll(a => a.Name == name);
            Save(settings);
        }

        public static void RemovePstAccount(string name)
        {
            var settings = Load();
            settings.PstAccounts.RemoveAll(a => a.Name == name);
            Save(settings);
        }
    }
}
