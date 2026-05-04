using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MailConverter
{
    /// <summary>
    /// EWS投递IMAP账户配置
    /// </summary>
    public class EwsToImapAccount
    {
        public string Name { get; set; } = "";
        public string ImapServer { get; set; } = "";
        public int ImapPort { get; set; } = 993;
        public string Username { get; set; } = "";
        public string Password { get; set; } = "";
        public bool UseSsl { get; set; } = true;
    }

    /// <summary>
    /// 注册表服务 - 将注册/激活信息存储在HKCU（不需要管理员权限）
    /// 同时与 registration.inf 双向同步
    /// 注册表结构:
    /// HKCU\SOFTWARE\MailConverter\
    ///   ├── register\    - 注册信息
    ///   └── activate\     - 激活信息
    /// </summary>
    public static class RegistryService
    {
        private const string RegistryPath = @"SOFTWARE\MailConverter";
        private const string RegisterPath = @"SOFTWARE\MailConverter\register";
        private const string ActivatePath = @"SOFTWARE\MailConverter\activate";
        private static readonly string InfFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config", "registration.inf");

        static RegistryService()
        {
            // 确保Config目录存在
            var configDir = Path.GetDirectoryName(InfFilePath);
            if (!string.IsNullOrEmpty(configDir))
                Directory.CreateDirectory(configDir);
        }

        #region 统一加载（注册表优先，无则读INF）

        /// <summary>
        /// 统一加载注册信息（注册表为主，INF为备）
        /// </summary>
        public static void LoadRegistration(AppSettings settings)
        {
            bool fromRegistry = false;
            bool fromInf = false;

            // 1. 先尝试从注册表加载
            if (HasRegistrationRecord())
            {
                LoadFromRegistry(settings);
                fromRegistry = true;
            }

            // 2. 如果注册表没有，尝试从INF加载
            if (!fromRegistry && File.Exists(InfFilePath))
            {
                LoadFromInf(settings);
                fromInf = true;
            }

            // 3. 如果从INF加载了数据，同步到注册表
            if (fromInf && settings.IsRegistered)
            {
                SaveToRegistry(settings);
            }
        }

        /// <summary>
        /// 统一保存注册信息（同时写入注册表和INF）
        /// </summary>
        public static void SaveRegistration(AppSettings settings)
        {
            // 同时保存到注册表和INF
            SaveToRegistry(settings);
            SaveToInf(settings);
        }

        /// <summary>
        /// 统一加载激活信息
        /// </summary>
        public static void LoadActivation(AppSettings settings)
        {
            // 优先从注册表加载
            if (HasActivationRecord())
            {
                LoadActivationFromRegistry(settings);
            }
            else
            {
                // 从INF加载
                LoadActivationFromInf(settings);
            }
        }

        /// <summary>
        /// 统一保存激活信息
        /// </summary>
        public static void SaveActivation(AppSettings settings)
        {
            SaveActivationToRegistry(settings);
            SaveActivationToInf(settings);
        }

        #endregion

        #region 注册表操作 (HKCU - 不需要管理员)

        private static void SaveToRegistry(AppSettings settings)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(RegisterPath))
                {
                    if (key == null)
                    {
                        Serilog.Log.Warning("无法创建注册表项: {Path}", RegisterPath);
                        return;
                    }

                    try { key.SetValue("IsRegistered", settings.IsRegistered ? 1 : 0); }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "保存IsRegistered失败"); }

                    try { key.SetValue("UserName", settings.RegisteredUserName ?? ""); }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "保存UserName失败"); }

                    try { key.SetValue("UserEmail", settings.RegisteredUserEmail ?? ""); }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "保存UserEmail失败"); }

                    try { key.SetValue("Organization", settings.RegisteredOrganization ?? ""); }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "保存Organization失败"); }

                    try { key.SetValue("MacAddress", settings.RegisteredMacAddress ?? ""); }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "保存MacAddress失败"); }

                    try { key.SetValue("RegisterDate", settings.RegisterDate?.ToString("yyyy-MM-dd") ?? ""); }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "保存RegisterDate失败"); }

                    try { key.SetValue("FirstRunDate", settings.FirstRunDate?.ToString("yyyy-MM-dd") ?? ""); }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "保存FirstRunDate失败"); }

                    try { key.SetValue("RemainingDays", settings.RegisterRemainingDays.HasValue ? settings.RegisterRemainingDays.Value : 0); }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "保存RemainingDays失败"); }

                    try { key.SetValue("ExpireDate", settings.RegisterExpireDate ?? ""); }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "保存ExpireDate失败"); }
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Error(ex, "保存注册信息到注册表失败");
            }
        }

        private static void LoadFromRegistry(AppSettings settings)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(RegisterPath))
                {
                    if (key == null) return;

                    try { settings.IsRegistered = ((int)key.GetValue("IsRegistered", 0)) == 1; }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "读取IsRegistered失败"); }

                    try { settings.RegisteredUserName = key.GetValue("UserName", "") as string ?? ""; }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "读取UserName失败"); }

                    try { settings.RegisteredUserEmail = key.GetValue("UserEmail", "") as string ?? ""; }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "读取UserEmail失败"); }

                    try { settings.RegisteredOrganization = key.GetValue("Organization", "") as string ?? ""; }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "读取Organization失败"); }

                    try { settings.RegisteredMacAddress = key.GetValue("MacAddress", "") as string ?? ""; }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "读取MacAddress失败"); }

                    try
                    {
                        var dateStr = key.GetValue("RegisterDate", "") as string;
                        if (!string.IsNullOrEmpty(dateStr) && DateTime.TryParse(dateStr, out var date))
                            settings.RegisterDate = date;
                    }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "读取RegisterDate失败"); }

                    try
                    {
                        var dateStr = key.GetValue("FirstRunDate", "") as string;
                        if (!string.IsNullOrEmpty(dateStr) && DateTime.TryParse(dateStr, out var date))
                            settings.FirstRunDate = date;
                    }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "读取FirstRunDate失败"); }

                    try
                    {
                        var remainingDays = key.GetValue("RemainingDays", 0);
                        if (remainingDays != null)
                            settings.RegisterRemainingDays = Convert.ToInt32(remainingDays);
                    }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "读取RemainingDays失败"); }

                    try { settings.RegisterExpireDate = key.GetValue("ExpireDate", "") as string ?? ""; }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "读取ExpireDate失败"); }
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Error(ex, "从注册表加载注册信息失败");
            }
        }

        private static void SaveActivationToRegistry(AppSettings settings)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(ActivatePath))
                {
                    if (key == null) return;

                    try { key.SetValue("SerialNumber", settings.RegisterSerialNumber ?? ""); }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "保存SerialNumber失败"); }

                    try { key.SetValue("ExpireDate", settings.RegisterExpireDate ?? ""); }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "保存ActivationExpireDate失败"); }

                    try { key.SetValue("RemainingDays", settings.RegisterRemainingDays.HasValue ? settings.RegisterRemainingDays.Value : 0); }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "保存ActivationRemainingDays失败"); }

                    try { key.SetValue("ActivatedDate", DateTime.Now.ToString("yyyy-MM-dd")); }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "保存ActivatedDate失败"); }
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Error(ex, "保存激活信息到注册表失败");
            }
        }

        private static void LoadActivationFromRegistry(AppSettings settings)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(ActivatePath))
                {
                    if (key == null) return;

                    try { settings.RegisterSerialNumber = key.GetValue("SerialNumber", "") as string ?? ""; }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "读取SerialNumber失败"); }

                    try
                    {
                        var expireDate = key.GetValue("ExpireDate", "") as string;
                        if (!string.IsNullOrEmpty(expireDate))
                            settings.RegisterExpireDate = expireDate;
                    }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "读取ActivationExpireDate失败"); }

                    try
                    {
                        var remainingDays = key.GetValue("RemainingDays", 0);
                        if (remainingDays != null)
                            settings.RegisterRemainingDays = Convert.ToInt32(remainingDays);
                    }
                    catch (Exception ex) { Serilog.Log.Warning(ex, "读取ActivationRemainingDays失败"); }
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Error(ex, "从注册表加载激活信息失败");
            }
        }

        #endregion

        #region INF文件操作

        private static void SaveToInf(AppSettings settings)
        {
            try
            {
                var sb = new System.Text.StringBuilder();
                sb.AppendLine($"IsRegistered={settings.IsRegistered}");
                sb.AppendLine($"RegisteredUserName={settings.RegisteredUserName}");
                sb.AppendLine($"RegisteredUserEmail={settings.RegisteredUserEmail}");
                sb.AppendLine($"RegisteredOrganization={settings.RegisteredOrganization}");
                sb.AppendLine($"RegisteredMacAddress={settings.RegisteredMacAddress}");
                if (settings.RegisterDate.HasValue)
                    sb.AppendLine($"RegisterDate={settings.RegisterDate.Value:yyyy-MM-dd HH:mm:ss}");
                if (settings.FirstRunDate.HasValue)
                    sb.AppendLine($"FirstRunDate={settings.FirstRunDate.Value:yyyy-MM-dd}");
                sb.AppendLine($"RegisterRemainingDays={settings.RegisterRemainingDays ?? 0}");
                sb.AppendLine($"RegisterExpireDate={settings.RegisterExpireDate}");
                File.WriteAllText(InfFilePath, sb.ToString());
            }
            catch (Exception ex)
            {
                Serilog.Log.Error(ex, "保存注册信息到INF失败");
            }
        }

        private static void LoadFromInf(AppSettings settings)
        {
            try
            {
                if (!File.Exists(InfFilePath)) return;

                var lines = File.ReadAllLines(InfFilePath);
                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line) || !line.Contains("="))
                        continue;

                    var parts = line.Split(new[] { '=' }, 2);
                    var key = parts[0].Trim().ToLower();
                    var value = parts.Length > 1 ? parts[1].Trim() : "";

                    switch (key)
                    {
                        case "isregistered": settings.IsRegistered = value.Equals("true", StringComparison.OrdinalIgnoreCase); break;
                        case "registeredusername": settings.RegisteredUserName = value; break;
                        case "registereduseremail": settings.RegisteredUserEmail = value; break;
                        case "registeredorganization": settings.RegisteredOrganization = value; break;
                        case "registeredmacaddress": settings.RegisteredMacAddress = value; break;
                        case "registerdate":
                            if (DateTime.TryParse(value, out var regDate))
                                settings.RegisterDate = regDate;
                            break;
                        case "firstrundate":
                            if (DateTime.TryParse(value, out var firstRun))
                                settings.FirstRunDate = firstRun;
                            break;
                        case "remainingdays":
                        case "registerremainingdays":
                            if (int.TryParse(value, out var days))
                                settings.RegisterRemainingDays = days;
                            break;
                        case "expiredate":
                        case "registerexpiredate":
                            settings.RegisterExpireDate = value;
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Error(ex, "从INF加载注册信息失败");
            }
        }

        private static void SaveActivationToInf(AppSettings settings)
        {
            try
            {
                // 读取现有INF，保留注册信息
                var existingLines = File.Exists(InfFilePath) ? File.ReadAllLines(InfFilePath).ToList() : new System.Collections.Generic.List<string>();

                // 查找并更新或添加激活相关字段
                bool hasSerial = false;
                bool hasExpire = false;
                bool hasRemaining = false;
                bool hasActivated = false;

                for (int i = 0; i < existingLines.Count; i++)
                {
                    var line = existingLines[i];
                    if (line.StartsWith("RegisterSerialNumber=", StringComparison.OrdinalIgnoreCase))
                    {
                        existingLines[i] = $"RegisterSerialNumber={settings.RegisterSerialNumber}";
                        hasSerial = true;
                    }
                    else if (line.StartsWith("ActivationExpireDate=", StringComparison.OrdinalIgnoreCase) || line.StartsWith("RegisterExpireDate=", StringComparison.OrdinalIgnoreCase))
                    {
                        existingLines[i] = $"RegisterExpireDate={settings.RegisterExpireDate}";
                        hasExpire = true;
                    }
                    else if (line.StartsWith("ActivationRemainingDays=", StringComparison.OrdinalIgnoreCase) || line.StartsWith("RegisterRemainingDays=", StringComparison.OrdinalIgnoreCase))
                    {
                        existingLines[i] = $"RegisterRemainingDays={settings.RegisterRemainingDays ?? 0}";
                        hasRemaining = true;
                    }
                    else if (line.StartsWith("ActivatedDate=", StringComparison.OrdinalIgnoreCase))
                    {
                        existingLines[i] = $"ActivatedDate={DateTime.Now:yyyy-MM-dd}";
                        hasActivated = true;
                    }
                }

                // 如果没有找到对应字段，追加
                if (!hasSerial && !string.IsNullOrEmpty(settings.RegisterSerialNumber))
                    existingLines.Add($"RegisterSerialNumber={settings.RegisterSerialNumber}");
                if (!hasExpire && !string.IsNullOrEmpty(settings.RegisterExpireDate))
                    existingLines.Add($"RegisterExpireDate={settings.RegisterExpireDate}");
                if (!hasRemaining && settings.RegisterRemainingDays.HasValue)
                    existingLines.Add($"RegisterRemainingDays={settings.RegisterRemainingDays.Value}");
                if (!hasActivated)
                    existingLines.Add($"ActivatedDate={DateTime.Now:yyyy-MM-dd}");

                File.WriteAllLines(InfFilePath, existingLines);
            }
            catch (Exception ex)
            {
                Serilog.Log.Error(ex, "保存激活信息到INF失败");
            }
        }

        private static void LoadActivationFromInf(AppSettings settings)
        {
            try
            {
                if (!File.Exists(InfFilePath)) return;

                var lines = File.ReadAllLines(InfFilePath);
                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line) || !line.Contains("="))
                        continue;

                    var parts = line.Split(new[] { '=' }, 2);
                    var key = parts[0].Trim().ToLower();
                    var value = parts.Length > 1 ? parts[1].Trim() : "";

                    switch (key)
                    {
                        case "registerserialnumber":
                        case "serialnumber":
                            settings.RegisterSerialNumber = value;
                            break;
                        case "registerexpiredate":
                        case "activationexpiredate":
                        case "expiredate":
                            settings.RegisterExpireDate = value;
                            break;
                        case "registerremainingdays":
                        case "activationremainingdays":
                        case "remainingdays":
                            if (int.TryParse(value, out var days))
                                settings.RegisterRemainingDays = days;
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Error(ex, "从INF加载激活信息失败");
            }
        }

        #endregion

        #region 检查方法

        /// <summary>
        /// 检查注册表是否有注册记录
        /// </summary>
        public static bool HasRegistrationRecord()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(RegisterPath))
                {
                    if (key == null) return false;
                    var mac = key.GetValue("MacAddress", "") as string;
                    return !string.IsNullOrEmpty(mac);
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 检查注册表是否有激活记录
        /// </summary>
        public static bool HasActivationRecord()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(ActivatePath))
                {
                    if (key == null) return false;
                    var serial = key.GetValue("SerialNumber", "") as string;
                    return !string.IsNullOrEmpty(serial);
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 清除所有注册和激活信息
        /// </summary>
        public static void ClearRegistration()
        {
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(RegisterPath, false);
            }
            catch (Exception ex)
            {
                Serilog.Log.Warning(ex, "删除register子键失败");
            }

            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(ActivatePath, false);
            }
            catch (Exception ex)
            {
                Serilog.Log.Warning(ex, "删除activate子键失败");
            }

            // 同时删除INF
            try
            {
                if (File.Exists(InfFilePath))
                    File.Delete(InfFilePath);
            }
            catch (Exception ex)
            {
                Serilog.Log.Warning(ex, "删除registration.inf失败");
            }
        }

        #endregion

        #region Exchange On-Premise 连接账户 (本地Exchange)

        private const string OnPremiseAccountsPath = @"SOFTWARE\MailConverter\OnPremiseAccounts";
        private static readonly string OnPremiseAccountsInfPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config", "localexchangeConnection.inf");

        /// <summary>
        /// 保存On-Premise连接账户到注册表和INF文件
        /// </summary>
        public static void SaveOnPremiseAccounts(List<OnPremiseAccount> accounts)
        {
            // 序列化为JSON
            var json = Newtonsoft.Json.JsonConvert.SerializeObject(accounts);

            // 1. 保存到注册表
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(OnPremiseAccountsPath))
                {
                    if (key != null)
                    {
                        key.SetValue("Accounts", json);
                        key.SetValue("LastUpdated", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    }
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Warning(ex, "保存OnPremise账户到注册表失败");
            }

            // 2. 保存到INF文件
            try
            {
                var configDir = Path.GetDirectoryName(OnPremiseAccountsInfPath);
                if (!string.IsNullOrEmpty(configDir))
                    Directory.CreateDirectory(configDir);

                var sb = new System.Text.StringBuilder();
                sb.AppendLine($"Accounts={json}");
                sb.AppendLine($"LastUpdated={DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                File.WriteAllText(OnPremiseAccountsInfPath, sb.ToString());
            }
            catch (Exception ex)
            {
                Serilog.Log.Warning(ex, "保存OnPremise账户到INF失败");
            }
        }

        /// <summary>
        /// 从注册表或INF加载On-Premise连接账户
        /// </summary>
        public static List<OnPremiseAccount> LoadOnPremiseAccounts()
        {
            var accounts = new List<OnPremiseAccount>();
            bool loaded = false;

            // 1. 优先从注册表加载
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(OnPremiseAccountsPath))
                {
                    if (key != null)
                    {
                        var json = key.GetValue("Accounts") as string;
                        if (!string.IsNullOrEmpty(json))
                        {
                            accounts = Newtonsoft.Json.JsonConvert.DeserializeObject<List<OnPremiseAccount>>(json) ?? new List<OnPremiseAccount>();
                            loaded = true;
                            Serilog.Log.Information("从注册表加载了 {Count} 个OnPremise账户", accounts.Count);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Warning(ex, "从注册表加载OnPremise账户失败");
            }

            // 2. 如果注册表没有，从INF加载
            if (!loaded && File.Exists(OnPremiseAccountsInfPath))
            {
                try
                {
                    var lines = File.ReadAllLines(OnPremiseAccountsInfPath);
                    foreach (var line in lines)
                    {
                        if (line.StartsWith("Accounts=", StringComparison.OrdinalIgnoreCase))
                        {
                            var json = line.Substring("Accounts=".Length);
                            if (!string.IsNullOrEmpty(json))
                            {
                                accounts = Newtonsoft.Json.JsonConvert.DeserializeObject<List<OnPremiseAccount>>(json) ?? new List<OnPremiseAccount>();
                                Serilog.Log.Information("从INF文件加载了 {Count} 个OnPremise账户", accounts.Count);

                                // 同步到注册表
                                SaveOnPremiseAccounts(accounts);
                                break;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Serilog.Log.Warning(ex, "从INF加载OnPremise账户失败");
                }
            }

            return accounts;
        }

        /// <summary>
        /// 清除On-Premise连接账户
        /// </summary>
        public static void ClearOnPremiseAccounts()
        {
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(OnPremiseAccountsPath, false);
            }
            catch (Exception ex)
            {
                Serilog.Log.Warning(ex, "删除OnPremise注册表失败");
            }

            try
            {
                if (File.Exists(OnPremiseAccountsInfPath))
                    File.Delete(OnPremiseAccountsInfPath);
            }
            catch (Exception ex)
            {
                Serilog.Log.Warning(ex, "删除OnPremise INF失败");
            }
        }

        #endregion

        #region EWS投递IMAP 账户

        private const string EwsToImapAccountsPath = @"SOFTWARE\MailConverter\EwsToImapAccounts";
        private static readonly string EwsToImapAccountsInfPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config", "ewsToImapConnection.inf");

        /// <summary>
        /// 保存EWS投递IMAP账户到注册表和INF文件
        /// </summary>
        public static void SaveEwsToImapAccounts(List<EwsToImapAccount> accounts)
        {
            var json = Newtonsoft.Json.JsonConvert.SerializeObject(accounts);

            // 1. 保存到注册表
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(EwsToImapAccountsPath))
                {
                    if (key != null)
                    {
                        key.SetValue("Accounts", json);
                        key.SetValue("LastUpdated", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    }
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Warning(ex, "保存EwsToImap账户到注册表失败");
            }

            // 2. 保存到INF文件
            try
            {
                var configDir = Path.GetDirectoryName(EwsToImapAccountsInfPath);
                if (!string.IsNullOrEmpty(configDir))
                    Directory.CreateDirectory(configDir);

                var sb = new System.Text.StringBuilder();
                sb.AppendLine($"Accounts={json}");
                sb.AppendLine($"LastUpdated={DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                File.WriteAllText(EwsToImapAccountsInfPath, sb.ToString());
            }
            catch (Exception ex)
            {
                Serilog.Log.Warning(ex, "保存EwsToImap账户到INF失败");
            }
        }

        /// <summary>
        /// 从注册表或INF加载EWS投递IMAP账户
        /// </summary>
        public static List<EwsToImapAccount> LoadEwsToImapAccounts()
        {
            var accounts = new List<EwsToImapAccount>();
            bool loaded = false;

            // 1. 优先从注册表加载
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(EwsToImapAccountsPath))
                {
                    if (key != null)
                    {
                        var json = key.GetValue("Accounts") as string;
                        if (!string.IsNullOrEmpty(json))
                        {
                            accounts = Newtonsoft.Json.JsonConvert.DeserializeObject<List<EwsToImapAccount>>(json) ?? new List<EwsToImapAccount>();
                            loaded = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Warning(ex, "从注册表加载EwsToImap账户失败");
            }

            // 2. 如果注册表没有，从INF加载
            if (!loaded && File.Exists(EwsToImapAccountsInfPath))
            {
                try
                {
                    var lines = File.ReadAllLines(EwsToImapAccountsInfPath);
                    foreach (var line in lines)
                    {
                        if (line.StartsWith("Accounts=", StringComparison.OrdinalIgnoreCase))
                        {
                            var json = line.Substring("Accounts=".Length);
                            if (!string.IsNullOrEmpty(json))
                            {
                                accounts = Newtonsoft.Json.JsonConvert.DeserializeObject<List<EwsToImapAccount>>(json) ?? new List<EwsToImapAccount>();
                                // 同步到注册表
                                SaveEwsToImapAccounts(accounts);
                                break;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Serilog.Log.Warning(ex, "从INF加载EwsToImap账户失败");
                }
            }

            return accounts;
        }

        #endregion
    }
}
