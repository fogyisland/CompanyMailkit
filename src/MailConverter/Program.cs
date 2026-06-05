using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using MimeKit;
using Serilog;

namespace MailConverter
{
    static class Program
    {
        // 日志目录
        public static string LogDirectory => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
        public static string O365LogDirectory => Path.Combine(LogDirectory, "O365Online");
        public static string LoginLogDirectory => Path.Combine(O365LogDirectory, "Login");
        public static string AccountManagementLogDirectory => Path.Combine(O365LogDirectory, "AccountManagement");
        public static string GroupManagementLogDirectory => Path.Combine(O365LogDirectory, "GroupManagement");
        public static string PurviewLogDirectory => Path.Combine(LogDirectory, "Purview");
        public static string BatchToO365LogDirectory => Path.Combine(LogDirectory, "batchToO365");
        public static string RegistrationLogDirectory => Path.Combine(LogDirectory, "Registration");
        public static string SyncCalendarLogDirectory => Path.Combine(LogDirectory, "syncCalandar");
        // 日历 ICS 导入专用日志(只放 ICS 同步,VCS/CSV 仍走 syncCalandar/)
        public static string O365SingleSyncIcsLogDirectory => Path.Combine(LogDirectory, "o365single", "SYNCICS");

        /// <summary>
        /// 获取Python可执行文件路径（仅使用嵌入式Python）
        /// </summary>
        public static string GetPythonExecutable()
        {
            var exeDir = AppDomain.CurrentDomain.BaseDirectory;
            var embeddedPython = Path.Combine(exeDir, "python", "python.exe");
            if (File.Exists(embeddedPython))
                return embeddedPython;
            return null; // 嵌入式Python不存在时返回null
        }

        /// <summary>
        /// 创建Python进程启动信息（禁用用户site-packages，避免引用外部Python环境）
        /// </summary>
        public static System.Diagnostics.ProcessStartInfo CreatePythonStartInfo(string pythonExe, string arguments)
        {
            var startInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = pythonExe,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = System.Text.Encoding.UTF8,
                StandardErrorEncoding = System.Text.Encoding.UTF8
            };
            // 禁用用户site-packages，避免引用外部Python环境
            startInfo.EnvironmentVariables["PYTHONNOUSERSITE"] = "1";
            startInfo.EnvironmentVariables["PYTHONUSERBASE"] = "";
            // 设置Python UTF-8模式
            startInfo.EnvironmentVariables["PYTHONIOENCODING"] = "UTF-8";
            return startInfo;
        }

        [STAThread]
        static void Main()
        {
            // 确保日志目录存在
            Directory.CreateDirectory(LogDirectory);
            Directory.CreateDirectory(O365LogDirectory);
            Directory.CreateDirectory(LoginLogDirectory);
            Directory.CreateDirectory(AccountManagementLogDirectory);
            Directory.CreateDirectory(GroupManagementLogDirectory);
            Directory.CreateDirectory(PurviewLogDirectory);
            Directory.CreateDirectory(BatchToO365LogDirectory);
            Directory.CreateDirectory(SyncCalendarLogDirectory);
            Directory.CreateDirectory(O365SingleSyncIcsLogDirectory);
            Directory.CreateDirectory(RegistrationLogDirectory);

            // 在WinForms初始化之前设置DPI感知（关键！必须最早调用）
            // 配置主日志
            var logPath = Path.Combine(LogDirectory, "oauth_import.log");
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .WriteTo.File(logPath,
                    rollingInterval: RollingInterval.Day,
                    retainedFileCountLimit: 7,
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                .CreateLogger();

            Log.Information("=== 应用程序启动 ===");

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // 启动时自动检查注册状态并同步
            Task.Run(async () =>
            {
                try
                {
                    await CheckAndSyncRegistrationAsync();
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "启动时注册状态同步失败");
                }
            });

            // 检查注册状态 - 优先检查注册表，INF文件作为备份
            var settings = ConfigService.LoadAll();

            // 如果INF文件没有注册信息，但注册表有，则从注册表恢复
            if (!settings.IsRegistered && RegistryService.HasRegistrationRecord())
            {
                RegistryService.LoadRegistration(settings);
                ConfigService.SaveRegistration(settings);
                Log.Information("从注册表恢复注册信息，剩余天数: {Days}", settings.RegisterRemainingDays);
            }

            if (!settings.IsRegistered)
            {
                var result = MessageBox.Show(
                    "您尚未注册本软件。\n软件使用前需要进行注册。\n\n是否现在注册？",
                    "软件注册",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);

                if (result == DialogResult.Yes)
                {
                    using (var regForm = new RegistrationForm())
                    {
                        if (regForm.ShowDialog() != DialogResult.OK)
                        {
                            MessageBox.Show("必须完成注册才能使用本软件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            Log.Information("用户取消注册，程序退出");
                            return;
                        }
                        // 重新加载设置（注册后会保存）
                        settings = ConfigService.LoadAll();
                    }
                }
                else
                {
                    MessageBox.Show("必须完成注册才能使用本软件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Log.Information("用户拒绝注册，程序退出");
                    return;
                }
            }

            try
            {
                Application.Run(new MainForm());
            }
            finally
            {
                Log.Information("=== 应用程序退出 ===");
                Log.CloseAndFlush();
            }
        }

        /// <summary>
        /// 登录日志记录器
        /// </summary>
        public static ILogger LoginLogger => new LoggerConfiguration()
            .MinimumLevel.Information()
            .WriteTo.File(
                Path.Combine(LoginLogDirectory, "login.log"),
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 7,
                outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
            .CreateLogger();

        /// <summary>
        /// 账户管理日志记录器
        /// </summary>
        public static ILogger AccountManagementLogger => new LoggerConfiguration()
            .MinimumLevel.Information()
            .WriteTo.File(
                Path.Combine(AccountManagementLogDirectory, "accountManagement.log"),
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 7,
                outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
            .CreateLogger();

        /// <summary>
        /// 组管理日志记录器
        /// </summary>
        public static ILogger GroupManagementLogger => new LoggerConfiguration()
            .MinimumLevel.Information()
            .WriteTo.File(
                Path.Combine(GroupManagementLogDirectory, "groupManagement.log"),
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 7,
                outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
            .CreateLogger();

        private static ILogger _purviewLogger;
        private static string _purviewLogDir;

        /// <summary>
        /// 获取 Purview 日志目录 (支持自定义路径)
        /// </summary>
        public static string GetPurviewLogDirectory()
        {
            if (string.IsNullOrEmpty(_purviewLogDir))
            {
                var settings = ConfigService.LoadAll();
                if (!string.IsNullOrEmpty(settings.PurviewLogPath) && Directory.Exists(settings.PurviewLogPath))
                {
                    _purviewLogDir = settings.PurviewLogPath;
                }
                else
                {
                    _purviewLogDir = PurviewLogDirectory;
                }
            }
            return _purviewLogDir;
        }

        /// <summary>
        /// PurView 日志记录器 (支持动态路径)
        /// </summary>
        public static ILogger PurviewLogger
        {
            get
            {
                if (_purviewLogger == null)
                {
                    var logDir = GetPurviewLogDirectory();
                    Directory.CreateDirectory(logDir);
                    _purviewLogger = new LoggerConfiguration()
                        .MinimumLevel.Information()
                        .WriteTo.File(
                            Path.Combine(logDir, "purview.log"),
                            rollingInterval: RollingInterval.Day,
                            retainedFileCountLimit: 7,
                            outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                        .CreateLogger();
                }
                return _purviewLogger;
            }
        }

        /// <summary>
        /// 清除 PurviewLogger 缓存 (用于重新加载设置)
        /// </summary>
        public static void ResetPurviewLogger()
        {
            _purviewLogger = null;
            _purviewLogDir = null;
        }

        private static ILogger _batchToO365Logger;

        /// <summary>
        /// 批量同步联系人到O365的日志记录器
        /// </summary>
        public static ILogger BatchToO365Logger
        {
            get
            {
                if (_batchToO365Logger == null)
                {
                    Directory.CreateDirectory(BatchToO365LogDirectory);
                    _batchToO365Logger = new LoggerConfiguration()
                        .MinimumLevel.Information()
                        .WriteTo.File(
                            Path.Combine(BatchToO365LogDirectory, "syncO365.log"),
                            rollingInterval: RollingInterval.Day,
                            retainedFileCountLimit: 7,
                            outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                        .CreateLogger();
                }
                return _batchToO365Logger;
            }
        }

        private static ILogger _registrationLogger;

        /// <summary>
        /// 启动时检查注册表中的注册信息并与云端同步
        /// 如果云端显示未注册或已过期，自动提交注册
        /// </summary>
        private static async Task CheckAndSyncRegistrationAsync()
        {
            // 检查注册表是否有注册记录
            if (!RegistryService.HasRegistrationRecord())
            {
                Log.Information("注册表中无注册记录，跳过云端同步");
                return;
            }

            // 从注册表加载注册信息
            var registrySettings = new AppSettings();
            RegistryService.LoadRegistration(registrySettings);

            if (string.IsNullOrEmpty(registrySettings.RegisteredUserEmail))
            {
                Log.Information("注册表中无邮箱信息，跳过云端同步");
                return;
            }

            Log.Information("检测到注册表记录，开始云端注册状态检查: {Email}", registrySettings.RegisteredUserEmail);

            var regService = new RegistrationService();

            // 获取MAC地址
            var macAddress = regService.GetPhysicalMacAddress();

            // 检查云端注册状态
            var checkResult = await regService.CheckRegistrationStatusAsync("xiaomingMailtoolkitCompany", registrySettings.RegisteredUserEmail);

            if (checkResult.Success)
            {
                Log.Information("云端注册状态有效，剩余天数: {Days}", checkResult.RemainingDays);

                // 更新本地设置
                var settings = ConfigService.LoadAll();
                settings.IsRegistered = true;
                settings.RegisterRemainingDays = checkResult.RemainingDays;
                settings.RegisterExpireDate = checkResult.ExpireDate;
                settings.RegisterDate = !string.IsNullOrEmpty(checkResult.InstallDate)
                    ? DateTime.TryParse(checkResult.InstallDate, out var dt) ? dt : DateTime.Now
                    : DateTime.Now;

                // 保存到注册表
                RegistryService.SaveRegistration(settings);

                // 保存到INF
                ConfigService.SaveAll(settings);

                RegistrationLogger.Information("云端注册状态同步成功，剩余天数: {Days}", checkResult.RemainingDays);
            }
            else
            {
                Log.Information("云端注册状态无效: {Message}，尝试自动注册", checkResult.Message);

                // 自动注册
                var registerResult = await regService.RegisterAsync(
                    "xiaomingMailtoolkitCompany",
                    "1.1.3",
                    registrySettings.RegisteredUserName ?? "",
                    registrySettings.RegisteredUserEmail,
                    registrySettings.RegisteredOrganization ?? "",
                    macAddress);

                if (registerResult.Success)
                {
                    Log.Information("自动注册成功，剩余天数: {Days}", registerResult.RemainingDays);

                    // 更新本地设置
                    var settings = new AppSettings
                    {
                        IsRegistered = true,
                        RegisteredUserName = registrySettings.RegisteredUserName,
                        RegisteredUserEmail = registrySettings.RegisteredUserEmail,
                        RegisteredOrganization = registrySettings.RegisteredOrganization,
                        RegisteredMacAddress = macAddress,
                        RegisterDate = !string.IsNullOrEmpty(registerResult.InstallDate)
                            ? DateTime.TryParse(registerResult.InstallDate, out var dt) ? dt : DateTime.Now
                            : DateTime.Now,
                        RegisterRemainingDays = registerResult.RemainingDays,
                        RegisterExpireDate = registerResult.ExpireDate
                    };

                    // 保存到注册表
                    RegistryService.SaveRegistration(settings);

                    // 保存到INF
                    ConfigService.SaveAll(settings);

                    RegistrationLogger.Information("自动注册成功，剩余天数: {Days}", registerResult.RemainingDays);
                }
                else
                {
                    Log.Warning("自动注册失败: {Message}", registerResult.Message);
                    RegistrationLogger.Warning("自动注册失败: {Message}", registerResult.Message);
                }
            }
        }

        /// <summary>
        /// 注册激活日志记录器
        /// </summary>
        public static ILogger RegistrationLogger
        {
            get
            {
                if (_registrationLogger == null)
                {
                    Directory.CreateDirectory(RegistrationLogDirectory);
                    _registrationLogger = new LoggerConfiguration()
                        .MinimumLevel.Information()
                        .WriteTo.File(
                            Path.Combine(RegistrationLogDirectory, "registration.log"),
                            rollingInterval: RollingInterval.Day,
                            retainedFileCountLimit: 30,
                            outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                        .CreateLogger();
                }
                return _registrationLogger;
            }
        }

        /// <summary>
        /// 日历同步日志记录器(每次同步生成新文件 calandarSync_yyyyMMdd_HHmmss.log)
        /// 日志目录: Logs/syncCalandar/
        /// </summary>
        public static ILogger CalendarSyncLogger(string sessionId = null)
        {
            Directory.CreateDirectory(SyncCalendarLogDirectory);
            var stamp = string.IsNullOrEmpty(sessionId)
                ? DateTime.Now.ToString("yyyyMMdd_HHmmss")
                : sessionId;
            return new LoggerConfiguration()
                .MinimumLevel.Information()
                .WriteTo.File(
                    Path.Combine(SyncCalendarLogDirectory, $"calandarSync_{stamp}.log"),
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] {Message:lj}{NewLine}{Exception}",
                    shared: true)
                .CreateLogger();
        }

        /// <summary>
        /// 日历 ICS 导入专用日志记录器(每次同步生成新文件 icsSync_yyyyMMdd_HHmmss.log)
        /// 日志目录: Logs/o365single/SYNCICS/
        /// 仅用于 ICS 同步(CSV/VCS 仍走 CalendarSyncLogger → Logs/syncCalandar/)
        /// </summary>
        public static ILogger CalendarIcsSyncLogger(string sessionId = null)
        {
            Directory.CreateDirectory(O365SingleSyncIcsLogDirectory);
            var stamp = string.IsNullOrEmpty(sessionId)
                ? DateTime.Now.ToString("yyyyMMdd_HHmmss")
                : sessionId;
            return new LoggerConfiguration()
                .MinimumLevel.Information()
                .WriteTo.File(
                    Path.Combine(O365SingleSyncIcsLogDirectory, $"icsSync_{stamp}.log"),
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] {Message:lj}{NewLine}{Exception}",
                    shared: true)
                .CreateLogger();
        }
    }
}
