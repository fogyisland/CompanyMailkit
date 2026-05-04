using System;
using System.IO;

namespace MailConverter
{
    /// <summary>
    /// PST转换日志服务 - 为每种转换类型创建独立日志文件
    /// </summary>
    public static class TransferPstLogService
    {
        private static readonly string LogDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs", "transferPST");
        private static string _currentLogType = "EML2PST";

        public static string[] LogTypes = new[] { "EML2PST", "OST2PST", "IMAP2PST", "IMAPMULTI2PST", "SingleUserSYNCO365", "BatchSYNCtoO365" };

        // BatchSYNCtoO365 子类型
        public static string[] BatchLogTypes = new[] { "Login", "PSTMAILSYNC", "PSTSYNCContact", "PSTSYNCCalendar", "CSVSYNCContact", "VCSYNCContact", "CSVSYNCCalendar", "PurViewSYNC" };
        // SingleUserSYNCO365 子类型
        public static string[] SingleUserLogTypes = new[] { "EMLPSTImport", "SyncContactCalendar" };
        // O365Toolkit 子类型
        public static string[] O365ToolkitLogTypes = new[] { "O365Login", "O365Account", "O365Group", "O365Mobile", "O365Traffic", "O365Migration", "O365Whois", "O365Dns" };

        /// <summary>
        /// 获取指定类型的日志路径
        /// </summary>
        public static string GetLogPath(string type)
        {
            string dir;
            if (type == "SingleUserSYNCO365")
            {
                // SingleUserSYNCO365 放在 logs/SingleUserSYNCO365/ 下（作为父目录）
                dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs", "SingleUserSYNCO365");
            }
            else if (Array.Exists(SingleUserLogTypes, t => t == type))
            {
                // SingleUserSYNCO365 子类型放在 logs/SingleUserSYNCO365/ 下
                dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs", "SingleUserSYNCO365", type);
            }
            else if (Array.Exists(BatchLogTypes, t => t == type))
            {
                // BatchSYNCtoO365 子类型放在 logs/BatchSYNCtoO365/ 下
                dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs", "BatchSYNCtoO365", type);
            }
            else if (type == "O365Toolkit")
            {
                // O365Toolkit 作为父目录
                dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs", "O365Toolkit");
            }
            else if (Array.Exists(O365ToolkitLogTypes, t => t == type))
            {
                // O365Toolkit 子类型放在 logs/O365Toolkit/ 下
                dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs", "O365Toolkit", type);
            }
            else
            {
                dir = Path.Combine(LogDir, type);
            }
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, $"{type}-{DateTime.Now:yyyy-MM-dd}.LOG");
        }

        /// <summary>
        /// 获取当前日志类型的路径
        /// </summary>
        public static string GetCurrentLogPath()
        {
            return GetLogPath(_currentLogType);
        }

        /// <summary>
        /// 写入日志（同时写入文件和显示到UI）
        /// </summary>
        public static void Log(string type, string message)
        {
            try
            {
                var logPath = GetLogPath(type);
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\n";
                File.AppendAllText(logPath, logEntry, System.Text.Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Serilog.Log.Warning(ex, "写入PST转换日志失败");
            }
        }

        /// <summary>
        /// 设置当前日志类型
        /// </summary>
        public static void SetCurrentLogType(string type)
        {
            _currentLogType = type;
        }

        /// <summary>
        /// 获取当前日志类型
        /// </summary>
        public static string GetCurrentLogType()
        {
            return _currentLogType;
        }

        /// <summary>
        /// 读取指定类型的最新日志内容
        /// </summary>
        public static string ReadLog(string type)
        {
            try
            {
                var logPath = GetLogPath(type);
                if (File.Exists(logPath))
                {
                    return File.ReadAllText(logPath, System.Text.Encoding.UTF8);
                }
            }
            catch { }
            return string.Empty;
        }

        /// <summary>
        /// 读取当前日志类型的日志
        /// </summary>
        public static string ReadCurrentLog()
        {
            return ReadLog(_currentLogType);
        }
    }
}
