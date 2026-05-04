using System;
using System.IO;

namespace MailConverter
{
    /// <summary>
    /// 邮件提取日志服务 - 为邮件提取功能创建独立日志文件
    /// </summary>
    public static class MailPickLogService
    {
        private static readonly string LogDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs", "mailpick");
        private static string _currentLogType = "EMLPSTImport";

        public static string[] LogTypes = new[] { "EMLPSTImport", "SyncContactCalendar" };

        /// <summary>
        /// 获取指定类型的日志路径
        /// </summary>
        public static string GetLogPath(string type)
        {
            var dir = Path.Combine(LogDir, type);
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, $"{type.ToUpper()}-{DateTime.Now:yyyy-MM-dd}.LOG");
        }

        /// <summary>
        /// 获取当前日志类型的路径
        /// </summary>
        public static string GetCurrentLogPath()
        {
            return GetLogPath(_currentLogType);
        }

        /// <summary>
        /// 写入日志
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
                Serilog.Log.Warning(ex, "写入邮件提取日志失败");
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
