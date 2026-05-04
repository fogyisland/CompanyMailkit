using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using Serilog;

namespace MailConverter
{
    /// <summary>
    /// Exchange On-Premise 管理工具箱
    /// </summary>
    public class ExchangeOnPremiseToolkitService
    {
        // 批量导出PST专用日志
        private static readonly ILogger _batchExportLogger = new LoggerConfiguration()
            .WriteTo.File(
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs", "batchEX2Pst", "batchExport.log"),
                rollingInterval: RollingInterval.Day,
                outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
            .CreateLogger();

        // 批量投递到IMAP专用日志
        private static readonly ILogger _batchToImapLogger = new LoggerConfiguration()
            .WriteTo.File(
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs", "batchToImap", "batchToImap.log"),
                rollingInterval: RollingInterval.Day,
                outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
            .CreateLogger();

        private ExchangeService _service;
        private string _email;
        private string _password;
        private string _serverUrl;
        private string _domain;

        /// <summary>
        /// 连接到 Exchange On-Premise
        /// </summary>
        public bool Connect(string email, string password, string serverUrl, string domain = null)
        {
            try
            {
                _email = email;
                _password = password;
                _serverUrl = serverUrl;
                _domain = domain;

                Log.Information("连接 Exchange On-Premise: {Email} @ {Server}", email, serverUrl);

                _service = new ExchangeService(ExchangeVersion.Exchange2016);

                if (!string.IsNullOrEmpty(serverUrl))
                {
                    // 确保URL包含正确的EWS路径
                    string ewsUrl = serverUrl;
                    if (serverUrl.IndexOf("/EWS/Exchange.asmx", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        // 移除末尾的斜杠（如果有）
                        ewsUrl = serverUrl.TrimEnd('/');
                        // 添加EWS路径
                        if (!ewsUrl.EndsWith("/EWS", StringComparison.OrdinalIgnoreCase))
                            ewsUrl += "/EWS";
                        if (!ewsUrl.EndsWith("/Exchange.asmx", StringComparison.OrdinalIgnoreCase))
                            ewsUrl += "/Exchange.asmx";
                    }
                    Log.Information("EWS URL: {EwsUrl}", ewsUrl);
                    _service.Url = new Uri(ewsUrl);
                }

                if (!string.IsNullOrEmpty(domain))
                {
                    _service.Credentials = new System.Net.NetworkCredential(email, password, domain);
                }
                else
                {
                    _service.Credentials = new System.Net.NetworkCredential(email, password);
                }

                // 尝试自动发现
                try { _service.AutodiscoverUrl(email, RedirectionUrlValidationCallback); }
                catch { Log.Warning("自动发现失败，请手动指定EWS地址"); }

                // 验证连接
                Folder.Bind(_service, WellKnownFolderName.Inbox);

                Log.Information("Exchange On-Premise 连接成功");
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Exchange On-Premise 连接失败: {Msg}", ex.Message);
                return false;
            }
        }

        private bool RedirectionUrlValidationCallback(string url)
        {
            return url.ToLower().StartsWith("https://");
        }

        /// <summary>
        /// 发送测试邮件
        /// </summary>
        public bool SendTestEmail(string to, string subject, string body)
        {
            try
            {
                var email = new EmailMessage(_service)
                {
                    Subject = subject,
                    Body = new MessageBody(BodyType.Text, body)
                };
                email.ToRecipients.Add(to);
                email.Send();
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "发送测试邮件失败");
                return false;
            }
        }

        /// <summary>
        /// 设置要模拟的用户（Impersonation）
        /// </summary>
        public bool SetImpersonatedUser(string smtpAddress)
        {
            try
            {
                if (_service == null)
                {
                    Log.Error("服务未初始化，请先调用Connect方法");
                    return false;
                }

                _service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, smtpAddress);
                Log.Information("已设置Impersonation用户: {User}", smtpAddress);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "设置Impersonation用户失败: {User}", smtpAddress);
                return false;
            }
        }

        /// <summary>
        /// 获取邮件的MimeContent（字符串）
        /// </summary>
        public string GetMimeContent(string itemId)
        {
            try
            {
                var item = Item.Bind(_service, itemId, new PropertySet(EmailMessageSchema.MimeContent)).GetAwaiter().GetResult();
                return item.MimeContent?.Content != null
                    ? new System.Text.UTF8Encoding(false).GetString(item.MimeContent.Content)
                    : null;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取MimeContent失败: {Msg}", ex.Message);
                return null;
            }
        }

        /// <summary>
        /// 清理文件名中的非法字符
        /// </summary>
        private string SanitizeFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return "无主题";
            var invalid = Path.GetInvalidFileNameChars();
            foreach (var c in invalid)
            {
                fileName = fileName.Replace(c, '_');
            }
            if (fileName.Length > 100) fileName = fileName.Substring(0, 100);
            return fileName;
        }

        /// <summary>
        /// 清理文件夹名称中的非法字符
        /// </summary>
        private string SanitizeFolderName(string folderName)
        {
            if (string.IsNullOrEmpty(folderName)) return "未知文件夹";
            var invalid = Path.GetInvalidFileNameChars();
            foreach (var c in invalid)
            {
                folderName = folderName.Replace(c, '_');
            }
            folderName = folderName.Replace(":", "_").Replace("*", "_").Replace("?", "_")
                .Replace("\"", "_").Replace("<", "_").Replace(">", "_").Replace("|", "_");
            return folderName;
        }

        /// <summary>
        /// 批量导出PST（EWS提取EML + Outlook创建PST）
        /// CSV格式：用户邮箱,PST文件名
        /// </summary>
        public bool BatchExportPst(
            string csvPath,
            string outputDir,
            DateTime startDate,
            DateTime endDate,
            int maxDegreeOfParallelism,
            bool exportEmail,
            bool exportCalendar,
            bool exportContacts,
            bool exportTasks,
            Action<int, int, string, string> progressCallback)
        {
            try
            {
                // 确保日志目录存在
                var logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs", "batchEX2Pst");
                Directory.CreateDirectory(logDir);

                if (!File.Exists(csvPath))
                {
                    _batchExportLogger.Error("CSV文件不存在: {Path}", csvPath);
                    return false;
                }

                _batchExportLogger.Information("========== 批量导出PST开始 ==========");
                _batchExportLogger.Information("CSV文件: {Path}", csvPath);
                _batchExportLogger.Information("输出目录: {OutputDir}", outputDir);
                _batchExportLogger.Information("日期范围: {Start} 至 {End}", startDate, endDate);
                _batchExportLogger.Information("导出选项: 邮件={Email}, 日历={Calendar}, 联系人={Contacts}, 任务={Tasks}",
                    exportEmail, exportCalendar, exportContacts, exportTasks);

                // 读取CSV文件
                var entries = new List<Tuple<string, string>>();
                var lines = File.ReadAllLines(csvPath, System.Text.Encoding.UTF8);
                foreach (var line in lines.Skip(1))
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    var parts = line.Split(',');
                    if (parts.Length >= 2)
                    {
                        var email = parts[0].Trim().Trim('"');
                        var pstName = parts[1].Trim().Trim('"');
                        if (!string.IsNullOrEmpty(email) && !string.IsNullOrEmpty(pstName))
                        {
                            entries.Add(Tuple.Create(email, pstName));
                        }
                    }
                }

                if (entries.Count == 0)
                {
                    _batchExportLogger.Error("CSV文件中没有有效数据");
                    return false;
                }

                _batchExportLogger.Information("开始批量导出PST，共 {Count} 个邮箱", entries.Count);

                int total = entries.Count;
                int completed = 0;
                int errors = 0;

                // EWS Impersonation不支持多线程并行，每个用户必须串行处理
                foreach (var entry in entries)
                {
                    string userEmail = entry.Item1;
                    string pstName = entry.Item2;
                    string tempDir = null;
                    bool success = false;

                    try
                    {
                        tempDir = Path.Combine(Path.GetTempPath(), $"batch_pst_{Guid.NewGuid():N}");
                        Directory.CreateDirectory(tempDir);

                        progressCallback?.Invoke(completed, total, userEmail, "正在提取EML...");

                        int totalExported = 0;
                        int totalErrors = 0;

                        // 1. 设置Impersonation
                        SetImpersonatedUser(userEmail);

                        // 2. 生成PST文件路径
                        string pstPath = Path.Combine(outputDir, pstName);
                        if (!pstName.EndsWith(".pst", StringComparison.OrdinalIgnoreCase))
                            pstPath += ".pst";

                        // 3. 导出邮件文件夹 (第一步)
                        if (exportEmail)
                        {
                            progressCallback?.Invoke(completed, total, userEmail, "正在提取邮件...");
                            var (exported, errCount) = ExportAllFoldersToEml(tempDir, startDate, endDate);
                            totalExported += exported;
                            totalErrors += errCount;
                        }

                        // 4. 调用Python脚本创建PST (第二步)
                        bool hasEmails = Directory.GetFiles(tempDir, "*.eml", SearchOption.AllDirectories).Length > 0;
                        if (hasEmails)
                        {
                            progressCallback?.Invoke(completed, total, userEmail, "正在创建PST...");

                            var exeDir = AppDomain.CurrentDomain.BaseDirectory;
                            var pythonExe = Path.Combine(exeDir, "python", "python.exe");
                            var scriptPath = Path.Combine(exeDir, "script", "create_pst.py");

                            if (!File.Exists(pythonExe))
                            {
                                throw new Exception("Python环境未找到");
                            }
                            if (!File.Exists(scriptPath))
                            {
                                throw new Exception("create_pst.py脚本不存在");
                            }

                            var args = $"\"{scriptPath}\" \"{pstPath}\" \"{tempDir}\"";
                            var startInfo = new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = pythonExe,
                                Arguments = args,
                                UseShellExecute = false,
                                RedirectStandardOutput = true,
                                RedirectStandardError = true,
                                CreateNoWindow = true,
                                StandardOutputEncoding = System.Text.Encoding.UTF8,
                                StandardErrorEncoding = System.Text.Encoding.UTF8
                            };
                            startInfo.EnvironmentVariables["PYTHONNOUSERSITE"] = "1";
                            startInfo.EnvironmentVariables["PYTHONIOENCODING"] = "UTF-8";

                            using (var process = System.Diagnostics.Process.Start(startInfo))
                            {
                                if (process != null)
                                {
                                    process.WaitForExit(300000);
                                    if (process.ExitCode == 0)
                                    {
                                        success = true;
                                        _batchExportLogger.Information("用户 {Email} PST邮件导入成功: {Path}", userEmail, pstPath);
                                    }
                                    else
                                    {
                                        string err = process.StandardError.ReadToEnd();
                                        _batchExportLogger.Error("用户 {Email} PST创建失败: {Error}", userEmail, err);
                                    }
                                }
                            }
                        }

                        // 5. 提取联系人和日历 (第三步 - 在PST创建之后)
                        if (success)
                        {
                            string dataDir = Path.Combine(Path.GetTempPath(), $"batch_data_{Guid.NewGuid():N}");
                            Directory.CreateDirectory(dataDir);

                            try
                            {
                                // 导出日历
                                if (exportCalendar)
                                {
                                    progressCallback?.Invoke(completed, total, userEmail, "正在提取日历...");
                                    var (exported, errCount) = ExportCalendarToJson(dataDir, startDate, endDate);
                                    totalExported += exported;
                                    totalErrors += errCount;
                                }

                                // 导出联系人
                                if (exportContacts)
                                {
                                    progressCallback?.Invoke(completed, total, userEmail, "正在提取联系人...");
                                    var (exported, errCount) = ExportContactsToJson(dataDir);
                                    totalExported += exported;
                                    totalErrors += errCount;
                                }

                                // 6. 写入日历和联系人到PST (第四步)
                                bool hasCalendarJson = File.Exists(Path.Combine(dataDir, "日历", "calendar.json"));
                                bool hasContactsJson = File.Exists(Path.Combine(dataDir, "联系人", "contacts.json"));

                                if (hasCalendarJson || hasContactsJson)
                                {
                                    progressCallback?.Invoke(completed, total, userEmail, "正在导入日历和联系人...");
                                    try
                                    {
                                        ImportCalendarAndContactsToPst(pstPath, dataDir);
                                        _batchExportLogger.Information("用户 {Email} 日历和联系人导入成功", userEmail);
                                    }
                                    catch (Exception ex)
                                    {
                                        _batchExportLogger.Error(ex, "用户 {Email} 日历和联系人导入失败", userEmail);
                                    }
                                }
                            }
                            finally
                            {
                                // 清理临时数据目录
                                try { Directory.Delete(dataDir, true); } catch { }
                            }
                        }

                        // 7. 导出任务
                        if (exportTasks)
                        {
                            progressCallback?.Invoke(completed, total, userEmail, "正在提取任务...");
                            var (exported, errCount) = ExportTasksToEml(tempDir, startDate, endDate);
                            totalExported += exported;
                            totalErrors += errCount;
                        }

                        if (totalExported == 0 && !success)
                        {
                            _batchExportLogger.Warning("用户 {Email} 没有内容可导出", userEmail);
                            progressCallback?.Invoke(completed, total, userEmail, "无内容");
                        }
                    }
                    catch (Exception ex)
                    {
                        errors++;
                        _batchExportLogger.Error(ex, "用户 {Email} 批量导出失败", userEmail);
                    }
                    finally
                    {
                        try
                        {
                            if (tempDir != null && Directory.Exists(tempDir))
                            {
                                Directory.Delete(tempDir, true);
                            }
                        }
                        catch { }

                        completed++;
                        progressCallback?.Invoke(completed, total, userEmail, success ? "成功" : "失败");
                    }
                }

                _batchExportLogger.Information("批量导出PST完成: 成功 {Success}, 失败 {Errors}", total - errors, errors);
                _batchExportLogger.Information("========== 批量导出PST结束 ==========");
                return errors == 0;
            }
            catch (Exception ex)
            {
                _batchExportLogger.Error(ex, "批量导出PST失败");
                return false;
            }
        }

        /// <summary>
        /// 批量投递到IMAP
        /// CSV格式：源邮箱,目标邮箱,目标密码
        /// </summary>
        public bool BatchDeliverToImap(
            string imapServer,
            int imapPort,
            bool useSsl,
            string csvPath,
            DateTime startDate,
            DateTime endDate,
            bool includeReceived,
            bool includeSent,
            int maxDegreeOfParallelism,
            Action<int, int, string, string> progressCallback)
        {
            try
            {
                var logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs", "batchToImap");
                Directory.CreateDirectory(logDir);

                if (!File.Exists(csvPath))
                {
                    _batchToImapLogger.Error("CSV文件不存在: {Path}", csvPath);
                    return false;
                }

                _batchToImapLogger.Information("========== 批量投递到IMAP开始 ==========");
                _batchToImapLogger.Information("IMAP服务器: {Server}:{Port}, SSL: {Ssl}", imapServer, imapPort, useSsl);
                _batchToImapLogger.Information("CSV文件: {Path}", csvPath);
                _batchToImapLogger.Information("日期范围: {Start} 至 {End}", startDate, endDate);

                // 读取CSV文件
                var entries = new List<Tuple<string, string, string>>();
                var lines = File.ReadAllLines(csvPath, System.Text.Encoding.UTF8);
                foreach (var line in lines.Skip(1))
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    var parts = line.Split(',');
                    if (parts.Length >= 3)
                    {
                        var sourceEmail = parts[0].Trim().Trim('"');
                        var targetEmail = parts[1].Trim().Trim('"');
                        var targetPassword = parts[2].Trim().Trim('"');
                        if (!string.IsNullOrEmpty(sourceEmail) && !string.IsNullOrEmpty(targetEmail))
                        {
                            entries.Add(Tuple.Create(sourceEmail, targetEmail, targetPassword));
                        }
                    }
                }

                if (entries.Count == 0)
                {
                    _batchToImapLogger.Error("CSV文件中没有有效数据");
                    return false;
                }

                _batchToImapLogger.Information("共 {Count} 个账户待处理", entries.Count);

                int total = entries.Count;
                int completed = 0;
                int errors = 0;

                foreach (var entry in entries)
                {
                    string sourceEmail = entry.Item1;
                    string targetEmail = entry.Item2;
                    string targetPassword = entry.Item3;

                    try
                    {
                        progressCallback?.Invoke(completed, total, sourceEmail, "正在处理...");

                        // 设置Impersonation
                        if (!SetImpersonatedUser(sourceEmail))
                        {
                            _batchToImapLogger.Error("设置Impersonation用户失败: {Email}", sourceEmail);
                            errors++;
                            completed++;
                            continue;
                        }

                        // 获取邮箱文件夹
                        var folders = GetAllFolders();
                        _batchToImapLogger.Information("用户 {Email} 共发现 {Count} 个文件夹", sourceEmail, folders.Count);
                        foreach (var f in folders)
                        {
                            _batchToImapLogger.Information("  文件夹: {Name}, Count={TotalCount}, Class={Class}", f.DisplayName, f.TotalCount, f.FolderClass);
                        }

                        // 收集所有邮件
                        var allEmails = new List<Tuple<string, byte[], DateTime>>();
                        foreach (var folder in folders)
                        {
                            bool isInbox = folder.DisplayName == "收件箱" || folder.DisplayName == "Inbox";
                            bool isSent = folder.DisplayName == "已发送" || folder.DisplayName == "已发送邮件" || folder.DisplayName == "Sent" || folder.DisplayName == "Sent Items" || folder.DisplayName == "发件箱";

                            _batchToImapLogger.Information("用户 {Email} 检查文件夹 {Folder}: isInbox={Inbox}, isSent={Sent}, includeRecv={Recv}, includeSent={Sent2}",
                                sourceEmail, folder.DisplayName, isInbox, isSent, includeReceived, includeSent);

                            if ((includeReceived && isInbox) || (includeSent && isSent))
                            {
                                var emails = GetFolderEmailsMime(folder.Id, startDate, endDate);
                                allEmails.AddRange(emails);
                                _batchToImapLogger.Information("用户 {Email} 文件夹 {Folder}: 获取 {Count} 封", sourceEmail, folder.DisplayName, emails.Count);
                            }
                        }

                        if (allEmails.Count == 0)
                        {
                            _batchToImapLogger.Warning("用户 {Email} 没有邮件可投递", sourceEmail);
                            progressCallback?.Invoke(completed + 1, total, sourceEmail, "无邮件");
                            completed++;
                            continue;
                        }

                        progressCallback?.Invoke(completed, total, sourceEmail, $"正在投递 {allEmails.Count} 封邮件...");

                        // 连接到IMAP服务器
                        using (var imapClient = new MailKit.Net.Imap.ImapClient())
                        {
                            imapClient.Connect(imapServer, imapPort, useSsl ? MailKit.Security.SecureSocketOptions.SslOnConnect : MailKit.Security.SecureSocketOptions.None);
                            imapClient.Authenticate(targetEmail, targetPassword);

                            var inbox = imapClient.Inbox;
                            int delivered = 0;
                            int failed = 0;

                            foreach (var email in allEmails)
                            {
                                try
                                {
                                    using (var stream = new MemoryStream(email.Item2))
                                    {
                                        var message = MimeKit.MimeMessage.Load(stream);
                                        var request = new MailKit.AppendRequest(message, MailKit.MessageFlags.None)
                                        {
                                            InternalDate = new DateTimeOffset(email.Item3)
                                        };
                                        inbox.Append(request, CancellationToken.None);
                                        delivered++;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    failed++;
                                    _batchToImapLogger.Error(ex, "投递邮件失败: {Subject}", email.Item1);
                                }
                            }

                            imapClient.Disconnect(true);
                            _batchToImapLogger.Information("用户 {Email} 投递完成: 成功 {Success}, 失败 {Failed}", sourceEmail, delivered, failed);
                        }

                        completed++;
                        progressCallback?.Invoke(completed, total, sourceEmail, "完成");
                    }
                    catch (Exception ex)
                    {
                        errors++;
                        completed++;
                        _batchToImapLogger.Error(ex, "用户 {Email} 处理失败", sourceEmail);
                        progressCallback?.Invoke(completed, total, sourceEmail, "失败: " + ex.Message);
                    }
                }

                _batchToImapLogger.Information("批量投递到IMAP完成: 成功 {Success}, 失败 {Errors}", total - errors, errors);
                _batchToImapLogger.Information("========== 批量投递到IMAP结束 ==========");
                return errors == 0;
            }
            catch (Exception ex)
            {
                _batchToImapLogger.Error(ex, "批量投递到IMAP失败");
                return false;
            }
        }

        /// <summary>
        /// 导出邮箱所有文件夹到EML（保留文件夹结构）
        /// </summary>
        private (int exported, int errors) ExportAllFoldersToEml(
            string outputDir,
            DateTime startDate,
            DateTime endDate)
        {
            int totalExported = 0;
            int totalErrors = 0;

            try
            {
                // 使用FolderView深度遍历获取所有文件夹
                var view = new FolderView(500);
                view.Traversal = FolderTraversal.Deep;
                view.PropertySet = new PropertySet(
                    BasePropertySet.IdOnly,
                    FolderSchema.DisplayName,
                    FolderSchema.FolderClass,
                    FolderSchema.TotalCount);

                var results = _service.FindFolders(WellKnownFolderName.MsgFolderRoot, view).GetAwaiter().GetResult();

                foreach (Folder folder in results.Folders)
                {
                    try
                    {
                        if (folder is SearchFolder) continue;

                        string folderClass = folder.FolderClass ?? "";
                        // 只导出邮件文件夹 (IPF.Note) 和通用文件夹 (IPF.)
                        // 跳过所有非邮件类型的系统文件夹
                        if (folderClass != "IPF.Note" && folderClass != "" && !folderClass.StartsWith("IPF."))
                            continue;

                        string folderName = folder.DisplayName ?? "";

                        // 跳过根目录和特殊系统文件夹
                        if (folderName == "Root" || folderName == "MsgFolderRoot")
                            continue;

                        // 跳过GUID命名的隐藏文件夹（如 {06967759-...}）
                        if (folderName.StartsWith("{") && folderName.EndsWith("}"))
                            continue;

                        // 跳过隐藏的通讯录/缓存文件夹
                        string lowerName = folderName.ToLower();
                        if (lowerName.Contains("recipient cache") ||
                            lowerName.Contains("gal contacts") ||
                            lowerName.Contains("externalcontacts") ||
                            lowerName.Contains("peoplecentric") ||
                            lowerName.Contains("conversation action") ||
                            lowerName.Contains("quick step") ||
                            lowerName.Contains("sync issues") ||
                            lowerName.Contains("rss") ||
                            lowerName.Contains("acl") ||
                            lowerName.Contains("access control"))
                            continue;

                        string folderOutputDir = Path.Combine(outputDir, SanitizeFolderName(folderName));
                        if (!Directory.Exists(folderOutputDir))
                        {
                            Directory.CreateDirectory(folderOutputDir);
                        }

                        // 导出该文件夹中的邮件
                        var searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                            new SearchFilter.IsGreaterThanOrEqualTo(EmailMessageSchema.DateTimeReceived, startDate),
                            new SearchFilter.IsLessThanOrEqualTo(EmailMessageSchema.DateTimeReceived, endDate));

                        var itemView = new ItemView(1000);
                        itemView.PropertySet = new PropertySet(BasePropertySet.IdOnly, EmailMessageSchema.Subject, EmailMessageSchema.DateTimeReceived);

                        var itemResults = _service.FindItems(folder.Id, searchFilter, itemView).GetAwaiter().GetResult();

                        foreach (var item in itemResults)
                        {
                            try
                            {
                                var mimeContent = GetMimeContent(item.Id.UniqueId);
                                if (!string.IsNullOrEmpty(mimeContent))
                                {
                                    string safeSubject = SanitizeFileName(item.Subject ?? "无主题");
                                    string fileName = $"{safeSubject}.eml";
                                    string filePath = Path.Combine(folderOutputDir, fileName);

                                    int counter = 1;
                                    while (File.Exists(filePath))
                                    {
                                        fileName = $"{safeSubject}_{counter}.eml";
                                        filePath = Path.Combine(folderOutputDir, fileName);
                                        counter++;
                                    }

                                    File.WriteAllText(filePath, mimeContent, new System.Text.UTF8Encoding(false));
                                    totalExported++;
                                }
                                else
                                {
                                    totalErrors++;
                                }
                            }
                            catch (Exception ex)
                            {
                                totalErrors++;
                                _batchExportLogger.Error(ex, "导出邮件失败: {Subject}", item.Subject ?? "(无主题)");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _batchExportLogger.Error(ex, "导出文件夹失败: {Folder}", folder.DisplayName);
                    }
                }

                _batchExportLogger.Information("导出所有文件夹完成: 共导出 {Exported} 封, 失败 {Errors} 封", totalExported, totalErrors);
            }
            catch (Exception ex)
            {
                _batchExportLogger.Error(ex, "导出所有文件夹失败");
            }

            return (totalExported, totalErrors);
        }

        /// <summary>
        /// 导出日历到ICS文件
        /// </summary>
        private (int exported, int errors) ExportCalendarToJson(
            string outputDir,
            DateTime startDate,
            DateTime endDate)
        {
            int exported = 0;
            int errors = 0;

            try
            {
                _batchExportLogger.Information("开始导出日历(JSON): startDate={Start}, endDate={End}", startDate, endDate);

                var calendarFolder = CalendarFolder.Bind(_service, WellKnownFolderName.Calendar).GetAwaiter().GetResult();
                _batchExportLogger.Information("日历文件夹绑定成功");

                var adjustedEndDate = endDate.AddDays(1).AddSeconds(-1);
                if ((adjustedEndDate - startDate).TotalDays > 730)
                {
                    adjustedEndDate = startDate.AddDays(729);
                }
                var calendarView = new CalendarView(startDate, adjustedEndDate, 1000);

                var appointments = calendarFolder.FindAppointments(calendarView).GetAwaiter().GetResult();
                _batchExportLogger.Information("找到 {Count} 个日历事件", appointments.Count());

                string folderOutputDir = Path.Combine(outputDir, "日历");
                Directory.CreateDirectory(folderOutputDir);

                // 创建日历属性集：FirstClassProperties + Body（强制Text格式）
                var calendarPropertySet = new PropertySet(BasePropertySet.FirstClassProperties, AppointmentSchema.Body);
                calendarPropertySet.RequestedBodyType = BodyType.Text;

                var calendarItems = new System.Collections.Generic.List<object>();
                int calProcessed = 0;

                foreach (var appointment in appointments)
                {
                    try
                    {
                        appointment.Load(calendarPropertySet);

                        string bodyText = "";
                        try { bodyText = appointment.Body?.Text ?? ""; } catch { }

                        calendarItems.Add(new
                        {
                            Subject = appointment.Subject ?? "",
                            Start = appointment.Start.ToString("yyyy-MM-ddTHH:mm:ss"),
                            End = appointment.End.ToString("yyyy-MM-ddTHH:mm:ss"),
                            Location = appointment.Location ?? "",
                            IsAllDayEvent = appointment.IsAllDayEvent,
                            Body = bodyText,
                            Organizer = appointment.Organizer?.Name ?? "",
                            IsRecurring = appointment.IsRecurring,
                            IsMeeting = appointment.IsMeeting,
                            IsCancelled = appointment.IsCancelled,
                            Categories = appointment.Categories != null ? string.Join(",", appointment.Categories) : ""
                        });
                        exported++;
                        calProcessed++;
                        if (calProcessed % 5 == 0)
                        {
                            _batchExportLogger.Information("处理日历进度: {Processed}/{Total}", calProcessed, appointments.Count());
                        }
                    }
                    catch (Exception ex)
                    {
                        errors++;
                        _batchExportLogger.Error(ex, "导出日历事件失败: {Subject}", appointment.Subject ?? "(无主题)");
                    }
                }

                // 写入JSON文件
                if (calendarItems.Count > 0)
                {
                    string jsonPath = Path.Combine(folderOutputDir, "calendar.json");
                    var jsonContent = System.Text.Json.JsonSerializer.Serialize(calendarItems);
                    File.WriteAllText(jsonPath, jsonContent, new System.Text.UTF8Encoding(false));
                    _batchExportLogger.Information("日历JSON已保存: {Count} 条记录", exported);
                }

                _batchExportLogger.Information("日历导出完成: 成功={Exported}, 失败={Errors}", exported, errors);
            }
            catch (Exception ex)
            {
                _batchExportLogger.Error(ex, "导出日历失败");
            }

            return (exported, errors);
        }

        /// <summary>
        /// 生成ICS日历内容
        /// </summary>
        private string GenerateIcsContent(Appointment appointment)
        {
            try
            {
                var sb = new System.Text.StringBuilder();
                sb.AppendLine("BEGIN:VCALENDAR");
                sb.AppendLine("VERSION:2.0");
                sb.AppendLine("PRODID:-//MailConverter//NONSGML v1.0//EN");
                sb.AppendLine("BEGIN:VEVENT");
                sb.AppendLine($"UID:{appointment.Id?.UniqueId ?? Guid.NewGuid().ToString()}");

                string dtStart = appointment.Start.ToUniversalTime().ToString("yyyyMMddTHHm:ssZ");
                string dtEnd = appointment.End.ToUniversalTime().ToString("yyyyMMddTHHmmssZ");
                sb.AppendLine($"DTSTART:{dtStart}");
                sb.AppendLine($"DTEND:{dtEnd}");
                sb.AppendLine($"SUMMARY:{EscapeIcsText(appointment.Subject ?? "")}");
                sb.AppendLine($"LOCATION:{EscapeIcsText(appointment.Location ?? "")}");
                sb.AppendLine($"DESCRIPTION:{EscapeIcsText(appointment.Subject ?? "")}");
                sb.AppendLine($"DTSTAMP:{DateTime.UtcNow.ToString("yyyyMMddTHHmmssZ")}");
                sb.AppendLine("END:VEVENT");
                sb.AppendLine("END:VCALENDAR");

                return sb.ToString();
            }
            catch
            {
                return null;
            }
        }

        private string EscapeIcs(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            return value.Replace("\\", "\\\\").Replace(",", "\\,").Replace(";", "\\;").Replace("\n", "\\n");
        }

        private string EscapeIcsText(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            return value.Replace("\\", "\\\\").Replace(",", "\\,").Replace(";", "\\;").Replace("\n", "\\n").Replace("\r", "");
        }

        /// <summary>
        /// 导出联系人到JSON文件
        /// </summary>
        private (int exported, int errors) ExportContactsToJson(string outputDir)
        {
            int exported = 0;
            int errors = 0;

            try
            {
                _batchExportLogger.Information("开始导出联系人(JSON)");

                var contactsFolder = ContactsFolder.Bind(_service, WellKnownFolderName.Contacts).GetAwaiter().GetResult();
                _batchExportLogger.Information("联系人文件夹绑定成功");

                // 关键：搜索阶段只拿ID，不要拿任何复杂属性
                var view = new ItemView(1000);
                view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ContactSchema.DisplayName);

                var results = _service.FindItems(contactsFolder.Id, view).GetAwaiter().GetResult();
                _batchExportLogger.Information("FindItems返回 {Count} 个结果", results.Count());

                string folderOutputDir = Path.Combine(outputDir, "联系人");
                Directory.CreateDirectory(folderOutputDir);

                // 批量加载联系人属性（FirstClassProperties已包含所有常用属性）
                if (results.Items.Count > 0)
                {
                    var contactPropertySet = new PropertySet(BasePropertySet.FirstClassProperties);
                    _service.LoadPropertiesForItems(results.Items, contactPropertySet);
                    _batchExportLogger.Information("批量加载联系人属性完成，共 {Count} 个", results.Items.Count);
                }

                var contactItems = new System.Collections.Generic.List<object>();
                int contactProcessed = 0;

                foreach (var item in results)
                {
                    try
                    {
                        if (item is Contact contact)
                        {
                            contactProcessed++;
                            if (contactProcessed % 10 == 0)
                            {
                                _batchExportLogger.Information("处理联系人进度: {Processed}/{Total}", contactProcessed, results.Items.Count);
                            }

                            string email1 = "", email2 = "", email3 = "";
                            if (contact.EmailAddresses != null)
                            {
                                if (contact.EmailAddresses.Contains(EmailAddressKey.EmailAddress1))
                                    email1 = contact.EmailAddresses[EmailAddressKey.EmailAddress1]?.Address ?? "";
                                if (contact.EmailAddresses.Contains(EmailAddressKey.EmailAddress2))
                                    email2 = contact.EmailAddresses[EmailAddressKey.EmailAddress2]?.Address ?? "";
                                if (contact.EmailAddresses.Contains(EmailAddressKey.EmailAddress3))
                                    email3 = contact.EmailAddresses[EmailAddressKey.EmailAddress3]?.Address ?? "";
                            }

                            string mobilePhone = "", homePhone = "", businessPhone = "";
                            if (contact.PhoneNumbers != null)
                            {
                                if (contact.PhoneNumbers.Contains(PhoneNumberKey.MobilePhone))
                                    mobilePhone = contact.PhoneNumbers[PhoneNumberKey.MobilePhone] ?? "";
                                if (contact.PhoneNumbers.Contains(PhoneNumberKey.HomePhone))
                                    homePhone = contact.PhoneNumbers[PhoneNumberKey.HomePhone] ?? "";
                                if (contact.PhoneNumbers.Contains(PhoneNumberKey.BusinessPhone))
                                    businessPhone = contact.PhoneNumbers[PhoneNumberKey.BusinessPhone] ?? "";
                            }

                            contactItems.Add(new
                            {
                                DisplayName = contact.DisplayName ?? "",
                                FirstName = contact.GivenName ?? "",
                                LastName = contact.Surname ?? "",
                                Email1 = email1,
                                Email2 = email2,
                                Email3 = email3,
                                MobilePhone = mobilePhone,
                                HomePhone = homePhone,
                                BusinessPhone = businessPhone,
                                CompanyName = contact.CompanyName ?? "",
                                JobTitle = contact.JobTitle ?? "",
                                Department = contact.Department ?? ""
                            });
                            exported++;
                        }
                    }
                    catch (Exception ex)
                    {
                        errors++;
                        _batchExportLogger.Error(ex, "导出联系人失败");
                    }
                }

                // 写入JSON文件
                if (contactItems.Count > 0)
                {
                    string jsonPath = Path.Combine(folderOutputDir, "contacts.json");
                    var jsonContent = System.Text.Json.JsonSerializer.Serialize(contactItems);
                    File.WriteAllText(jsonPath, jsonContent, new System.Text.UTF8Encoding(false));
                    _batchExportLogger.Information("联系人JSON已保存: {Count} 条记录", exported);
                }

                _batchExportLogger.Information("联系人导出完成: 成功={Exported}, 失败={Errors}", exported, errors);
            }
            catch (Exception ex)
            {
                _batchExportLogger.Error(ex, "导出联系人文件夹失败");
            }

            return (exported, errors);
        }

        /// <summary>
        /// 使用Outlook Interop导入日历和联系人到PST
        /// </summary>
        private void ImportCalendarAndContactsToPst(string pstPath, string tempDir)
        {
            try
            {
                _batchExportLogger.Information("开始导入日历和联系人到PST: {Path}", pstPath);

                // 启动Outlook
                var outlookType = Type.GetTypeFromProgID("Outlook.Application");
                dynamic outlook = Activator.CreateInstance(outlookType);
                var ns = outlook.GetNamespace("MAPI");

                // 挂载PST
                dynamic pstFolder = null;
                foreach (dynamic store in ns.Stores)
                {
                    try
                    {
                        if (store.FilePath != null && store.FilePath.Equals(pstPath, StringComparison.OrdinalIgnoreCase))
                        {
                            pstFolder = store.GetRootFolder();
                            break;
                        }
                    }
                    catch { }
                }

                if (pstFolder == null)
                {
                    _batchExportLogger.Error("无法找到PST文件夹: {Path}", pstPath);
                    Marshal.ReleaseComObject(ns);
                    Marshal.ReleaseComObject(outlook);
                    return;
                }

                // 获取或创建日历文件夹
                dynamic calendarFolder = null;
                try
                {
                    calendarFolder = pstFolder.Folders.Item("日历");
                }
                catch
                {
                    calendarFolder = pstFolder.Folders.Add("日历", 9); // olFolderCalendar = 9
                }

                // 获取或创建联系人文件夹
                dynamic contactsFolder = null;
                try
                {
                    contactsFolder = pstFolder.Folders.Item("联系人");
                }
                catch
                {
                    contactsFolder = pstFolder.Folders.Add("联系人", 10); // olFolderContacts = 10
                }

                // 导入日历
                string calendarJsonPath = Path.Combine(tempDir, "日历", "calendar.json");
                if (File.Exists(calendarJsonPath))
                {
                    string jsonContent = File.ReadAllText(calendarJsonPath, System.Text.Encoding.UTF8);

                    // 解析JSON数组
                    int calendarCount = 0;
                    using (var doc = System.Text.Json.JsonDocument.Parse(jsonContent))
                    {
                        foreach (var element in doc.RootElement.EnumerateArray())
                        {
                            dynamic appt = null;
                            try
                            {
                                appt = calendarFolder.Items.Add(1); // olAppointmentItem = 1
                                // 强制设置MessageClass确保被识别为日历
                                try { appt.MessageClass = "IPM.Appointment"; } catch { }

                                appt.Subject = element.GetProperty("Subject").GetString() ?? "";

                                string startStr = element.GetProperty("Start").GetString();
                                if (DateTime.TryParse(startStr, out DateTime start))
                                    appt.Start = start;

                                string endStr = element.GetProperty("End").GetString();
                                if (DateTime.TryParse(endStr, out DateTime end))
                                    appt.End = end;

                                appt.Location = element.GetProperty("Location").GetString() ?? "";
                                appt.Body = element.GetProperty("Body").GetString() ?? "";
                                appt.AllDayEvent = element.GetProperty("IsAllDayEvent").GetBoolean();

                                appt.Save();
                                calendarCount++;
                            }
                            catch (Exception ex)
                            {
                                _batchExportLogger.Error(ex, "导入日历项失败");
                            }
                            finally
                            {
                                try { if (appt != null) Marshal.ReleaseComObject(appt); } catch { }
                            }
                        }
                    }
                    _batchExportLogger.Information("日历导入完成: {Count} 条", calendarCount);
                }

                // 导入联系人
                string contactsJsonPath = Path.Combine(tempDir, "联系人", "contacts.json");
                if (File.Exists(contactsJsonPath))
                {
                    string jsonContent = File.ReadAllText(contactsJsonPath, System.Text.Encoding.UTF8);

                    int contactCount = 0;
                    using (var doc = System.Text.Json.JsonDocument.Parse(jsonContent))
                    {
                        foreach (var element in doc.RootElement.EnumerateArray())
                        {
                            dynamic contact = null;
                            try
                            {
                                contact = contactsFolder.Items.Add(2); // olContactItem = 2

                                contact.FullName = element.GetProperty("DisplayName").GetString() ?? "";
                                contact.FirstName = element.GetProperty("FirstName").GetString() ?? "";
                                contact.LastName = element.GetProperty("LastName").GetString() ?? "";
                                contact.Email1Address = element.GetProperty("Email1").GetString() ?? "";
                                contact.Email2Address = element.GetProperty("Email2").GetString() ?? "";
                                contact.Email3Address = element.GetProperty("Email3").GetString() ?? "";
                                contact.MobileTelephoneNumber = element.GetProperty("MobilePhone").GetString() ?? "";
                                contact.HomeTelephoneNumber = element.GetProperty("HomePhone").GetString() ?? "";
                                contact.BusinessTelephoneNumber = element.GetProperty("BusinessPhone").GetString() ?? "";
                                contact.CompanyName = element.GetProperty("CompanyName").GetString() ?? "";
                                contact.JobTitle = element.GetProperty("JobTitle").GetString() ?? "";
                                contact.Department = element.GetProperty("Department").GetString() ?? "";

                                contact.Save();
                                contactCount++;
                            }
                            catch (Exception ex)
                            {
                                _batchExportLogger.Error(ex, "导入联系人失败");
                            }
                            finally
                            {
                                try { if (contact != null) Marshal.ReleaseComObject(contact); } catch { }
                            }
                        }
                    }
                    _batchExportLogger.Information("联系人导入完成: {Count} 条", contactCount);
                }

                // 释放COM对象
                try { if (calendarFolder != null) Marshal.ReleaseComObject(calendarFolder); } catch { }
                try { if (contactsFolder != null) Marshal.ReleaseComObject(contactsFolder); } catch { }
                try { if (pstFolder != null) Marshal.ReleaseComObject(pstFolder); } catch { }
                try { if (ns != null) Marshal.ReleaseComObject(ns); } catch { }
                try { if (outlook != null) Marshal.ReleaseComObject(outlook); } catch { }

                // 强制垃圾回收，确保Outlook进程释放
                GC.Collect();
                GC.WaitForPendingFinalizers();

                _batchExportLogger.Information("日历和联系人导入PST完成");
            }
            catch (Exception ex)
            {
                _batchExportLogger.Error(ex, "导入日历和联系人到PST失败");
                throw;
            }
        }

        /// <summary>
        /// 生成VCF联系人内容
        /// </summary>
        private string GenerateVcfContent(Contact contact)
        {
            try
            {
                var sb = new System.Text.StringBuilder();
                sb.AppendLine("BEGIN:VCARD");
                sb.AppendLine("VERSION:2.1");
                sb.AppendLine($"FN:{contact.DisplayName ?? ""}");

                string email = null;
                if (contact.EmailAddresses != null && contact.EmailAddresses.Contains(EmailAddressKey.EmailAddress1))
                {
                    email = contact.EmailAddresses[EmailAddressKey.EmailAddress1]?.Address;
                }
                if (!string.IsNullOrEmpty(email))
                {
                    sb.AppendLine($"EMAIL;TYPE=PREF:{email}");
                }

                if (contact.PhoneNumbers != null && contact.PhoneNumbers.Contains(PhoneNumberKey.BusinessPhone))
                {
                    sb.AppendLine($"TEL;TYPE=WORK:{contact.PhoneNumbers[PhoneNumberKey.BusinessPhone]}");
                }
                if (contact.PhoneNumbers != null && contact.PhoneNumbers.Contains(PhoneNumberKey.MobilePhone))
                {
                    sb.AppendLine($"TEL;TYPE=CELL:{contact.PhoneNumbers[PhoneNumberKey.MobilePhone]}");
                }
                if (contact.PhoneNumbers != null && contact.PhoneNumbers.Contains(PhoneNumberKey.HomePhone))
                {
                    sb.AppendLine($"TEL;TYPE=HOME:{contact.PhoneNumbers[PhoneNumberKey.HomePhone]}");
                }

                sb.AppendLine("END:VCARD");
                return sb.ToString();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 导出任务到TXT文件
        /// </summary>
        private (int exported, int errors) ExportTasksToEml(
            string outputDir,
            DateTime startDate,
            DateTime endDate)
        {
            int exported = 0;
            int errors = 0;

            try
            {
                var tasksFolder = Folder.Bind(_service, new FolderId(WellKnownFolderName.Tasks)).GetAwaiter().GetResult();

                var searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                    new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeCreated, startDate),
                    new SearchFilter.IsLessThanOrEqualTo(ItemSchema.DateTimeCreated, endDate));

                var view = new ItemView(1000);
                view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.DateTimeCreated);

                var results = _service.FindItems(tasksFolder.Id, searchFilter, view).GetAwaiter().GetResult();

                string folderOutputDir = Path.Combine(outputDir, "任务");
                Directory.CreateDirectory(folderOutputDir);

                foreach (var item in results)
                {
                    try
                    {
                        string safeSubject = SanitizeFileName(item.Subject ?? "无主题");
                        string fileName = $"{safeSubject}.txt";
                        string filePath = Path.Combine(folderOutputDir, fileName);

                        int counter = 1;
                        while (File.Exists(filePath))
                        {
                            fileName = $"{safeSubject}_{counter}.txt";
                            filePath = Path.Combine(folderOutputDir, fileName);
                            counter++;
                        }

                        string content = $"主题: {item.Subject}\n创建时间: {item.DateTimeCreated}\n注意: 任务内容需要手动导出";
                        File.WriteAllText(filePath, content, new System.Text.UTF8Encoding(false));
                        exported++;
                    }
                    catch (Exception ex)
                    {
                        errors++;
                        _batchExportLogger.Error(ex, "导出任务失败: {Subject}", item.Subject ?? "(无主题)");
                    }
                }
            }
            catch (Exception ex)
            {
                _batchExportLogger.Error(ex, "导出任务文件夹失败");
            }

            return (exported, errors);
        }

        /// <summary>
        /// 是否已连接
        /// </summary>
        public bool IsConnected => _service != null;

        /// <summary>
        /// 导出邮件到EML文件
        /// </summary>
        public (int exported, int errors) ExportEmailsToEml(string outputDir, DateTime startDate, DateTime endDate)
        {
            return ExportAllFoldersToEml(outputDir, startDate, endDate);
        }

        /// <summary>
        /// 获取Outlook本地账户列表
        /// </summary>
        public List<string> GetOutlookAccounts()
        {
            var accounts = new List<string>();
            try
            {
                var outlook = new Microsoft.Office.Interop.Outlook.Application();
                var namespaces = outlook.GetNamespace("MAPI");
                var folders = namespaces.Folders;
                foreach (Microsoft.Office.Interop.Outlook.MAPIFolder folder in folders)
                {
                    try
                    {
                        accounts.Add(folder.Name);
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取Outlook账户失败");
            }
            return accounts;
        }

        /// <summary>
        /// 导出到PST（使用create_pst.py）
        /// </summary>
        public bool ExportToPst(string pstPath, string emlDir)
        {
            try
            {
                var exeDir = AppDomain.CurrentDomain.BaseDirectory;
                var pythonExe = Path.Combine(exeDir, "python", "python.exe");
                var scriptPath = Path.Combine(exeDir, "script", "create_pst.py");

                if (!File.Exists(pythonExe))
                {
                    Log.Error("Python环境未找到");
                    return false;
                }
                if (!File.Exists(scriptPath))
                {
                    Log.Error("create_pst.py脚本不存在");
                    return false;
                }

                var args = $"\"{scriptPath}\" \"{pstPath}\" \"{emlDir}\"";
                var startInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = pythonExe,
                    Arguments = args,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = System.Text.Encoding.UTF8,
                    StandardErrorEncoding = System.Text.Encoding.UTF8
                };
                startInfo.EnvironmentVariables["PYTHONNOUSERSITE"] = "1";
                startInfo.EnvironmentVariables["PYTHONIOENCODING"] = "UTF-8";

                using (var process = System.Diagnostics.Process.Start(startInfo))
                {
                    if (process != null)
                    {
                        process.WaitForExit(300000);
                        return process.ExitCode == 0;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "导出PST失败");
                return false;
            }
        }

        /// <summary>
        /// 测试Impersonation邮箱访问
        /// </summary>
        public (int inboxCount, string inboxError) TestImpersonationInbox()
        {
            try
            {
                if (_service == null)
                    return (0, "服务未初始化");

                var inbox = Folder.Bind(_service, WellKnownFolderName.Inbox).GetAwaiter().GetResult();
                return (inbox.TotalCount, null);
            }
            catch (Exception ex)
            {
                return (0, ex.Message);
            }
        }

        /// <summary>
        /// 文件夹信息类
        /// </summary>
        public class FolderInfo
        {
            public string Id { get; set; }
            public string DisplayName { get; set; }
            public string FolderClass { get; set; }
            public int TotalCount { get; set; }
        }

        /// <summary>
        /// 获取所有文件夹列表
        /// </summary>
        public List<FolderInfo> GetAllFolders()
        {
            var folders = new List<FolderInfo>();
            try
            {
                var view = new FolderView(500);
                view.Traversal = FolderTraversal.Deep;
                view.PropertySet = new PropertySet(
                    BasePropertySet.IdOnly,
                    FolderSchema.DisplayName,
                    FolderSchema.FolderClass,
                    FolderSchema.TotalCount);

                var results = _service.FindFolders(WellKnownFolderName.MsgFolderRoot, view).GetAwaiter().GetResult();

                foreach (Folder folder in results.Folders)
                {
                    try
                    {
                        if (folder is SearchFolder) continue;
                        string folderClass = folder.FolderClass ?? "";
                        if (folderClass != "IPF.Note" && folderClass != "" && !folderClass.StartsWith("IPF."))
                            continue;
                        folders.Add(new FolderInfo
                        {
                            Id = folder.Id.UniqueId,
                            DisplayName = folder.DisplayName,
                            FolderClass = folderClass,
                            TotalCount = folder.TotalCount
                        });
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取文件夹列表失败");
            }
            return folders;
        }

        /// <summary>
        /// 获取文件夹中的邮件MimeContent列表
        /// </summary>
        public List<Tuple<string, byte[], DateTime>> GetFolderEmailsMime(string folderId, DateTime startDate, DateTime endDate)
        {
            var mimeContents = new List<Tuple<string, byte[], DateTime>>();
            try
            {
                var folder = Folder.Bind(_service, new FolderId(folderId)).GetAwaiter().GetResult();

                var searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                    new SearchFilter.IsGreaterThanOrEqualTo(EmailMessageSchema.DateTimeReceived, startDate),
                    new SearchFilter.IsLessThanOrEqualTo(EmailMessageSchema.DateTimeReceived, endDate));

                var view = new ItemView(1000);
                view.PropertySet = new PropertySet(
                    BasePropertySet.IdOnly,
                    EmailMessageSchema.Subject,
                    EmailMessageSchema.DateTimeReceived);

                var items = _service.FindItems(folder.Id, searchFilter, view).GetAwaiter().GetResult();
                _batchToImapLogger.Information("FindItems返回 {Count} 个项目", items.Items.Count);

                // 批量加载MimeContent和Subject
                var emails = items.Where(i => i is EmailMessage).Cast<EmailMessage>().ToList();
                _batchToImapLogger.Information("其中 {Count} 个是EmailMessage", emails.Count);
                if (emails.Count > 0)
                {
                    var mimePropertySet = new PropertySet(
                        BasePropertySet.IdOnly,
                        EmailMessageSchema.MimeContent,
                        EmailMessageSchema.Subject,
                        EmailMessageSchema.DateTimeReceived);
                    _service.LoadPropertiesForItems(emails, mimePropertySet).GetAwaiter().GetResult();
                }

                foreach (var item in items)
                {
                    try
                    {
                        if (item is EmailMessage msg)
                        {
                            var subject = msg.Subject ?? "";
                            DateTime receivedDate = msg.DateTimeReceived;
                            byte[] mimeBytes = null;

                            if (msg.MimeContent != null && msg.MimeContent.Content != null)
                            {
                                mimeBytes = msg.MimeContent.Content;
                                _batchToImapLogger.Information("邮件 {Subject}: MimeContent大小={Size}", subject, mimeBytes.Length);
                            }
                            else
                            {
                                _batchToImapLogger.Warning("邮件 {Subject}: MimeContent为空或Content为null", subject);
                            }

                            if (mimeBytes != null && mimeBytes.Length > 0)
                            {
                                mimeContents.Add(Tuple.Create(subject, mimeBytes, receivedDate));
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _batchToImapLogger.Error(ex, "处理邮件时出错: {Item}", item.Subject ?? "(无主题)");
                    }
                }
            }
            catch (Exception ex)
            {
                _batchToImapLogger.Error(ex, "获取文件夹邮件失败");
            }
            return mimeContents;
        }

        /// <summary>
        /// 获取联系人列表
        /// </summary>
        public List<ContactInfo> GetContacts(string targetMailbox = null)
        {
            var contacts = new List<ContactInfo>();
            try
            {
                if (!string.IsNullOrEmpty(targetMailbox))
                {
                    SetImpersonatedUser(targetMailbox);
                }

                var contactsFolder = ContactsFolder.Bind(_service, WellKnownFolderName.Contacts).GetAwaiter().GetResult();
                var view = new ItemView(1000);
                view.PropertySet = new PropertySet(
                    BasePropertySet.IdOnly,
                    ContactSchema.DisplayName,
                    ContactSchema.EmailAddress1,
                    ContactSchema.EmailAddress2,
                    ContactSchema.EmailAddress3,
                    ContactSchema.CompanyName,
                    ContactSchema.PhoneNumbers);

                var results = _service.FindItems(contactsFolder.Id, view).GetAwaiter().GetResult();
                foreach (var item in results)
                {
                    try
                    {
                        if (item is Contact contact)
                        {
                            var info = new ContactInfo
                            {
                                DisplayName = contact.DisplayName,
                                Email1 = contact.EmailAddresses?.Contains(EmailAddressKey.EmailAddress1) == true
                                    ? contact.EmailAddresses[EmailAddressKey.EmailAddress1]?.Address : null,
                                Email2 = contact.EmailAddresses?.Contains(EmailAddressKey.EmailAddress2) == true
                                    ? contact.EmailAddresses[EmailAddressKey.EmailAddress2]?.Address : null,
                                Email3 = contact.EmailAddresses?.Contains(EmailAddressKey.EmailAddress3) == true
                                    ? contact.EmailAddresses[EmailAddressKey.EmailAddress3]?.Address : null,
                                Company = contact.CompanyName,
                                Phone = contact.PhoneNumbers?.Contains(PhoneNumberKey.BusinessPhone) == true
                                    ? contact.PhoneNumbers[PhoneNumberKey.BusinessPhone] : null
                            };
                            contacts.Add(info);
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取联系人失败");
            }
            return contacts;
        }

        /// <summary>
        /// 导出联系人到CSV
        /// </summary>
        public bool ExportContactsToCsv(string csvPath)
        {
            try
            {
                var contacts = GetContacts();
                using (var writer = new System.IO.StreamWriter(csvPath, false, System.Text.Encoding.UTF8))
                {
                    writer.WriteLine("DisplayName,Email1,Email2,Email3,Company,Phone");
                    foreach (var c in contacts)
                    {
                        writer.WriteLine($"\"{c.DisplayName ?? ""}\",\"{c.Email1 ?? ""}\",\"{c.Email2 ?? ""}\",\"{c.Email3 ?? ""}\",\"{c.Company ?? ""}\",\"{c.Phone ?? ""}\"");
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "导出联系人CSV失败");
                return false;
            }
        }

        /// <summary>
        /// 联系人信息类
        /// </summary>
        public class ContactInfo
        {
            public string DisplayName { get; set; }
            public string Email1 { get; set; }
            public string Email2 { get; set; }
            public string Email3 { get; set; }
            public string Company { get; set; }
            public string Phone { get; set; }
        }

        /// <summary>
        /// Exchange On-Premise用户信息类
        /// </summary>
        public class OnPremiseUserInfo
        {
            public string Id { get; set; }
            public string DisplayName { get; set; }
            public string Title { get; set; }
            public string Email { get; set; }
            public string Department { get; set; }
            public bool IsEnabled { get; set; }
            public string RecipientType { get; set; }
        }

        /// <summary>
        /// 邮件搜索结果类
        /// </summary>
        public class EmailSearchResult
        {
            public string Id { get; set; }
            public string Subject { get; set; }
            public string From { get; set; }
            public DateTime DateTimeReceived { get; set; }
            public bool HasAttachments { get; set; }
            public int Size { get; set; }
            public string Preview { get; set; }
        }

        /// <summary>
        /// 导出邮件到EML文件
        /// </summary>
        public (int exported, int errors) ExportEmailsToEml(
            string targetMailbox,
            Microsoft.Exchange.WebServices.Data.WellKnownFolderName folderName,
            DateTime startDate,
            DateTime endDate,
            string outputDir,
            Action<int, int> progressCallback,
            Action<string> errorCallback)
        {
            int exported = 0;
            int errors = 0;
            try
            {
                if (!string.IsNullOrEmpty(targetMailbox))
                {
                    SetImpersonatedUser(targetMailbox);
                }

                var folder = Folder.Bind(_service, folderName).GetAwaiter().GetResult();

                var searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                    new SearchFilter.IsGreaterThanOrEqualTo(EmailMessageSchema.DateTimeReceived, startDate),
                    new SearchFilter.IsLessThanOrEqualTo(EmailMessageSchema.DateTimeReceived, endDate));

                var view = new ItemView(1000);
                view.PropertySet = new PropertySet(
                    BasePropertySet.IdOnly,
                    EmailMessageSchema.Subject,
                    EmailMessageSchema.DateTimeReceived,
                    EmailMessageSchema.HasAttachments,
                    EmailMessageSchema.Size);

                var items = _service.FindItems(folder.Id, searchFilter, view).GetAwaiter().GetResult();
                int total = items.TotalCount;
                int processed = 0;

                foreach (var item in items)
                {
                    try
                    {
                        var mimeContent = GetMimeContent(item.Id.UniqueId);
                        if (!string.IsNullOrEmpty(mimeContent))
                        {
                            string safeSubject = SanitizeFileName(item.Subject ?? "无主题");
                            string fileName = $"{safeSubject}_{processed}.eml";
                            string filePath = Path.Combine(outputDir, fileName);
                            File.WriteAllText(filePath, mimeContent, new System.Text.UTF8Encoding(false));
                            exported++;
                        }
                        processed++;
                        progressCallback?.Invoke(processed, total);
                    }
                    catch (Exception ex)
                    {
                        errors++;
                        errorCallback?.Invoke(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                errorCallback?.Invoke(ex.Message);
            }
            return (exported, errors);
        }

        /// <summary>
        /// 导出到PST文件
        /// </summary>
        public bool ExportToPst(
            string targetEmail,
            string outputFile,
            bool includeReceived,
            bool includeSent,
            bool includeContacts,
            bool includeCalendar,
            Action<int, int, string> progressCallback)
        {
            try
            {
                if (!string.IsNullOrEmpty(targetEmail))
                {
                    SetImpersonatedUser(targetEmail);
                }

                // 创建临时目录
                string tempDir = Path.Combine(Path.GetTempPath(), $"pst_export_{Guid.NewGuid():N}");
                Directory.CreateDirectory(tempDir);

                try
                {
                    int processed = 0;
                    int total = 0;

                    // 导出邮件
                    if (includeReceived || includeSent)
                    {
                        progressCallback?.Invoke(processed, total, "正在导出邮件...");

                        var folders = GetAllFolders();
                        var emailFolders = folders.Where(f => f.TotalCount > 0).ToList();
                        total += emailFolders.Sum(f => f.TotalCount);

                        foreach (var folder in emailFolders)
                        {
                            try
                            {
                                if ((folder.DisplayName == "收件箱" || folder.DisplayName == "Inbox") && includeReceived)
                                {
                                    var emails = GetFolderEmailsMime(folder.Id, DateTime.MinValue, DateTime.MaxValue);
                                    foreach (var email in emails)
                                    {
                                        string filePath = Path.Combine(tempDir, $"{SanitizeFileName(email.Item1 ?? "无主题")}_{processed}.eml");
                                        File.WriteAllBytes(filePath, email.Item2);
                                        processed++;
                                        progressCallback?.Invoke(processed, total, $"正在导出邮件... {processed}/{total}");
                                    }
                                }
                                else if ((folder.DisplayName == "已发送" || folder.DisplayName == "Sent") && includeSent)
                                {
                                    var emails = GetFolderEmailsMime(folder.Id, DateTime.MinValue, DateTime.MaxValue);
                                    foreach (var email in emails)
                                    {
                                        string filePath = Path.Combine(tempDir, $"{SanitizeFileName(email.Item1 ?? "无主题")}_{processed}.eml");
                                        File.WriteAllBytes(filePath, email.Item2);
                                        processed++;
                                        progressCallback?.Invoke(processed, total, $"正在导出邮件... {processed}/{total}");
                                    }
                                }
                            }
                            catch { }
                        }
                    }

                    // 生成PST文件
                    progressCallback?.Invoke(processed, total, "正在创建PST...");

                    var exeDir = AppDomain.CurrentDomain.BaseDirectory;
                    var pythonExe = Path.Combine(exeDir, "python", "python.exe");
                    var scriptPath = Path.Combine(exeDir, "script", "create_pst.py");

                    if (!File.Exists(pythonExe))
                    {
                        progressCallback?.Invoke(processed, total, "错误: Python环境未找到");
                        return false;
                    }

                    var args = $"\"{scriptPath}\" \"{outputFile}\" \"{tempDir}\"";
                    var startInfo = new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = pythonExe,
                        Arguments = args,
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        CreateNoWindow = true,
                        StandardOutputEncoding = System.Text.Encoding.UTF8,
                        StandardErrorEncoding = System.Text.Encoding.UTF8
                    };
                    startInfo.EnvironmentVariables["PYTHONNOUSERSITE"] = "1";
                    startInfo.EnvironmentVariables["PYTHONIOENCODING"] = "UTF-8";

                    using (var process = System.Diagnostics.Process.Start(startInfo))
                    {
                        if (process != null)
                        {
                            process.WaitForExit(300000);
                            if (process.ExitCode == 0)
                            {
                                progressCallback?.Invoke(processed, total, "完成");
                                return true;
                            }
                        }
                    }
                    progressCallback?.Invoke(processed, total, "错误: PST创建失败");
                    return false;
                }
                finally
                {
                    try { Directory.Delete(tempDir, true); } catch { }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "导出PST失败");
                return false;
            }
        }

        /// <summary>
        /// 导出联系人到VCF
        /// </summary>
        public bool ExportContactsToVcf(string outputDir, int offset = 0, int limit = 100)
        {
            try
            {
                var contactsFolder = ContactsFolder.Bind(_service, WellKnownFolderName.Contacts).GetAwaiter().GetResult();
                var view = new ItemView(limit, offset, OffsetBasePoint.Beginning);
                view.PropertySet = new PropertySet(
                    BasePropertySet.IdOnly,
                    ContactSchema.DisplayName,
                    ContactSchema.EmailAddress1,
                    ContactSchema.EmailAddress2,
                    ContactSchema.EmailAddress3,
                    ContactSchema.CompanyName,
                    ContactSchema.PhoneNumbers,
                    ContactSchema.HomePhone,
                    ContactSchema.MobilePhone);

                var results = _service.FindItems(contactsFolder.Id, view).GetAwaiter().GetResult();
                int exported = 0;

                foreach (var item in results)
                {
                    try
                    {
                        if (item is Contact contact)
                        {
                            var vcfContent = GenerateVcfContent(contact);
                            if (!string.IsNullOrEmpty(vcfContent))
                            {
                                string safeName = SanitizeFileName(contact.DisplayName ?? "联系人");
                                string fileName = $"{safeName}_{exported}.vcf";
                                string filePath = Path.Combine(outputDir, fileName);
                                File.WriteAllText(filePath, vcfContent, new System.Text.UTF8Encoding(true));
                                exported++;
                            }
                        }
                    }
                    catch { }
                }

                return exported > 0;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "导出联系人VCF失败");
                return false;
            }
        }

        /// <summary>
        /// 日历事件信息类
        /// </summary>
        public class CalendarInfo
        {
            public string Id { get; set; }
            public string Subject { get; set; }
            public DateTime Start { get; set; }
            public DateTime End { get; set; }
            public string Location { get; set; }
            public bool IsAllDayEvent { get; set; }
        }

        /// <summary>
        /// 获取日历事件列表
        /// </summary>
        public List<CalendarInfo> GetCalendars(string targetMailbox, DateTime startDate, DateTime endDate)
        {
            var calendars = new List<CalendarInfo>();
            try
            {
                if (!string.IsNullOrEmpty(targetMailbox))
                {
                    SetImpersonatedUser(targetMailbox);
                }

                var calendarFolder = CalendarFolder.Bind(_service, WellKnownFolderName.Calendar).GetAwaiter().GetResult();

                var adjustedEndDate = endDate.AddDays(1).AddSeconds(-1);
                if ((adjustedEndDate - startDate).TotalDays > 730)
                    adjustedEndDate = startDate.AddDays(729);

                var calendarView = new CalendarView(startDate, adjustedEndDate, 1000);
                var appointments = calendarFolder.FindAppointments(calendarView).GetAwaiter().GetResult();

                foreach (var apt in appointments)
                {
                    try
                    {
                        calendars.Add(new CalendarInfo
                        {
                            Id = apt.Id?.UniqueId ?? Guid.NewGuid().ToString(),
                            Subject = apt.Subject,
                            Start = apt.Start,
                            End = apt.End,
                            Location = apt.Location ?? "",
                            IsAllDayEvent = apt.IsAllDayEvent
                        });
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取日历事件失败");
            }
            return calendars;
        }

        /// <summary>
        /// 导出日历到CSV
        /// </summary>
        public bool ExportCalendarsToCsv(string csvPath, List<CalendarInfo> calendars)
        {
            try
            {
                using (var writer = new System.IO.StreamWriter(csvPath, false, System.Text.Encoding.UTF8))
                {
                    writer.WriteLine("Subject,Start,End,Location,IsAllDayEvent,Duration");

                    foreach (var cal in calendars)
                    {
                        try
                        {
                            writer.WriteLine($"\"{cal.Subject ?? ""}\",\"{cal.Start:yyyy-MM-dd HH:mm:ss}\",\"{cal.End:yyyy-MM-dd HH:mm:ss}\",\"{cal.Location ?? ""}\",{cal.IsAllDayEvent},{(cal.End - cal.Start).TotalMinutes}");
                        }
                        catch { }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "导出日历CSV失败");
                return false;
            }
        }

        /// <summary>
        /// 导出日历到ICS
        /// </summary>
        public bool ExportCalendarsToIcs(string icsPath, List<CalendarInfo> calendars)
        {
            try
            {
                using (var writer = new System.IO.StreamWriter(icsPath, false, System.Text.Encoding.UTF8))
                {
                    writer.WriteLine("BEGIN:VCALENDAR");
                    writer.WriteLine("VERSION:2.0");
                    writer.WriteLine("PRODID:-//MailConverter//NONSGML v1.0//EN");

                    foreach (var cal in calendars)
                    {
                        try
                        {
                            writer.WriteLine("BEGIN:VEVENT");
                            writer.WriteLine($"UID:{cal.Id}");
                            writer.WriteLine($"DTSTART:{cal.Start.ToUniversalTime():yyyyMMddTHHmmssZ}");
                            writer.WriteLine($"DTEND:{cal.End.ToUniversalTime():yyyyMMddTHHmmssZ}");
                            writer.WriteLine($"SUMMARY:{EscapeIcsText(cal.Subject ?? "")}");
                            writer.WriteLine($"LOCATION:{EscapeIcsText(cal.Location ?? "")}");
                            writer.WriteLine($"DTSTAMP:{DateTime.UtcNow:yyyyMMddTHHmmssZ}");
                            writer.WriteLine("END:VEVENT");
                        }
                        catch { }
                    }

                    writer.WriteLine("END:VCALENDAR");
                }
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "导出日历ICS失败");
                return false;
            }
        }

        /// <summary>
        /// 获取用户列表（通过OWA/Exchange Web Services枚举）
        /// </summary>
        public List<OnPremiseUserInfo> GetUsers()
        {
            var users = new List<OnPremiseUserInfo>();
            try
            {
                // 使用全局地址列表
                var view = new ItemView(1000);
                view.PropertySet = new PropertySet(
                    BasePropertySet.IdOnly,
                    Microsoft.Exchange.WebServices.Data.FolderSchema.DisplayName,
                    Microsoft.Exchange.WebServices.Data.FolderSchema.FolderClass);

                // 由于Exchange On-Premise版本差异，这里返回一个空列表
                // 实际使用时需要通过PowerShell或其他方式获取用户列表
                Log.Warning("GetUsers方法在On-Premise模式下需要额外的AD/PowerShell集成");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取用户列表失败");
            }
            return users;
        }

        /// <summary>
        /// 设置用户启用/禁用状态
        /// </summary>
        public bool SetUserEnabled(string userEmail, bool enabled)
        {
            try
            {
                // Exchange On-Premise需要通过AD或PowerShell设置用户状态
                // 这里只是记录日志
                Log.Warning("SetUserEnabled需要通过AD/PowerShell集成: {Email}, Enabled: {Enabled}", userEmail, enabled);
                return false;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "设置用户状态失败");
                return false;
            }
        }

        /// <summary>
        /// 重置用户密码
        /// </summary>
        public bool ResetUserPassword(string userEmail, string newPassword)
        {
            try
            {
                // Exchange On-Premise需要通过AD或PowerShell重置密码
                Log.Warning("ResetUserPassword需要通过AD/PowerShell集成: {Email}", userEmail);
                return false;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "重置密码失败");
                return false;
            }
        }

        /// <summary>
        /// 搜索邮件（获取邮件数量）
        /// </summary>
        public int SearchEmails(DateTime startDate, DateTime? endDate, string from, string subject, bool isReceived)
        {
            try
            {
                var folderName = isReceived ? WellKnownFolderName.Inbox : WellKnownFolderName.SentItems;

                var searchFilters = new List<SearchFilter>();

                searchFilters.Add(new SearchFilter.IsGreaterThanOrEqualTo(EmailMessageSchema.DateTimeReceived, startDate));
                if (endDate.HasValue)
                    searchFilters.Add(new SearchFilter.IsLessThanOrEqualTo(EmailMessageSchema.DateTimeReceived, endDate.Value));
                if (!string.IsNullOrEmpty(from))
                    searchFilters.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.From, from));
                if (!string.IsNullOrEmpty(subject))
                    searchFilters.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.Subject, subject));

                var view = new ItemView(1);
                view.PropertySet = new PropertySet(BasePropertySet.IdOnly);

                SearchFilter finalFilter = null;
                if (searchFilters.Count > 0)
                    finalFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchFilters.ToArray());

                var items = _service.FindItems(folderName, finalFilter, view).GetAwaiter().GetResult();
                return items.TotalCount;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "搜索邮件失败");
                return 0;
            }
        }

        /// <summary>
        /// 获取总邮件数
        /// </summary>
        public int GetTotalMailCount(bool isReceived)
        {
            try
            {
                var folderName = isReceived ? WellKnownFolderName.Inbox : WellKnownFolderName.SentItems;
                var folder = Folder.Bind(_service, folderName).GetAwaiter().GetResult();
                return folder.TotalCount;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取总邮件数失败");
                return 0;
            }
        }

        /// <summary>
        /// 获取联系人总数
        /// </summary>
        public int GetContactCount()
        {
            try
            {
                var contactsFolder = ContactsFolder.Bind(_service, WellKnownFolderName.Contacts).GetAwaiter().GetResult();
                return contactsFolder.TotalCount;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取联系人总数失败");
                return 0;
            }
        }

        /// <summary>
        /// 搜索传输日志（模拟实现）
        /// </summary>
        /// <summary>
        /// 传输日志信息类
        /// </summary>
        public class TransportLogInfo
        {
            public string Status { get; set; }
            public string From { get; set; }
            public string To { get; set; }
            public string Subject { get; set; }
            public DateTime Timestamp { get; set; }
            public string Source { get; set; }
            public string EventId { get; set; }
            public string MessageId { get; set; }
        }

        /// <summary>
        /// 搜索传输日志
        /// </summary>
        public List<TransportLogInfo> SearchTransportLogs(
            string serverUrl,
            string email,
            string password,
            string domain,
            DateTime startDate,
            DateTime endDate,
            string from,
            string to,
            int maxResults)
        {
            var logs = new List<TransportLogInfo>();
            try
            {
                // Exchange On-Premise传输日志需要通过PowerShell或IIS日志访问
                // 这里返回一个模拟的结果
                Log.Warning("SearchTransportLogs需要额外的日志集成");
                logs.Add(new TransportLogInfo
                {
                    Status = "DELIVERED",
                    From = from,
                    To = to,
                    Subject = "传输日志搜索需要PowerShell集成",
                    Timestamp = DateTime.Now,
                    Source = "Exchange",
                    EventId = "1",
                    MessageId = Guid.NewGuid().ToString()
                });
            }
            catch (Exception ex)
            {
                Log.Error(ex, "搜索传输日志失败");
            }
            return logs;
        }

        /// <summary>
        /// 根据ID获取邮件
        /// </summary>
        public EmailSearchResult GetEmailById(string itemId, string targetMailbox = null)
        {
            try
            {
                if (!string.IsNullOrEmpty(targetMailbox))
                {
                    SetImpersonatedUser(targetMailbox);
                }

                var item = Item.Bind(_service, itemId).GetAwaiter().GetResult();
                if (item is EmailMessage msg)
                {
                    return new EmailSearchResult
                    {
                        Id = item.Id.UniqueId,
                        Subject = msg.Subject,
                        From = msg.From?.Address ?? "",
                        DateTimeReceived = msg.DateTimeReceived,
                        HasAttachments = msg.HasAttachments,
                        Size = msg.Size,
                        Preview = (msg.Subject ?? "").Length > 50 ? msg.Subject.Substring(0, 50) : msg.Subject
                    };
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取邮件失败");
            }
            return null;
        }
    }
}
