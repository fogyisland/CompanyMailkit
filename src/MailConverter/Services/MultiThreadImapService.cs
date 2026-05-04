using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using MailKit;
using MailKit.Search;

namespace MailConverter
{
    /// <summary>
    /// IMAP账户配置
    /// </summary>
    public class ImapAccount
    {
        public string Email { get; set; }
        public string Password { get; set; }
        public string PstFilePath { get; set; }
        public string DisplayName { get; set; }
        public string Host { get; set; }
        public int Port { get; set; } = 993;
        public bool UseSsl { get; set; } = true;
    }

    /// <summary>
    /// IMAP同步进度信息
    /// </summary>
    public class ImapSyncProgress
    {
        public int CompletedAccounts { get; set; }
        public int TotalAccounts { get; set; }
        public string CurrentEmail { get; set; }
        public string CurrentFolder { get; set; }
        public int FolderEmailCount { get; set; }
        public int TotalEmailsAll { get; set; }
        public int TotalEmailsCurrentFolder { get; set; }
        public int TotalFolders { get; set; }
        public int CompletedFolders { get; set; }
    }

    /// <summary>
    /// 多线程IMAP同步服务
    /// </summary>
    public class MultiThreadImapService
    {
        private Action<string> _logCallback;
        private Action<ImapSyncProgress> _progressCallback;
        private int _completedCount;
        private int _totalCount;
        private int _totalEmailsAll;
        private readonly object _lockObj = new object();

        public void SetCallbacks(Action<string> logCallback, Action<ImapSyncProgress> progressCallback)
        {
            _logCallback = logCallback;
            _progressCallback = progressCallback;
        }

        /// <summary>
        /// 批量导入IMAP账户并同步到各自的PST
        /// </summary>
        public async Task<bool> SyncAccountsAsync(
            List<ImapAccount> accounts,
            int maxParallel = 5,
            int maxEmailsPerAccount = 1000,
            int daysBack = 0)
        {
            _totalCount = accounts.Count;
            _completedCount = 0;

            DateTime? since = null;
            if (daysBack > 0)
            {
                since = DateTime.Now.AddDays(-daysBack);
                Log($"日期过滤: 下载最近 {daysBack} 天的邮件 (从 {since:yyyy-MM-dd} 至今)");
            }

            Log($"开始同步 {_totalCount} 个账户，最大并行数: {maxParallel}");

            // 使用SemaphoreSlim限制并行数
            using (var semaphore = new SemaphoreSlim(maxParallel))
            {
                var tasks = accounts.Select(async account =>
                {
                    await semaphore.WaitAsync();
                    try
                    {
                        await SyncSingleAccountAsync(account, maxEmailsPerAccount, since);
                    }
                    finally
                    {
                        semaphore.Release();

                        lock (_lockObj)
                        {
                            _completedCount++;
                            _progressCallback?.Invoke(new ImapSyncProgress
                            {
                                CompletedAccounts = _completedCount,
                                TotalAccounts = _totalCount,
                                CurrentEmail = account.Email
                            });
                        }
                    }
                });

                await Task.WhenAll(tasks);
            }

            Log($"全部完成！ {_totalCount} 个账户同步完成");
            return true;
        }

        private async Task SyncSingleAccountAsync(ImapAccount account, int maxEmails, DateTime? since = null)
        {
            string logPrefix = $"[{account.Email}]";
            Log($"{logPrefix} 开始同步...");

            try
            {
                // 使用指定的服务器配置
                var config = new ImapAccountConfig();
                config.Email = account.Email;
                config.Password = account.Password;

                if (!string.IsNullOrEmpty(account.Host))
                {
                    config.Host = account.Host;
                    config.Port = account.Port > 0 ? account.Port : 993;
                    config.UseSsl = account.UseSsl;
                }
                else
                {
                    // 自动发现IMAP配置
                    var autoDiscover = new AutoDiscoverService();
                    var autoConfig = await autoDiscover.AutoDiscoverAsync(account.Email, account.Password);
                    config.Host = autoConfig.Host;
                    config.Port = autoConfig.Port;
                    config.UseSsl = autoConfig.UseSsl;
                }

                Log($"{logPrefix} IMAP服务器: {config.Host}:{config.Port}");

                // 创建临时目录
                var tempDir = Path.Combine(Path.GetTempPath(), $"imap_{Guid.NewGuid():N}");
                Directory.CreateDirectory(tempDir);

                try
                {
                    // 下载邮件并投递到PST（内部会创建PST并导入）
                    await DownloadAndDeliverAsync(config, account.PstFilePath, maxEmails, logPrefix, since);

                    Log($"{logPrefix} 同步完成: {account.PstFilePath}");
                }
                finally
                {
                    // 清理临时目录
                    try { Directory.Delete(tempDir, true); } catch { }
                }
            }
            catch (Exception ex)
            {
                Log($"{logPrefix} 同步失败: {ex.Message}");
                Log($"{logPrefix} 堆栈: {ex.StackTrace}");
                if (ex.InnerException != null)
                    Log($"{logPrefix} 内部错误: {ex.InnerException.Message}");
            }
        }

        private async Task DownloadEmailsAsync(ImapAccountConfig config, string outputDir, int maxEmails, string logPrefix)
        {
            using (var client = new MailKit.Net.Imap.ImapClient())
            {
                await client.ConnectAsync(config.Host, config.Port, config.UseSsl);
                await client.AuthenticateAsync(config.Email, config.Password);

                // 获取所有文件夹
                var folders = await client.GetFoldersAsync(client.PersonalNamespaces[0]);
                Log($"{logPrefix} 发现 {folders.Count} 个文件夹");

                int totalDownloaded = 0;

                foreach (var folder in folders)
                {
                    if (totalDownloaded >= maxEmails)
                        break;

                    try
                    {
                        await folder.OpenAsync(MailKit.FolderAccess.ReadOnly);
                        int count = folder.Count;

                        if (count == 0)
                        {
                            await folder.CloseAsync();
                            continue;
                        }

                        Log($"{logPrefix} 处理文件夹: {folder.Name} ({count} 封)");

                        // 获取邮件UID
                        var uids = folder.Search(SearchQuery.All);
                        var messages = folder.Fetch(uids, MessageSummaryItems.Envelope | MessageSummaryItems.Flags);

                        int folderDownloaded = 0;
                        foreach (var msg in messages)
                        {
                            if (totalDownloaded >= maxEmails)
                                break;

                            try
                            {
                                var mimeMsg = await folder.GetMessageAsync(msg.UniqueId);

                                string subject = mimeMsg.Subject ?? "No Subject";
                                foreach (char c in Path.GetInvalidFileNameChars())
                                {
                                    subject = subject.Replace(c, '_');
                                }
                                if (subject.Length > 80)
                                    subject = subject.Substring(0, 80);

                                string emlPath = Path.Combine(outputDir, $"{subject}_{msg.UniqueId}.eml");

                                // 处理重复文件名
                                int counter = 1;
                                while (File.Exists(emlPath))
                                {
                                    emlPath = Path.Combine(outputDir, $"{subject}_{msg.UniqueId}_{counter}.eml");
                                    counter++;
                                }

                                await mimeMsg.WriteToAsync(emlPath);
                                totalDownloaded++;
                                folderDownloaded++;

                                if (folderDownloaded % 20 == 0)
                                {
                                    Log($"{logPrefix} 已下载: {totalDownloaded} 封");
                                }
                            }
                            catch (Exception ex)
                            {
                                Log($"{logPrefix} 下载邮件失败: {ex.Message}");
                            }
                        }

                        await folder.CloseAsync();
                        Log($"{logPrefix} 文件夹 {folder.Name} 完成: {folderDownloaded} 封");
                    }
                    catch (Exception ex)
                    {
                        Log($"{logPrefix} 处理文件夹失败: {folder.Name} - {ex.Message}");
                        try { await folder.CloseAsync(); } catch { }
                    }
                }

                await client.DisconnectAsync(true);
                Log($"{logPrefix} 共下载: {totalDownloaded} 封邮件");
            }
        }

        private async Task CreatePstFromEml(string emlDir, string pstPath, string pstName)
        {
            // 删除旧PST
            if (File.Exists(pstPath))
            {
                File.Delete(pstPath);
            }

            var exeDir = AppDomain.CurrentDomain.BaseDirectory;
            var scriptPath = Path.Combine(exeDir, "script", "create_pst.py");

            if (!File.Exists(scriptPath))
            {
                throw new Exception("Python脚本不存在: " + scriptPath);
            }

            // 检查EML文件数量
            var emlFiles = Directory.GetFiles(emlDir, "*.eml", SearchOption.AllDirectories);
            Log($"[CreatePstFromEml] EML目录: {emlDir}, 文件数: {emlFiles.Length}");

            var pythonExe = Program.GetPythonExecutable();
            if (string.IsNullOrEmpty(pythonExe))
            {
                throw new Exception("Python环境未找到，请重新安装程序");
            }

            var cmdLine = $"\"{scriptPath}\" \"{pstPath}\" \"{emlDir}\"";
            Log($"[CreatePstFromEml] 执行命令: {pythonExe} {cmdLine}");

            var startInfo = Program.CreatePythonStartInfo(pythonExe, cmdLine);

            using (var process = System.Diagnostics.Process.Start(startInfo))
            {
                if (process != null)
                {
                    process.WaitForExit(300000);
                    var output = process.StandardOutput.ReadToEnd();
                    var error = process.StandardError.ReadToEnd();
                    var exitCode = process.ExitCode;

                    Log($"[CreatePstFromEml] Python退出码: {exitCode}");
                    if (!string.IsNullOrEmpty(output))
                        Log("PST创建: " + output);
                    if (!string.IsNullOrEmpty(error))
                        Log("PST错误: " + error);
                }
            }

            if (!File.Exists(pstPath))
            {
                throw new Exception("PST文件创建失败");
            }
        }

        /// <summary>
        /// 创建空的PST文件
        /// </summary>
        private async Task CreateEmptyPst(string pstPath, string pstName)
        {
            // 删除旧PST
            if (File.Exists(pstPath))
            {
                File.Delete(pstPath);
            }

            var exeDir = AppDomain.CurrentDomain.BaseDirectory;
            var scriptPath = Path.Combine(exeDir, "script", "create_pst.py");

            if (!File.Exists(scriptPath))
            {
                throw new Exception("Python脚本不存在: " + scriptPath);
            }

            var pythonExe = Program.GetPythonExecutable();
            if (string.IsNullOrEmpty(pythonExe))
            {
                throw new Exception("Python环境未找到，请重新安装程序");
            }

            var startInfo = Program.CreatePythonStartInfo(pythonExe, $"\"{scriptPath}\" \"{pstPath}\" \"{pstName}\"");

            using (var process = System.Diagnostics.Process.Start(startInfo))
            {
                if (process != null)
                {
                    process.WaitForExit(60000);
                }
            }

            if (!File.Exists(pstPath))
            {
                throw new Exception("PST文件创建失败");
            }
        }

        /// <summary>
        /// 下载并实时投递邮件到PST
        /// </summary>
        private async Task DownloadAndDeliverAsync(ImapAccountConfig config, string pstPath, int maxEmails, string logPrefix, DateTime? since = null)
        {
            var exeDir = AppDomain.CurrentDomain.BaseDirectory;
            var deliverScript = Path.Combine(exeDir, "script", "deliver_mail.py");

            // 始终使用临时文件方式，因为 deliver_mail.py 需要 PST 已存在
            bool useTempFiles = true;

            var tempDir = Path.Combine(Path.GetTempPath(), $"deliver_{Guid.NewGuid():N}");
            Directory.CreateDirectory(tempDir);

            try
            {
                using (var client = new MailKit.Net.Imap.ImapClient())
                {
                    client.CheckCertificateRevocation = false;
                    await client.ConnectAsync(config.Host, config.Port, config.UseSsl);
                    await client.AuthenticateAsync(config.Email, config.Password);

                    var folders = await client.GetFoldersAsync(client.PersonalNamespaces[0]);
                    Log($"{logPrefix} 发现 {folders.Count} 个文件夹");

                    int totalDelivered = 0;

                    foreach (var folder in folders)
                    {
                        if (totalDelivered >= maxEmails)
                            break;

                        try
                        {
                            await folder.OpenAsync(MailKit.FolderAccess.ReadOnly);
                            int count = folder.Count;

                            if (count == 0)
                            {
                                await folder.CloseAsync();
                                continue;
                            }

                            string folderName = folder.Name;
                            Log($"{logPrefix} 处理文件夹: {folderName} ({count} 封)");

                            // 创建文件夹子目录（按文件夹分类）
                            var folderSubDir = Path.Combine(tempDir, SanitizeFolderName(folderName));
                            Directory.CreateDirectory(folderSubDir);

                            // 【优化】使用IMAP服务器端搜索，按日期范围过滤
                            IList<UniqueId> uidsToFetch;
                            SearchQuery searchQuery = SearchQuery.All;

                            if (since.HasValue)
                            {
                                searchQuery = SearchQuery.DeliveredAfter(since.Value);
                                Log($"{logPrefix} 使用日期过滤: 从 {since.Value:yyyy-MM-dd} 至今");
                            }

                            var searchResult = folder.Search(searchQuery);
                            uidsToFetch = searchResult;

                            if (uidsToFetch.Count == 0)
                            {
                                Log($"{logPrefix} 没有符合条件的邮件，跳过");
                                await folder.CloseAsync();
                                continue;
                            }

                            // 按UID从最新到最旧排序（优先下载最新的）
                            var sortedUids = uidsToFetch.OrderByDescending(u => u.Id).ToList();
                            Log($"{logPrefix} 文件夹 {folderName} 符合条件 {sortedUids.Count} 封，开始下载...");

                            int folderDownloaded = 0;
                            foreach (var uid in sortedUids)
                            {
                                if (totalDelivered >= maxEmails)
                                    break;

                                try
                                {
                                    var mimeMsg = await folder.GetMessageAsync(uid);

                                    if (useTempFiles)
                                    {
                                        // 保存到临时文件，累积后一起写入PST
                                        string subject = mimeMsg.Subject ?? "No Subject";
                                        string safeSubject = SanitizeFileName(subject);
                                        if (safeSubject.Length > 80)
                                            safeSubject = safeSubject.Substring(0, 80);

                                        string emlPath = Path.Combine(folderSubDir, $"{safeSubject}_{uid}.eml");
                                        int counter = 1;
                                        while (File.Exists(emlPath))
                                        {
                                            emlPath = Path.Combine(folderSubDir, $"{safeSubject}_{uid}_{counter}.eml");
                                            counter++;
                                        }

                                        await mimeMsg.WriteToAsync(emlPath);
                                    }
                                    else
                                    {
                                        // 方式2: 实时投递到PST
                                        await DeliverMailToPst(mimeMsg, pstPath, deliverScript, logPrefix);
                                    }

                                    totalDelivered++;
                                    folderDownloaded++;

                                    if (totalDelivered % 10 == 0)
                                    {
                                        Log($"{logPrefix} 已投递: {totalDelivered} 封");
                                        _progressCallback?.Invoke(new ImapSyncProgress
                                        {
                                            CompletedAccounts = _completedCount,
                                            TotalAccounts = _totalCount,
                                            CurrentEmail = config.Email,
                                            CurrentFolder = folder.Name,
                                            FolderEmailCount = totalDelivered,
                                            TotalEmailsAll = totalDelivered
                                        });
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Log($"{logPrefix} 投递失败: {ex.Message}");
                                }
                            }

                            await folder.CloseAsync();
                            Log($"{logPrefix} 文件夹 {folder.Name} 完成");
                        }
                        catch (Exception ex)
                        {
                            Log($"{logPrefix} 处理文件夹失败: {folder.Name} - {ex.Message}");
                            try { await folder.CloseAsync(); } catch { }
                        }
                    }

                    await client.DisconnectAsync(true);

                    Log($"{logPrefix} 断开连接，useTempFiles={useTempFiles}, totalDelivered={totalDelivered}");

                    // 如果使用临时文件方式，最后写入PST
                    if (useTempFiles)
                    {
                        Log($"{logPrefix} useTempFiles=true，准备写入PST");
                        if (totalDelivered > 0)
                        {
                            Log($"{logPrefix} 正在写入PST...");
                            await CreatePstFromEml(tempDir, pstPath, "ImapSync");
                        }
                        else
                        {
                            // 即使没有邮件，也创建空PST
                            Log($"{logPrefix} 没有邮件，创建空PST...");
                            await CreateEmptyPst(pstPath, config.Email);
                        }
                    }
                    else
                    {
                        Log($"{logPrefix} useTempFiles=false，跳过PST写入");
                    }

                    Log($"{logPrefix} 共投递: {totalDelivered} 封邮件");
                }
            }
            finally
            {
                try { Directory.Delete(tempDir, true); } catch { }
            }
        }

        /// <summary>
        /// 实时投递单封邮件到PST
        /// </summary>
        private async Task DeliverMailToPst(MimeKit.MimeMessage mimeMsg, string pstPath, string scriptPath, string logPrefix)
        {
            // 保存到临时EML文件
            var tempEml = Path.Combine(Path.GetTempPath(), $"temp_{Guid.NewGuid():N}.eml");
            try
            {
                await mimeMsg.WriteToAsync(tempEml);

                var pythonExe = Program.GetPythonExecutable();
                if (string.IsNullOrEmpty(pythonExe))
                {
                    throw new Exception("Python环境未找到，请重新安装程序");
                }

                // 调用Python脚本投递到PST
                var startInfo = Program.CreatePythonStartInfo(pythonExe, $"\"{scriptPath}\" \"{pstPath}\" \"{tempEml}\"");

                using (var process = System.Diagnostics.Process.Start(startInfo))
                {
                    if (process != null)
                    {
                        process.WaitForExit(30000);
                    }
                }
            }
            finally
            {
                try { File.Delete(tempEml); } catch { }
            }
        }

        /// <summary>
        /// 清理文件夹名称，移除无效字符
        /// </summary>
        private string SanitizeFolderName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "Inbox";
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name;
        }

        /// <summary>
        /// 清理文件名，移除无效字符
        /// </summary>
        private string SanitizeFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "NoSubject";
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            // 移除可能导致问题的其他字符
            name = name.Replace(':', '_').Replace('|', '_').Replace('?', '_').Replace('*', '_');
            return name.Trim();
        }

        private void Log(string message)
        {
            _logCallback?.Invoke(message);

            var logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs", "mailconverter.log");
            File.AppendAllText(logPath, $"[MultiThreadImap] {message}\n");
        }
    }
}
