using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using MailKit;
using MimeKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;

namespace MailConverter
{
    /// <summary>
    /// IMAP收件服务 - 从IMAP服务器下载邮件到PST
    /// </summary>
    public class ImapFetchService
    {
        private readonly AutoDiscoverService _autoDiscover;
        private Action<string> _logCallback;
        private Action<int, int> _progressCallback;
        private int _consecutiveFailures = 0;

        public ImapFetchService()
        {
            _autoDiscover = new AutoDiscoverService();
        }

        /// <summary>
        /// 根据邮件服务器获取安全的并发连接数
        /// </summary>
        private int GetSafePoolSize(string host)
        {
            var lowerHost = host.ToLower();
            if (lowerHost.Contains("gmail") || lowerHost.Contains("google"))
                return 8;  // Gmail限制15，用8安全
            if (lowerHost.Contains("outlook") || lowerHost.Contains("office") || lowerHost.Contains("hotmail"))
                return 10; // Outlook限制较高
            if (lowerHost.Contains("qq"))
                return 5;  // QQ邮箱保守设置
            if (lowerHost.Contains("163") || lowerHost.Contains("126") || lowerHost.Contains("yeah.net"))
                return 5;  // 网易邮箱保守设置
            return 5;      // 未知服务器默认5
        }

        /// <summary>
        /// 带重试的邮件下载
        /// </summary>
        private async Task<MimeMessage> DownloadWithRetryAsync(IMailFolder folder, UniqueId uid, int maxRetries = 3)
        {
            int retryCount = 0;
            while (retryCount < maxRetries)
            {
                try
                {
                    return await folder.GetMessageAsync(uid);
                }
                catch (Exception ex)
                {
                    retryCount++;
                    _consecutiveFailures++;

                    if (retryCount >= maxRetries)
                    {
                        throw new Exception($"下载邮件 {uid} 失败，已重试 {maxRetries} 次: {ex.Message}");
                    }

                    // 指数退避: 1秒, 2秒, 4秒
                    var delay = TimeSpan.FromSeconds(Math.Pow(2, retryCount - 1));
                    Log($"邮件 {uid} 下载失败，{delay.TotalSeconds}秒后重试... ({ex.Message})");
                    await Task.Delay(delay);
                }
            }
            throw new Exception($"下载邮件 {uid} 失败");
        }

        public void SetCallbacks(Action<string> logCallback, Action<int, int> progressCallback)
        {
            _logCallback = logCallback;
            _progressCallback = progressCallback;
        }

        /// <summary>
        /// 获取IMAP文件夹列表
        /// </summary>
        public async Task<List<string>> GetFoldersAsync(string email, string password)
        {
            var folders = new List<string>();
            try
            {
                Log($"正在获取文件夹列表: {email}");

                var config = await _autoDiscover.AutoDiscoverAsync(email, password);
                config.Email = email;
                config.Password = password;

                Log($"使用配置: {config.Host}:{config.Port}");

                using (var client = new ImapClient())
                {
                    // 禁用证书吊销检查，避免SSL握手卡顿
                    client.CheckCertificateRevocation = false;
                    await client.ConnectAsync(config.Host, config.Port, config.UseSsl);
                    await client.AuthenticateAsync(email, password);

                    // 使用更简单的方法：直接从 PersonalNamespaces[0] 获取
                    if (client.PersonalNamespaces != null && client.PersonalNamespaces.Count > 0)
                    {
                        var allFolders = await client.GetFoldersAsync(client.PersonalNamespaces[0]);

                        foreach (var folder in allFolders)
                        {
                            if (!string.IsNullOrEmpty(folder.Name))
                            {
                                folders.Add(folder.Name);
                            }
                        }
                    }
                    else
                    {
                        // 备用：尝试获取 Inbox 的子文件夹
                        var inbox = client.Inbox;
                        await inbox.OpenAsync(MailKit.FolderAccess.ReadOnly);
                        var subfolders = await inbox.GetSubfoldersAsync();

                        foreach (var folder in subfolders)
                        {
                            if (!string.IsNullOrEmpty(folder.Name))
                            {
                                folders.Add(folder.Name);
                            }
                        }
                    }

                    await client.DisconnectAsync(true);
                }

                Log($"获取到 {folders.Count} 个文件夹");
            }
            catch (Exception ex)
            {
                Log($"获取文件夹失败: {ex.Message}");
                System.Diagnostics.Debug.WriteLine(ex.StackTrace);
            }
            return folders;
        }

        /// <summary>
        /// 获取各年份的邮件数量统计
        /// </summary>
        public async Task<Dictionary<int, int>> GetYearlyStatsAsync(string email, string password, List<string> selectedFolders)
        {
            var stats = new Dictionary<int, int>();

            try
            {
                Log($"正在统计年度邮件数量...");

                var config = await _autoDiscover.AutoDiscoverAsync(email, password);
                config.Email = email;
                config.Password = password;

                using (var client = new ImapClient())
                {
                    client.CheckCertificateRevocation = false;
                    await client.ConnectAsync(config.Host, config.Port, config.UseSsl);
                    await client.AuthenticateAsync(email, password);

                    var folders = await client.GetFoldersAsync(client.PersonalNamespaces[0]);

                    if (selectedFolders != null && selectedFolders.Count > 0)
                    {
                        var selectedSet = new HashSet<string>(selectedFolders, StringComparer.OrdinalIgnoreCase);
                        folders = folders.Where(f => selectedSet.Contains(f.Name)).ToList();
                    }

                    foreach (var folder in folders)
                    {
                        try
                        {
                            await folder.OpenAsync(MailKit.FolderAccess.ReadOnly);
                            var uids = folder.Search(MailKit.Search.SearchQuery.All);

                            // 获取所有邮件的摘要（包含日期）
                            var summaries = await folder.FetchAsync(uids, MessageSummaryItems.Envelope);

                            foreach (var summary in summaries)
                            {
                                if (summary.Date != default)
                                {
                                    int year = summary.Date.LocalDateTime.Year;
                                    if (!stats.ContainsKey(year))
                                        stats[year] = 0;
                                    stats[year]++;
                                }
                            }

                            await folder.CloseAsync();
                            Log($"文件夹 {folder.Name} 统计完成");
                        }
                        catch (Exception ex)
                        {
                            Log($"统计文件夹 {folder.Name} 失败: {ex.Message}");
                        }
                    }

                    await client.DisconnectAsync(true);
                }

                Log($"年度统计完成，共 {stats.Values.Sum()} 封邮件");
            }
            catch (Exception ex)
            {
                Log($"年度统计失败: {ex.Message}");
            }

            return stats;
        }

        /// <summary>
        /// 从IMAP服务器收取邮件到PST文件
        /// </summary>
        /// <param name="email">邮箱地址</param>
        /// <param name="password">密码</param>
        /// <param name="outputPstPath">输出PST路径</param>
        /// <param name="maxEmails">最大邮件数</param>
        /// <param name="since">起始日期过滤</param>
        /// <param name="selectedFolders">要同步的文件夹列表（null表示所有文件夹）</param>
        /// <param name="cancellationToken">取消令牌</param>
        /// <param name="batchSize">每批获取的邮件数</param>
        public async Task<bool> FetchEmailsToPstAsync(
            string email,
            string password,
            string outputPstPath,
            int maxEmails = 1000,
            DateTime? since = null,
            List<string> selectedFolders = null,
            CancellationToken cancellationToken = default,
            int batchSize = 100,
            DateTime? endDate = null)
        {
            try
            {
                Log($"开始收取邮件: {email}");

                // 自动发现IMAP配置
                var config = await _autoDiscover.AutoDiscoverAsync(email, password);
                config.Email = email;
                config.Password = password;

                Log($"IMAP服务器: {config.Host}:{config.Port}");

                using (var client = new ImapClient())
                {
                    Log($"创建ImapClient成功");

                    // 禁用证书吊销检查，避免SSL握手卡顿
                    client.CheckCertificateRevocation = false;
                    Log($"开始ConnectAsync...");
                    await client.ConnectAsync(config.Host, config.Port, config.UseSsl);
                    Log($"ConnectAsync完成");

                    Log($"开始AuthenticateAsync...");
                    await client.AuthenticateAsync(email, password);
                    Log($"AuthenticateAsync完成");

                    // 获取所有文件夹
                    Log($"开始GetFoldersAsync...");
                    var folders = await client.GetFoldersAsync(client.PersonalNamespaces[0]);
                    Log($"发现 {folders.Count} 个文件夹");

                    // 创建PST文件
                    var exeDir = AppDomain.CurrentDomain.BaseDirectory;
                    var createPstScript = Path.Combine(exeDir, "script", "create_pst.py");

                    if (!File.Exists(createPstScript))
                    {
                        throw new Exception("Python脚本不存在: " + createPstScript);
                    }

                    // 创建临时EML目录
                    var tempDir = Path.Combine(Path.GetTempPath(), "imap_fetch_" + Guid.NewGuid().ToString("N"));
                    Directory.CreateDirectory(tempDir);

                    int totalEmails = 0;

                    // 如果指定了文件夹列表，只处理选中的文件夹
                    if (selectedFolders != null && selectedFolders.Count > 0)
                    {
                        var selectedSet = new HashSet<string>(selectedFolders, StringComparer.OrdinalIgnoreCase);
                        folders = folders.Where(f => selectedSet.Contains(f.Name)).ToList();
                        Log($"将同步以下文件夹: {string.Join(", ", selectedFolders)}");
                    }

                    try
                    {
                        // 处理每个文件夹
                        int folderIndex = 0;
                        foreach (var folder in folders)
                        {
                            folderIndex++;
                            Log($"=== 开始处理文件夹 {folderIndex}/{folders.Count}: {folder.Name} ===");
                            if (totalEmails >= maxEmails)
                                break;

                            try
                            {
                                totalEmails = await ProcessFolderAsync(client, folder, tempDir, totalEmails, maxEmails, since, cancellationToken, batchSize, endDate);
                            }
                            catch (Exception ex)
                            {
                                Log($"文件夹 {folder.Name} 处理异常: {ex.GetType().Name} - {ex.Message}");
                                Log($"堆栈: {ex.StackTrace}");
                                continue;
                            }
                            Log($"=== 文件夹 {folder.Name} 处理完成，当前总计: {totalEmails} ===");
                        }

                        Log($"共收取 {totalEmails} 封邮件");

                        // 将EML导入到PST
                        if (totalEmails > 0)
                        {
                            Log("正在将邮件导入PST...");
                            Log($"PST路径: {outputPstPath}");
                            Log($"EML目录: {tempDir}");

                            // 检查EML文件是否存在
                            var emlFiles = Directory.GetFiles(tempDir, "*.eml", SearchOption.AllDirectories);
                            Log($"EML文件数量: {emlFiles.Length}");

                            var importScript = Path.Combine(exeDir, "script", "create_pst.py");
                            var importArgs = $"\"{importScript}\" \"{outputPstPath}\" \"{tempDir}\"";
                            Log("执行命令: python " + importArgs);

                            var pythonExe = Program.GetPythonExecutable();
                            if (string.IsNullOrEmpty(pythonExe))
                            {
                                throw new Exception("Python环境未找到，请重新安装程序");
                            }

                            var importInfo = Program.CreatePythonStartInfo(pythonExe, importArgs);

                            using (var process = System.Diagnostics.Process.Start(importInfo))
                            {
                                process?.WaitForExit(300000); // 等待5分钟
                                var output = process?.StandardOutput.ReadToEnd();
                                var error = process?.StandardError.ReadToEnd();
                                Log($"Python退出码: {process?.ExitCode}");
                                if (!string.IsNullOrEmpty(output))
                                    Log("导入输出: " + output);
                                if (!string.IsNullOrEmpty(error))
                                    Log("导入错误: " + error);
                            }

                            // 检查PST文件是否存在
                            if (File.Exists(outputPstPath))
                            {
                                var pstInfo = new FileInfo(outputPstPath);
                                Log($"PST文件已创建，大小: {pstInfo.Length} 字节");
                            }
                            else
                            {
                                Log("警告: PST文件未创建!");
                            }
                            Log("导入PST完成");
                        }
                    }
                    finally
                    {
                        // 清理临时目录
                        try { Directory.Delete(tempDir, true); } catch { }
                    }

                    await client.DisconnectAsync(true);
                }

                Log($"收取完成: {outputPstPath}");
                return true;
            }
            catch (Exception ex)
            {
                Log($"收取失败: {ex.GetType().Name}: {ex.Message}");
                Log($"StackTrace: {ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Log($"InnerException: {ex.InnerException.GetType().Name}: {ex.InnerException.Message}");
                }
                throw;
            }
        }

        /// <summary>
        /// 从IMAP服务器收取邮件到EML目录（不转换为PST）
        /// </summary>
        public async Task<bool> FetchEmailsToEmlAsync(
            string email,
            string password,
            string outputDir,
            int maxEmails = 1000,
            DateTime? since = null,
            List<string> selectedFolders = null,
            DateTime? endDate = null)
        {
            try
            {
                Log($"开始收取邮件到EML: {email}");

                var config = await _autoDiscover.AutoDiscoverAsync(email, password);
                config.Email = email;
                config.Password = password;

                Log($"IMAP服务器: {config.Host}:{config.Port}");

                using (var client = new ImapClient())
                {
                    // 禁用证书吊销检查，避免SSL握手卡顿
                    client.CheckCertificateRevocation = false;
                    await client.ConnectAsync(config.Host, config.Port, config.UseSsl);
                    await client.AuthenticateAsync(email, password);

                    var folders = await client.GetFoldersAsync(client.PersonalNamespaces[0]);
                    Log($"发现 {folders.Count} 个文件夹");

                    // 创建输出目录
                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    int totalEmails = 0;

                    // 如果指定了文件夹列表，只处理选中的文件夹
                    if (selectedFolders != null && selectedFolders.Count > 0)
                    {
                        var selectedSet = new HashSet<string>(selectedFolders, StringComparer.OrdinalIgnoreCase);
                        folders = folders.Where(f => selectedSet.Contains(f.Name)).ToList();
                        Log($"将同步以下文件夹: {string.Join(", ", selectedFolders)}");
                    }

                    foreach (var folder in folders)
                    {
                        if (totalEmails >= maxEmails)
                            break;

                        totalEmails = await ProcessFolderToEmlAsync(client, folder, outputDir, totalEmails, maxEmails, since, endDate);
                    }

                    Log($"共收取 {totalEmails} 封邮件到EML目录: {outputDir}");

                    await client.DisconnectAsync(true);
                }

                return true;
            }
            catch (Exception ex)
            {
                Log($"收取失败: {ex.Message}");
                throw;
            }
        }

        private async Task<int> ProcessFolderToEmlAsync(
            ImapClient client,
            IMailFolder folder,
            string outputDir,
            int currentTotal,
            int maxEmails,
            DateTime? since,
            DateTime? endDate = null)
        {
            int totalEmails = currentTotal;
            try
            {
                await folder.OpenAsync(MailKit.FolderAccess.ReadOnly);
                int count = folder.Count;

                if (count == 0)
                {
                    await folder.CloseAsync();
                    return currentTotal;
                }

                string folderName = folder.Name;
                if (string.IsNullOrEmpty(folderName))
                    folderName = "Inbox";

                Log($"处理文件夹: {folderName} ({count} 封邮件)");

                // 创建子目录
                var folderPath = Path.Combine(outputDir, folderName);
                Directory.CreateDirectory(folderPath);

                // 获取已存在的文件列表用于去重
                var existingFiles = new HashSet<string>(Directory.GetFiles(folderPath, "*.eml")
                    .Select(f => Path.GetFileName(f)));
                Log($"检测到已存在 {existingFiles.Count} 个EML文件，将跳过重复邮件");

                // 【优化】使用IMAP服务器端搜索，按日期范围过滤
                IList<UniqueId> uidsToFetch;
                SearchQuery searchQuery = SearchQuery.All;

                // 构建日期搜索条件
                if (since.HasValue || endDate.HasValue)
                {
                    if (since.HasValue && endDate.HasValue)
                    {
                        searchQuery = SearchQuery.DeliveredAfter(since.Value).And(SearchQuery.DeliveredBefore(endDate.Value.AddDays(1)));
                        Log($"使用日期范围搜索: {since.Value:yyyy-MM-dd} 到 {endDate.Value:yyyy-MM-dd}");
                    }
                    else if (since.HasValue)
                    {
                        searchQuery = SearchQuery.DeliveredAfter(since.Value);
                        Log($"使用起始日期搜索: {since.Value:yyyy-MM-dd} 至今");
                    }
                    else if (endDate.HasValue)
                    {
                        searchQuery = SearchQuery.DeliveredBefore(endDate.Value.AddDays(1));
                        Log($"使用截止日期搜索: 至今到 {endDate.Value:yyyy-MM-dd}");
                    }
                }

                Log("正在搜索匹配的邮件UID...");
                var searchResult = folder.Search(searchQuery);
                uidsToFetch = searchResult;
                Log($"服务器返回 {uidsToFetch.Count} 个匹配的UID (count={count})");

                if (uidsToFetch.Count == 0)
                {
                    Log("搜索结果为0，尝试不使用日期过滤重新搜索...");
                    searchResult = folder.Search(SearchQuery.All);
                    uidsToFetch = searchResult;
                    Log($"无日期过滤搜索返回 {uidsToFetch.Count} 个UID");
                }

                if (uidsToFetch.Count == 0)
                {
                    await folder.CloseAsync();
                    Log("没有找到符合条件的邮件");
                    return currentTotal;
                }

                // 按UID从最新到最旧排序
                var sortedUids = uidsToFetch.OrderByDescending(u => u.Id).ToList();
                int totalMatched = sortedUids.Count;
                Log($"文件夹 {folderName} 符合条件 {totalMatched} 封邮件，开始下载...");

                int uidIndex = 0;
                while (uidIndex < sortedUids.Count && totalEmails < maxEmails)
                {
                    var uid = sortedUids[uidIndex];
                    try
                    {
                        var mimeMessage = await folder.GetMessageAsync(uid);

                        if (mimeMessage == null)
                        {
                            uidIndex++;
                            continue;
                        }

                        Log($"[{DateTime.Now:HH:mm:ss}] 下载邮件 {totalEmails + 1}/{maxEmails} (UID: {uid})...");

                        string subject = mimeMessage.Subject ?? "No Subject";
                        foreach (char c in Path.GetInvalidFileNameChars())
                        {
                            subject = subject.Replace(c, '_');
                        }
                        if (subject.Length > 100)
                            subject = subject.Substring(0, 100);

                        string emlFileName = $"{subject}.eml";

                        // 快速检查是否已存在
                        if (existingFiles.Contains(emlFileName))
                        {
                            uidIndex++;
                            continue;
                        }

                        string emlPath = Path.Combine(folderPath, emlFileName);

                        int fileCounter = 1;
                        while (File.Exists(emlPath))
                        {
                            emlFileName = $"{subject}_{fileCounter}.eml";
                            if (existingFiles.Contains(emlFileName))
                            {
                                uidIndex++;
                                break;
                            }
                            emlPath = Path.Combine(folderPath, emlFileName);
                            fileCounter++;
                        }

                        if (File.Exists(emlPath))
                        {
                            uidIndex++;
                            continue;
                        }

                        await mimeMessage.WriteToAsync(emlPath);
                        existingFiles.Add(emlFileName);

                        totalEmails++;

                        _progressCallback?.Invoke(totalEmails, maxEmails);

                        if (totalEmails % 10 == 0)
                        {
                            Log($"已收取: {totalEmails} 封邮件");
                        }
                        // 每50封增加较长延迟，避免服务器限制
                        if (totalEmails % 50 == 0 && totalEmails > 0)
                        {
                            Log("稍作休息...");
                            await Task.Delay(500);
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"UID {uid} 下载失败: {ex.Message}");
                    }

                    uidIndex++;
                    await Task.Delay(100);
                }

                await folder.CloseAsync();
                Log($"文件夹 {folderName} 处理完成: {totalEmails - currentTotal} 封");
            }
            catch (Exception ex)
            {
                Log($"处理文件夹失败: {ex.Message}");
                try { await folder.CloseAsync(); } catch { }
            }
            return totalEmails;
        }

        private async Task<int> ProcessFolderAsync(
            ImapClient client,
            IMailFolder folder,
            string tempDir,
            int currentTotal,
            int maxEmails,
            DateTime? since,
            CancellationToken cancellationToken = default,
            int batchSize = 100,
            DateTime? endDate = null)
        {
            int totalEmails = currentTotal;
            try
            {
                Log($"[ProcessFolderAsync] folder={folder?.Name}, client.IsConnected={client?.IsConnected}");
                await folder.OpenAsync(MailKit.FolderAccess.ReadOnly);
                int count = folder.Count;

                if (count == 0)
                {
                    await folder.CloseAsync();
                    return currentTotal;
                }

                string folderName = folder.Name;
                if (string.IsNullOrEmpty(folderName))
                    folderName = "Inbox";

                Log($"处理文件夹: {folderName} ({count} 封邮件)");

                // 创建子目录
                var folderPath = Path.Combine(tempDir, folderName);
                Directory.CreateDirectory(folderPath);

                // 获取已存在的文件列表用于去重
                var existingFiles = new HashSet<string>(Directory.GetFiles(folderPath, "*.eml")
                    .Select(f => Path.GetFileName(f)));
                Log($"检测到已存在 {existingFiles.Count} 个EML文件，将跳过重复邮件");

                // 【优化】使用IMAP服务器端搜索，按日期范围过滤
                IList<UniqueId> uidsToFetch;
                SearchQuery searchQuery = SearchQuery.All;

                // 构建日期搜索条件
                if (since.HasValue || endDate.HasValue)
                {
                    if (since.HasValue && endDate.HasValue)
                    {
                        // 日期范围: since <= date <= endDate
                        searchQuery = SearchQuery.DeliveredAfter(since.Value).And(SearchQuery.DeliveredBefore(endDate.Value.AddDays(1)));
                        Log($"使用日期范围搜索: {since.Value:yyyy-MM-dd} 到 {endDate.Value:yyyy-MM-dd}");
                    }
                    else if (since.HasValue)
                    {
                        searchQuery = SearchQuery.DeliveredAfter(since.Value);
                        Log($"使用起始日期搜索: {since.Value:yyyy-MM-dd} 至今");
                    }
                    else if (endDate.HasValue)
                    {
                        searchQuery = SearchQuery.DeliveredBefore(endDate.Value.AddDays(1));
                        Log($"使用截止日期搜索: 至今到 {endDate.Value:yyyy-MM-dd}");
                    }
                }

                Log("正在搜索匹配的邮件UID...");
                var searchResult = folder.Search(searchQuery);
                uidsToFetch = searchResult;
                Log($"服务器返回 {uidsToFetch.Count} 个匹配的UID (count={count})");

                if (uidsToFetch.Count == 0)
                {
                    Log("搜索结果为0，尝试不使用日期过滤重新搜索...");
                    searchResult = folder.Search(SearchQuery.All);
                    uidsToFetch = searchResult;
                    Log($"无日期过滤搜索返回 {uidsToFetch.Count} 个UID");
                }

                if (uidsToFetch.Count == 0)
                {
                    await folder.CloseAsync();
                    Log("没有找到符合条件的邮件");
                    return currentTotal;
                }

                // 按UID从最新到最旧排序
                var sortedUids = uidsToFetch.OrderByDescending(u => u.Id).ToList();
                int totalMatched = sortedUids.Count;
                Log($"文件夹 {folderName} 符合条件 {totalMatched} 封邮件，开始下载...");

                int processed = 0;
                int uidIndex = 0;

                while (uidIndex < sortedUids.Count && totalEmails < maxEmails)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var uid = sortedUids[uidIndex];

                    try
                    {
                        // 使用UID获取邮件
                        var mimeMessage = await folder.GetMessageAsync(uid);

                        if (mimeMessage == null)
                        {
                            uidIndex++;
                            continue;
                        }

                        Log($"[{DateTime.Now:HH:mm:ss}] 下载邮件 {totalEmails + 1}/{maxEmails} (UID: {uid})...");

                        string subject = mimeMessage.Subject ?? "No Subject";
                        foreach (char c in Path.GetInvalidFileNameChars())
                        {
                            subject = subject.Replace(c, '_');
                        }
                        if (subject.Length > 100)
                            subject = subject.Substring(0, 100);

                        // 只使用主题命名
                        string emlFileName = $"{subject}.eml";

                        // 快速检查是否已存在
                        if (existingFiles.Contains(emlFileName))
                        {
                            uidIndex++;
                            continue;
                        }

                        string emlPath = Path.Combine(folderPath, emlFileName);

                        int fileCounter = 1;
                        while (File.Exists(emlPath))
                        {
                            emlFileName = $"{subject}_{fileCounter}.eml";
                            if (existingFiles.Contains(emlFileName))
                            {
                                uidIndex++;
                                break;
                            }
                            emlPath = Path.Combine(folderPath, emlFileName);
                            fileCounter++;
                        }

                        if (File.Exists(emlPath))
                        {
                            uidIndex++;
                            continue;
                        }

                        await mimeMessage.WriteToAsync(emlPath);
                        existingFiles.Add(emlFileName);

                        totalEmails++;
                        processed++;

                        _progressCallback?.Invoke(totalEmails, maxEmails);

                        if (totalEmails % 10 == 0)
                        {
                            Log($"已收取: {totalEmails} 封邮件");
                        }
                        // 每50封增加较长延迟，避免服务器限制
                        if (totalEmails % 50 == 0 && totalEmails > 0)
                        {
                            Log("稍作休息...");
                            await Task.Delay(500);
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"UID {uid} 下载失败: {ex.Message}");
                    }

                    uidIndex++;
                    await Task.Delay(100); // 增加延迟避免被限制
                }

                await folder.CloseAsync();
                Log($"文件夹 {folderName} 处理完成: {processed} 封邮件");

                return totalEmails;
            }
            catch (Exception ex)
            {
                Log($"处理文件夹失败: {folder.Name} - {ex.Message}");
                try { await folder.CloseAsync(); } catch { }
                return currentTotal;
            }
        }

        private void Log(string message)
        {
            if (_logCallback == null)
            {
                System.Diagnostics.Debug.WriteLine("[ImapFetchService Log NULL] " + message);
            }
            else
            {
                _logCallback(message);
            }

            var logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs", "mailconverter.log");
            File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [ImapFetch] {message}\n");
        }
    }
}
