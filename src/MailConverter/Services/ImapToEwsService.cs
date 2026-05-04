using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using MailKit;
using MailKit.Search;

namespace MailConverter
{
    /// <summary>
    /// EWS配置
    /// </summary>
    public class EwsConfig
    {
        public string Email { get; set; }
        public string Password { get; set; }
        public string ServerUrl { get; set; }
    }

    /// <summary>
    /// IMAP到EWS同步配置
    /// </summary>
    public class ImapToEwsAccount
    {
        public string ImapEmail { get; set; }
        public string ImapPassword { get; set; }
        public string ImapHost { get; set; }
        public int ImapPort { get; set; } = 993;
        public bool ImapUseSsl { get; set; } = true;
        public string EwsEmail { get; set; }
        public string EwsPassword { get; set; }
        public string EwsServerUrl { get; set; }
        public string TargetFolder { get; set; }  // 目标EWS文件夹，如 "Archive/2024" 或 "已归档"
    }

    /// <summary>
    /// IMAP到EWS同步服务
    /// </summary>
    public class ImapToEwsService
    {
        private Action<string> _logCallback;
        private Action<int, int, string> _progressCallback;
        private int _completedCount;
        private int _totalCount;

        public void SetCallbacks(Action<string> logCallback, Action<int, int, string> progressCallback)
        {
            _logCallback = logCallback;
            _progressCallback = progressCallback;
        }

        /// <summary>
        /// 批量同步IMAP到EWS
        /// </summary>
        public async Task<bool> SyncAccountsAsync(
            List<ImapToEwsAccount> accounts,
            int maxParallel = 5,
            int maxEmailsPerAccount = 1000)
        {
            _totalCount = accounts.Count;
            _completedCount = 0;

            Log($"开始同步 {_totalCount} 个账户到EWS，最大并行数: {maxParallel}");

            using (var semaphore = new SemaphoreSlim(maxParallel))
            {
                var tasks = accounts.Select(async account =>
                {
                    await semaphore.WaitAsync();
                    try
                    {
                        await SyncSingleAccountAsync(account, maxEmailsPerAccount);
                    }
                    finally
                    {
                        semaphore.Release();
                        _completedCount++;
                        _progressCallback?.Invoke(_completedCount, _totalCount, account.ImapEmail);
                    }
                });

                await Task.WhenAll(tasks);
            }

            Log($"全部完成！ {_totalCount} 个账户同步到EWS");
            return true;
        }

        private async Task SyncSingleAccountAsync(ImapToEwsAccount account, int maxEmails)
        {
            string logPrefix = $"[{account.ImapEmail} -> {account.EwsEmail}]";
            Log($"{logPrefix} 开始同步...");

            try
            {
                // 连接IMAP
                var imapConfig = new ImapAccountConfig
                {
                    Email = account.ImapEmail,
                    Password = account.ImapPassword,
                    Host = string.IsNullOrEmpty(account.ImapHost) ? $"imap.{account.ImapEmail.Split('@')[1]}" : account.ImapHost,
                    Port = account.ImapPort,
                    UseSsl = account.ImapUseSsl
                };

                Log($"{logPrefix} 连接IMAP: {imapConfig.Host}:{imapConfig.Port}");

                // 收集邮件
                var emails = new List<MimeKit.MimeMessage>();
                using (var client = new MailKit.Net.Imap.ImapClient())
                {
                    await client.ConnectAsync(imapConfig.Host, imapConfig.Port, imapConfig.UseSsl);
                    await client.AuthenticateAsync(imapConfig.Email, imapConfig.Password);

                    var folders = await client.GetFoldersAsync(client.PersonalNamespaces[0]);
                    Log($"{logPrefix} IMAP文件夹: {folders.Count}");

                    int totalCollected = 0;
                var folderMap = new Dictionary<string, List<(MimeKit.MimeMessage msg, string folderName)>>();

                    foreach (var folder in folders)
                    {
                        if (totalCollected >= maxEmails)
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
                            Log($"{logPrefix} 处理: {folderName} ({count}封)");

                            var uids = folder.Search(SearchQuery.All);
                            var messages = folder.Fetch(uids, MessageSummaryItems.Envelope);

                            foreach (var msg in messages)
                            {
                                if (totalCollected >= maxEmails)
                                    break;

                                try
                                {
                                    var mimeMsg = await folder.GetMessageAsync(msg.UniqueId);

                                    // 按文件夹分组存储
                                    if (!folderMap.ContainsKey(folderName))
                                        folderMap[folderName] = new List<(MimeKit.MimeMessage, string)>();
                                    folderMap[folderName].Add((mimeMsg, folderName));

                                    totalCollected++;

                                    if (totalCollected % 10 == 0)
                                        Log($"{logPrefix} 已收集: {totalCollected}");
                                }
                                catch { }
                            }

                            await folder.CloseAsync();
                        }
                        catch { }
                    }

                    await client.DisconnectAsync(true);

                    Log($"{logPrefix} 共收集 {totalCollected} 封邮件，文件夹: {string.Join(", ", folderMap.Keys)}");
                    Log($"{logPrefix} 开始投递到EWS...");

                    // 按文件夹投递
                    int delivered = 0;
                    foreach (var kvp in folderMap)
                    {
                        string sourceFolderName = kvp.Key;
                        var messages = kvp.Value;

                        try
                        {
                            string targetFolderPath = account.TargetFolder;
                            if (!string.IsNullOrEmpty(targetFolderPath))
                            {
                                // 目标文件夹 + 来源文件夹名
                                targetFolderPath = targetFolderPath + "/" + sourceFolderName;
                            }
                            else
                            {
                                // 直接使用来源文件夹名
                                targetFolderPath = sourceFolderName;
                            }

                            Log($"{logPrefix} 投递文件夹: {sourceFolderName} -> {targetFolderPath} ({messages.Count}封)");

                            foreach (var (mimeMsg, _) in messages)
                            {
                                try
                                {
                                    await DeliverToEws(mimeMsg, account, targetFolderPath);
                                    delivered++;

                                    if (delivered % 10 == 0)
                                    {
                                        Log($"{logPrefix} 已投递: {delivered}/{totalCollected}");
                                        _progressCallback?.Invoke(_completedCount, _totalCount, $"{account.ImapEmail}: {delivered}封");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Log($"{logPrefix} 投递失败: {ex.Message}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Log($"{logPrefix} 文件夹 {sourceFolderName} 投递失败: {ex.Message}");
                        }
                    }

                    Log($"{logPrefix} 同步完成: {delivered} 封");
                }
            }
            catch (Exception ex)
            {
                Log($"{logPrefix} 同步失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 投递邮件到EWS/Office 365
        /// </summary>
        private async Task DeliverToEws(MimeKit.MimeMessage mimeMsg, ImapToEwsAccount account, string targetFolder = null)
        {
            // 使用Python EWS来投递
            await Task.Run(() =>
            {
                // 保存到临时EML
                var tempEml = Path.Combine(Path.GetTempPath(), $"ews_{Guid.NewGuid():N}.eml");
                try
                {
                    mimeMsg.WriteTo(tempEml);

                    // 调用Python EWS脚本
                    var exeDir = AppDomain.CurrentDomain.BaseDirectory;
                    var scriptPath = Path.Combine(exeDir, "script", "ews_deliver.py");

                    if (!File.Exists(scriptPath))
                    {
                        throw new Exception("EWS投递脚本不存在");
                    }

                    var args = $"\"{scriptPath}\" \"{account.EwsEmail}\" \"{account.EwsPassword}\" \"{account.EwsServerUrl}\" \"{tempEml}\"";
                    if (!string.IsNullOrEmpty(targetFolder))
                    {
                        args += $" \"{targetFolder}\"";
                    }

                    var pythonExe = Program.GetPythonExecutable();
                    if (string.IsNullOrEmpty(pythonExe))
                    {
                        throw new Exception("Python环境未找到，请重新安装程序");
                    }

                    var startInfo = Program.CreatePythonStartInfo(pythonExe, args);

                    using (var process = System.Diagnostics.Process.Start(startInfo))
                    {
                        if (process != null)
                        {
                            process.WaitForExit(30000);
                            var output = process.StandardOutput.ReadToEnd();
                            var error = process.StandardError.ReadToEnd();

                            if (!string.IsNullOrEmpty(error) && !error.Contains("Traceback"))
                            {
                                // 只有错误没有traceback才记录
                            }
                        }
                    }
                }
                finally
                {
                    try { File.Delete(tempEml); } catch { }
                }
            });
        }

        private void Log(string message)
        {
            _logCallback?.Invoke(message);
            File.AppendAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs", "mailconverter.log"), $"[ImapToEws] {message}\n");
        }
    }
}
