using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Serilog;

namespace MailConverter
{
    /// <summary>
    /// 邮件搜索结果
    /// </summary>
    public class MailSearchResult
    {
        public string Id { get; set; }
        public string Subject { get; set; }
        public string From { get; set; }
        public DateTime ReceivedDateTime { get; set; }
        public bool HasAttachments { get; set; }
        public string ToRecipients { get; set; }
        public bool IsSelected { get; set; } = true;
    }

    /// <summary>
    /// Office 365 邮件搜索导出删除服务
    /// 使用 Microsoft Graph API 进行邮件搜索和操作
    /// </summary>
    public class Office365MailSearchService
    {
        private readonly string _accessToken;
        private readonly string _email;
        private static readonly HttpClient _httpClient = new HttpClient();

        public Office365MailSearchService(string accessToken, string email)
        {
            _accessToken = accessToken;
            _email = email;
            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {_accessToken}");
            _httpClient.DefaultRequestHeaders.Add("Accept", "application/json");
        }

        /// <summary>
        /// 搜索邮件
        /// </summary>
        /// <param name="keyword">关键字搜索（主题/正文）</param>
        /// <param name="attachmentName">附件名搜索</param>
        /// <param name="startDate">开始日期</param>
        /// <param name="endDate">结束日期</param>
        /// <param name="maxResults">最大结果数</param>
        /// <returns>邮件列表</returns>
        public async Task<List<MailSearchResult>> SearchEmailsAsync(
            string keyword,
            string attachmentName,
            DateTime? startDate,
            DateTime? endDate,
            int maxResults = 100)
        {
            var results = new List<MailSearchResult>();

            try
            {
                // 构建搜索查询
                var searchParts = new List<string>();

                if (!string.IsNullOrWhiteSpace(keyword))
                {
                    // 关键字搜索 - 使用 subject: 前缀只搜索主题
                    // 如果包含空格则用引号包围
                    if (keyword.Contains(' '))
                        searchParts.Add($"subject:\"{keyword}\"");
                    else
                        searchParts.Add($"subject:{keyword}");
                }

                if (!string.IsNullOrWhiteSpace(attachmentName))
                {
                    // 附件名搜索
                    searchParts.Add($"attachment:{attachmentName}");
                }

                if (searchParts.Count == 0)
                {
                    Log.Warning("搜索关键字和附件名不能同时为空");
                    return results;
                }

                var searchQuery = string.Join(" AND ", searchParts);
                Log.Information("开始搜索邮件, Query: {Query}, MaxResults: {Max}", searchQuery, maxResults);

                // 构建请求 URL
                var url = $"https://graph.microsoft.com/v1.0/users/{_email}/messages" +
                         $"?$search=\"{Uri.EscapeDataString(searchQuery)}\"" +
                         $"&$select=id,subject,from,receivedDateTime,hasAttachments,toRecipients" +
                         $"&$top={maxResults}" +
                         $"&$orderby=receivedDateTime desc";

                // 如果有日期范围，使用筛选代替搜索（搜索+日期筛选较复杂）
                if (startDate.HasValue || endDate.HasValue)
                {
                    // 改用筛选方式结合搜索
                    url = $"https://graph.microsoft.com/v1.0/users/{_email}/messages" +
                         $"?$search=\"{Uri.EscapeDataString(searchQuery)}\"" +
                         $"&$filter=";

                    var dateFilters = new List<string>();
                    if (startDate.HasValue)
                        dateFilters.Add($"receivedDateTime ge {startDate.Value:yyyy-MM-ddTHH:mm:ssZ}");
                    if (endDate.HasValue)
                        dateFilters.Add($"receivedDateTime le {endDate.Value:yyyy-MM-ddTHH:mm:ssZ}");

                    url += string.Join(" and ", dateFilters);
                    url += $"&$select=id,subject,from,receivedDateTime,hasAttachments,toRecipients&$top={maxResults}&$orderby=receivedDateTime desc";
                }

                // 分页获取所有结果
                while (!string.IsNullOrEmpty(url) && results.Count < maxResults)
                {
                    var response = await _httpClient.GetAsync(url);
                    var content = await response.Content.ReadAsStringAsync();

                    if (!response.IsSuccessStatusCode)
                    {
                        Log.Error("搜索邮件失败: {StatusCode}, {Content}", response.StatusCode, content);
                        break;
                    }

                    using (var doc = JsonDocument.Parse(content))
                    {
                        var root = doc.RootElement;

                        if (root.TryGetProperty("value", out var messages))
                        {
                            foreach (var msg in messages.EnumerateArray())
                            {
                                var result = new MailSearchResult
                                {
                                    Id = GetStringProperty(msg, "id"),
                                    Subject = GetStringProperty(msg, "subject"),
                                    From = GetFromAddress(msg),
                                    ReceivedDateTime = GetDateTimeProperty(msg, "receivedDateTime"),
                                    HasAttachments = GetBoolProperty(msg, "hasAttachments"),
                                    ToRecipients = GetToRecipients(msg)
                                };
                                results.Add(result);
                            }
                        }

                        // 获取下一页链接
                        if (root.TryGetProperty("@odata.nextLink", out var nextLink))
                        {
                            url = nextLink.GetString();
                            if (results.Count >= maxResults) break;
                        }
                        else
                        {
                            break;
                        }
                    }
                }

                Log.Information("搜索完成, 找到 {Count} 封邮件", results.Count);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "搜索邮件异常");
            }

            return results;
        }

        /// <summary>
        /// 导出邮件到 PST 文件（通过 EML 临时目录）
        /// </summary>
        public async Task<string> ExportToPstAsync(List<MailSearchResult> emails, string outputPstPath, IProgress<int> progress = null)
        {
            if (emails == null || emails.Count == 0)
            {
                Log.Warning("没有邮件可导出");
                return null;
            }

            try
            {
                Log.Information("开始导出 {Count} 封邮件到 PST", emails.Count);

                // 创建临时目录存放 EML 文件
                var tempDir = Path.Combine(Path.GetTempPath(), $"MailConverter_Export_{Guid.NewGuid():N}");
                Directory.CreateDirectory(tempDir);
                Log.Information("临时目录: {TempDir}", tempDir);

                // 获取邮件内容并保存为 EML
                int count = 0;
                foreach (var email in emails.Where(e => e.IsSelected))
                {
                    var emlContent = await GetMailRawContentAsync(email.Id);
                    if (!string.IsNullOrEmpty(emlContent))
                    {
                        var emlPath = Path.Combine(tempDir, $"{count + 1:D5}_{SanitizeFileName(email.Subject)}.eml");
                        await Task.Run(() => File.WriteAllText(emlPath, emlContent, Encoding.UTF8));
                    }
                    count++;
                    progress?.Report(count * 100 / emails.Count);
                }

                // 使用 Python 脚本创建 PST
                var pythonExe = Program.GetPythonExecutable();
                if (string.IsNullOrEmpty(pythonExe))
                {
                    Log.Error("Python 环境未找到");
                    return null;
                }

                var scriptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "script", "create_pst.py");
                if (!File.Exists(scriptPath))
                {
                    Log.Error("create_pst.py 脚本不存在");
                    return null;
                }

                var arguments = $"\"{scriptPath}\" \"{outputPstPath}\" \"{tempDir}\"";
                var startInfo = Program.CreatePythonStartInfo(pythonExe, $"\"{scriptPath}\" \"{outputPstPath}\" \"{tempDir}\"");

                Log.Information("执行 Python 脚本: {Script}", scriptPath);

                using (var process = Process.Start(startInfo))
                {
                    if (process != null)
                    {
                        process.WaitForExit(600000); // 10 分钟超时
                        var output = await process.StandardOutput.ReadToEndAsync();
                        var error = await process.StandardError.ReadToEndAsync();

                        if (!string.IsNullOrEmpty(output))
                            Log.Information("Python 输出: {Output}", output);
                        if (!string.IsNullOrEmpty(error))
                            Log.Error("Python 错误: {Error}", error);
                    }
                }

                // 清理临时目录
                try
                {
                    Directory.Delete(tempDir, true);
                    Log.Information("已清理临时目录");
                }
                catch (Exception ex)
                {
                    Log.Warning("清理临时目录失败: {Msg}", ex.Message);
                }

                if (File.Exists(outputPstPath))
                {
                    Log.Information("PST 文件已创建: {Path}", outputPstPath);
                    return outputPstPath;
                }
                else
                {
                    Log.Error("PST 文件创建失败");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "导出 PST 异常");
                return null;
            }
        }

        /// <summary>
        /// 导出并删除邮件
        /// </summary>
        public async Task<(int exported, int deleted)> ExportAndDeleteAsync(
            List<MailSearchResult> emails,
            string outputPstPath,
            IProgress<int> progress = null)
        {
            int exported = 0;
            int deleted = 0;

            try
            {
                // 先导出
                var pstPath = await ExportToPstAsync(emails, outputPstPath, progress);
                if (!string.IsNullOrEmpty(pstPath))
                {
                    exported = emails.Count(e => e.IsSelected);

                    // 再删除
                    var selectedEmails = emails.Where(e => e.IsSelected).ToList();
                    deleted = await DeleteEmailsAsync(selectedEmails.Select(e => e.Id).ToList(), progress);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "导出并删除异常");
            }

            return (exported, deleted);
        }

        /// <summary>
        /// 删除邮件
        /// </summary>
        /// <param name="mailIds">邮件 ID 列表</param>
        /// <param name="progress">进度回调</param>
        /// <returns>删除成功的数量</returns>
        public async Task<int> DeleteEmailsAsync(List<string> mailIds, IProgress<int> progress = null)
        {
            int deleted = 0;
            int total = mailIds.Count;

            try
            {
                Log.Information("开始删除 {Count} 封邮件", total);

                for (int i = 0; i < mailIds.Count; i++)
                {
                    var mailId = mailIds[i];

                    // 软删除 - 移动到已删除邮件
                    var url = $"https://graph.microsoft.com/v1.0/users/{_email}/messages/{mailId}";

                    var response = await _httpClient.DeleteAsync(url);

                    if (response.IsSuccessStatusCode)
                    {
                        deleted++;
                        Log.Debug("已删除邮件: {MailId}", mailId);
                    }
                    else
                    {
                        var error = await response.Content.ReadAsStringAsync();
                        Log.Warning("删除邮件失败: {MailId}, Status: {Status}, Error: {Error}",
                            mailId, response.StatusCode, error);
                    }

                    progress?.Report((i + 1) * 100 / total);

                    // 添加小延迟避免限流
                    if (i > 0 && i % 10 == 0)
                    {
                        await Task.Delay(100);
                    }
                }

                Log.Information("删除完成, 成功删除 {Deleted}/{Total} 封", deleted, total);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "删除邮件异常");
            }

            return deleted;
        }

        /// <summary>
        /// 获取邮件原始内容
        /// </summary>
        private async Task<string> GetMailRawContentAsync(string mailId)
        {
            try
            {
                var url = $"https://graph.microsoft.com/v1.0/users/{_email}/messages/{mailId}/$value";
                var response = await _httpClient.GetAsync(url);

                if (response.IsSuccessStatusCode)
                {
                    var bytes = await response.Content.ReadAsByteArrayAsync();
                    return Encoding.UTF8.GetString(bytes);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取邮件内容失败: {MailId}", mailId);
            }
            return null;
        }

        /// <summary>
        /// 获取发件人地址
        /// </summary>
        private string GetFromAddress(JsonElement msg)
        {
            if (msg.TryGetProperty("from", out var from) &&
                from.TryGetProperty("emailAddress", out var emailAddr) &&
                emailAddr.TryGetProperty("address", out var address))
            {
                return address.GetString();
            }
            return "";
        }

        /// <summary>
        /// 获取收件人列表
        /// </summary>
        private string GetToRecipients(JsonElement msg)
        {
            if (msg.TryGetProperty("toRecipients", out var toRecipients))
            {
                var addrs = new List<string>();
                foreach (var r in toRecipients.EnumerateArray())
                {
                    if (r.TryGetProperty("emailAddress", out var emailAddr) &&
                        emailAddr.TryGetProperty("address", out var address))
                    {
                        addrs.Add(address.GetString());
                    }
                }
                return string.Join(", ", addrs);
            }
            return "";
        }

        private string GetStringProperty(JsonElement element, string propertyName)
        {
            if (element.TryGetProperty(propertyName, out var prop))
            {
                return prop.GetString() ?? "";
            }
            return "";
        }

        private DateTime GetDateTimeProperty(JsonElement element, string propertyName)
        {
            if (element.TryGetProperty(propertyName, out var prop))
            {
                if (DateTime.TryParse(prop.GetString(), out var dt))
                    return dt;
            }
            return DateTime.MinValue;
        }

        private bool GetBoolProperty(JsonElement element, string propertyName)
        {
            if (element.TryGetProperty(propertyName, out var prop))
            {
                return prop.GetBoolean();
            }
            return false;
        }

        /// <summary>
        /// 清理文件名中的非法字符
        /// </summary>
        private string SanitizeFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return "untitled";

            var invalid = Path.GetInvalidFileNameChars();
            var sanitized = new StringBuilder();
            foreach (var c in fileName)
            {
                if (!invalid.Contains(c))
                    sanitized.Append(c);
            }
            return sanitized.ToString().Trim();
        }
    }
}
