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
        public string UserId { get; set; }
        public string Subject { get; set; }
        public string From { get; set; }
        public DateTime ReceivedDateTime { get; set; }
        public bool HasAttachments { get; set; }
        public string ToRecipients { get; set; }
        public string AttachmentNames { get; set; }
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

        // 进度回调
        public Action<string, int, int> OnProgress { get; set; }

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
            int maxResults = 100,
            bool onlyWithAttachments = false)
        {
            var results = new List<MailSearchResult>();

            try
            {
                // 构建搜索查询 - 使用 KQL 语法
                var searchParts = new List<string>();

                if (!string.IsNullOrWhiteSpace(keyword))
                {
                    // 如果包含空格则用引号包围，直接搜索
                    if (keyword.Contains(' '))
                        searchParts.Add($"\"{keyword}\"");
                    else
                        searchParts.Add(keyword);
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
                Log.Information("开始搜索所有用户邮件, Query: {Query}, MaxResults: {Max}", searchQuery, maxResults);

                // 首先获取所有用户列表
                var users = await GetAllUsersAsync();
                Log.Information("找到 {Count} 个用户", users.Count);

                if (users.Count == 0)
                {
                    Log.Warning("未找到任何用户");
                    return results;
                }

                // 对每个用户搜索邮件
                int currentUserIndex = 0;
                foreach (var user in users)
                {
                    currentUserIndex++;
                    OnProgress?.Invoke($"正在搜索用户 {currentUserIndex}/{users.Count}: {user}...", currentUserIndex * 100 / users.Count, results.Count);

                    if (results.Count >= maxResults) break;

                    var userResults = await SearchUserMessagesAsync(user, keyword, attachmentName, startDate, endDate, maxResults - results.Count, onlyWithAttachments);
                    results.AddRange(userResults);
                }

                OnProgress?.Invoke($"搜索完成，共 {results.Count} 封邮件", 100, results.Count);
                Log.Information("搜索完成, 共找到 {Count} 封邮件", results.Count);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "搜索邮件异常");
                OnProgress?.Invoke($"搜索异常: {ex.Message}", 0, 0);
            }

            return results;
        }

        /// <summary>
        /// 获取所有用户列表
        /// </summary>
        private async Task<List<string>> GetAllUsersAsync()
        {
            var users = new List<string>();
            var url = "https://graph.microsoft.com/v1.0/users?$select=id,mail&$top=999";

            while (!string.IsNullOrEmpty(url))
            {
                var response = await _httpClient.GetAsync(url);
                var content = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    Log.Error("获取用户列表失败: {StatusCode}, {Content}", response.StatusCode, content);
                    break;
                }

                using (var doc = JsonDocument.Parse(content))
                {
                    var root = doc.RootElement;

                    if (root.TryGetProperty("value", out var userList))
                    {
                        foreach (var user in userList.EnumerateArray())
                        {
                            var userId = GetStringProperty(user, "id");
                            var mail = GetStringProperty(user, "mail");
                            if (!string.IsNullOrEmpty(userId))
                            {
                                users.Add(string.IsNullOrEmpty(mail) ? userId : mail);
                            }
                        }
                    }

                    // 获取下一页链接
                    if (root.TryGetProperty("@odata.nextLink", out var nextLink))
                    {
                        url = nextLink.GetString();
                    }
                    else
                    {
                        break;
                    }
                }
            }

            return users;
        }

        /// <summary>
        /// 搜索指定用户的邮件
        /// </summary>
        private async Task<List<MailSearchResult>> SearchUserMessagesAsync(
            string userId,
            string keyword,
            string attachmentName,
            DateTime? startDate,
            DateTime? endDate,
            int maxResults,
            bool onlyWithAttachments = false)
        {
            var results = new List<MailSearchResult>();

            try
            {
                // 构建过滤条件 - 使用 $filter with contains 替代 $search
                var filters = new List<string>();

                // 日期过滤
                if (startDate.HasValue)
                    filters.Add($"receivedDateTime ge {startDate.Value:yyyy-MM-ddTHH:mm:ssZ}");
                if (endDate.HasValue)
                    filters.Add($"receivedDateTime le {endDate.Value:yyyy-MM-ddTHH:mm:ssZ}");

                // 关键字过滤（搜索主题）
                if (!string.IsNullOrWhiteSpace(keyword))
                {
                    filters.Add($"contains(subject,'{EscapeFilterString(keyword)}')");
                }

                // 仅显示有附件的邮件
                if (onlyWithAttachments)
                {
                    filters.Add("hasAttachments eq true");
                }

                // 构建基础 URL
                var url = $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(userId)}/messages";

                // 添加过滤条件
                if (filters.Count > 0)
                {
                    url += "?$filter=" + string.Join(" and ", filters);
                }

                url += $"&$select=id,subject,from,receivedDateTime,hasAttachments,toRecipients&$top={maxResults}&$orderby=receivedDateTime desc";

                Log.Information("搜索 URL: {Url}", url);

                // 分页获取所有结果
                while (!string.IsNullOrEmpty(url) && results.Count < maxResults)
                {
                    var response = await _httpClient.GetAsync(url);
                    var content = await response.Content.ReadAsStringAsync();

                    if (!response.IsSuccessStatusCode)
                    {
                        // 某些用户可能没有邮件权限，跳过
                        if (response.StatusCode != System.Net.HttpStatusCode.Forbidden)
                        {
                            // 记录详细错误信息
                            string errorDetail = "";
                            try
                            {
                                using (var errDoc = JsonDocument.Parse(content))
                                {
                                    if (errDoc.RootElement.TryGetProperty("error", out var error))
                                    {
                                        if (error.TryGetProperty("message", out var errMsg))
                                            errorDetail = errMsg.GetString();
                                        else if (error.TryGetProperty("code", out var errCode))
                                            errorDetail = errCode.GetString();
                                        else
                                            errorDetail = error.GetRawText();
                                    }
                                    else
                                    {
                                        errorDetail = content;
                                    }
                                }
                            }
                            catch
                            {
                                errorDetail = string.IsNullOrEmpty(content) ? "无响应内容" : content;
                            }
                            Log.Warning("搜索用户 {User} 邮件失败: {StatusCode} - {Error}", userId, response.StatusCode, errorDetail);
                            OnProgress?.Invoke($"搜索用户 {userId} 失败: {errorDetail}", 0, results.Count);
                        }
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
                                    UserId = userId,
                                    Subject = GetStringProperty(msg, "subject"),
                                    From = GetFromAddress(msg),
                                    ReceivedDateTime = GetDateTimeProperty(msg, "receivedDateTime"),
                                    HasAttachments = GetBoolProperty(msg, "hasAttachments"),
                                    ToRecipients = GetToRecipients(msg),
                                    AttachmentNames = ""
                                };

                                // 如果邮件有附件，获取附件名称列表
                                if (result.HasAttachments)
                                {
                                    result.AttachmentNames = await GetAttachmentNamesAsync(userId, result.Id);
                                }

                                // 如果指定了附件名过滤，且邮件有附件，则获取附件列表进行过滤
                                if (!string.IsNullOrWhiteSpace(attachmentName) && result.HasAttachments)
                                {
                                    // 获取邮件附件信息（这里简化处理，假设附件名匹配）
                                    // 实际应该调用获取附件列表的API进行验证
                                    var attachmentFound = await CheckAttachmentExistsAsync(userId, result.Id, attachmentName);
                                    if (!attachmentFound)
                                        continue;
                                }

                                results.Add(result);
                            }
                        }

                        // 获取下一页链接
                        if (root.TryGetProperty("@odata.nextLink", out var nextLink))
                        {
                            url = nextLink.GetString();
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "搜索用户 {User} 邮件异常", userId);
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

                // 创建临时目录存放 EML 文件，放在"搜索导出"子目录下
                var tempBaseDir = Path.Combine(Path.GetTempPath(), $"MailConverter_Export_{Guid.NewGuid():N}");
                var searchExportDir = Path.Combine(tempBaseDir, "搜索导出");
                Directory.CreateDirectory(searchExportDir);
                Log.Information("临时目录: {TempDir}", searchExportDir);

                // 获取邮件内容并保存为 EML
                int count = 0;
                int savedCount = 0;
                foreach (var email in emails.Where(e => e.IsSelected))
                {
                    var emlContent = await GetMailRawContentAsync(email.UserId, email.Id);
                    if (emlContent != null && emlContent.Length > 0)
                    {
                        var emlPath = Path.Combine(searchExportDir, $"{count + 1:D5}_{SanitizeFileName(email.Subject)}.eml");
                        await Task.Run(() => File.WriteAllBytes(emlPath, emlContent));
                        savedCount++;
                        Log.Information("已保存 EML: {Path}, 大小: {Size} bytes", emlPath, emlContent.Length);
                    }
                    else
                    {
                        Log.Warning("邮件内容为空: {Subject} ({Id})", email.Subject, email.Id);
                    }
                    count++;
                    progress?.Report(count * 100 / emails.Count);
                }

                Log.Information("共保存 {SavedCount}/{TotalCount} 封邮件", savedCount, count);

                if (savedCount == 0)
                {
                    Log.Error("没有邮件内容可导出");
                    return null;
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

                var arguments = $"\"{scriptPath}\" \"{outputPstPath}\" \"{tempBaseDir}\"";
                var startInfo = Program.CreatePythonStartInfo(pythonExe, $"\"{scriptPath}\" \"{outputPstPath}\" \"{tempBaseDir}\"");

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
                    Directory.Delete(tempBaseDir, true);
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
                    deleted = await DeleteEmailsAsync(selectedEmails.Select(e => (e.UserId, e.Id)).ToList(), progress);
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
        public async Task<int> DeleteEmailsAsync(List<(string UserId, string MailId)> mailIds, IProgress<int> progress = null)
        {
            int deleted = 0;
            int total = mailIds.Count;

            try
            {
                Log.Information("开始删除 {Count} 封邮件", total);

                for (int i = 0; i < mailIds.Count; i++)
                {
                    var (userId, mailId) = mailIds[i];

                    // 使用标准 DELETE API 移动到已删除邮件文件夹
                    var url = $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(userId)}/messages/{mailId}";

                    var response = await _httpClient.DeleteAsync(url);

                    if (response.IsSuccessStatusCode)
                    {
                        deleted++;
                        Log.Debug("已删除邮件: {MailId} (用户: {UserId})", mailId, userId);
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
        private async Task<byte[]> GetMailRawContentAsync(string userId, string mailId)
        {
            try
            {
                var url = $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(userId)}/messages/{mailId}/$value";
                var response = await _httpClient.GetAsync(url);

                if (response.IsSuccessStatusCode)
                {
                    return await response.Content.ReadAsByteArrayAsync();
                }
                else
                {
                    var error = await response.Content.ReadAsStringAsync();
                    Log.Warning("获取邮件内容失败: {MailId}, Status: {Status}, Error: {Error}", mailId, response.StatusCode, error);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取邮件内容异常: {MailId}", mailId);
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
        /// 获取邮件附件名称列表
        /// </summary>
        private async Task<string> GetAttachmentNamesAsync(string userId, string messageId)
        {
            try
            {
                var url = $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(userId)}/messages/{messageId}/attachments?$select=name";
                var response = await _httpClient.GetAsync(url);

                if (response.IsSuccessStatusCode)
                {
                    var content = await response.Content.ReadAsStringAsync();
                    using (var doc = JsonDocument.Parse(content))
                    {
                        if (doc.RootElement.TryGetProperty("value", out var attachments))
                        {
                            var names = new List<string>();
                            foreach (var att in attachments.EnumerateArray())
                            {
                                var name = GetStringProperty(att, "name");
                                if (!string.IsNullOrEmpty(name))
                                    names.Add(name);
                            }
                            return string.Join(", ", names);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Debug("获取附件名称失败: {MessageId}, {Error}", messageId, ex.Message);
            }
            return "";
        }

        /// <summary>
        /// 检查邮件是否包含指定名称的附件
        /// </summary>
        private async Task<bool> CheckAttachmentExistsAsync(string userId, string messageId, string attachmentName)
        {
            try
            {
                var url = $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(userId)}/messages/{messageId}/attachments?$select=name";
                var response = await _httpClient.GetAsync(url);

                if (response.IsSuccessStatusCode)
                {
                    var content = await response.Content.ReadAsStringAsync();
                    using (var doc = JsonDocument.Parse(content))
                    {
                        if (doc.RootElement.TryGetProperty("value", out var attachments))
                        {
                            foreach (var att in attachments.EnumerateArray())
                            {
                                var name = GetStringProperty(att, "name");
                                if (name.IndexOf(attachmentName, StringComparison.OrdinalIgnoreCase) >= 0)
                                    return true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Debug("检查附件失败: {MessageId}, {Error}", messageId, ex.Message);
            }
            return false;
        }

        /// <summary>
        /// 转义 filter 字符串中的单引号
        /// </summary>
        private string EscapeFilterString(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;
            return input.Replace("'", "''");
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
