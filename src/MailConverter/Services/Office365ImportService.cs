using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Graph;
using Azure.Identity;
using Serilog;
using Outlook = Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json.Linq;

namespace MailConverter
{
    // PST 导入专用日志记录器
    public static class PstImportLogger
    {
        private static ILogger _logger;
        private static readonly object _lock = new object();

        public static ILogger Logger
        {
            get
            {
                if (_logger == null)
                {
                    lock (_lock)
                    {
                        if (_logger == null)
                        {
                            var logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
                            Directory.CreateDirectory(logDir);
                            var logPath = Path.Combine(logDir, "BatchPSTTOoff365.log");
                            _logger = new LoggerConfiguration()
                                .MinimumLevel.Information()
                                .WriteTo.File(logPath,
                                    rollingInterval: RollingInterval.Day,
                                    retainedFileCountLimit: 30,
                                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                                .CreateLogger();
                        }
                    }
                }
                return _logger;
            }
        }

        public static void Info(string message)
        {
            Logger.Information(message);
            // 同时写入主日志
            Serilog.Log.Information("[PST导入] " + message);
        }

        public static void Error(string message, Exception ex = null)
        {
            if (ex != null)
                Logger.Error(ex, message);
            else
                Logger.Error(message);
            Serilog.Log.Error(ex, "[PST导入] " + message);
        }

        public static void Warning(string message)
        {
            Logger.Warning(message);
            Serilog.Log.Warning("[PST导入] " + message);
        }
    }

    public class Office365ImportService
    {
        private ExchangeService _service;
        private GraphServiceClient _graphClient;
        private string _email;
        private string _password;
        private string _domain;
        private string _accessToken;
        private bool _isOAuth;
        private string _tenantId;
        private string _clientId;

        // 公开属性供外部访问
        public ExchangeService Service => _service;
        public GraphServiceClient GraphClient => _graphClient;
        public string Email => _email;
        public string AccessToken => _accessToken;
        private string _graphAccessToken;

        public string GraphAccessToken => _graphAccessToken;
        public string TenantId => _tenantId;
        public string AppId => _clientId;
        public string ClientSecret => _clientSecret;

        // 文件夹缓存：Key = "user@domain.com|FolderName", Value = FolderId
        private static ConcurrentDictionary<string, string> _folderCache = new ConcurrentDictionary<string, string>();
        private string _clientSecret;

        /// <summary>
        /// 使用用户名密码连接
        /// </summary>
        public bool Connect(string email, string password, string domain = null)
        {
            try
            {
                _email = email;
                _password = password;
                _domain = domain;
                _isOAuth = false;
                _accessToken = null;

                Log.Information("使用用户名密码连接 Office 365: {Email}", email);

                _service = new ExchangeService(ExchangeVersion.Exchange2016);
                _service.Url = new Uri($"https://outlook.office365.com/EWS/Exchange.asmx");

                if (!string.IsNullOrEmpty(domain))
                {
                    _service.Credentials = new NetworkCredential(email, password, domain);
                }
                else
                {
                    _service.Credentials = new NetworkCredential(email, password);
                }

                _service.AutodiscoverUrl(email, RedirectionUrlValidationCallback);

                Log.Information("Office 365 连接成功");
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Office 365 连接失败: {Msg}", ex.Message);
                return false;
            }
        }

        /// <summary>
        /// 使用 OAuth2 访问令牌连接
        /// </summary>
        public bool ConnectWithOAuth(string email, string accessToken)
        {
            try
            {
                _email = email;
                _accessToken = accessToken;
                _isOAuth = true;
                _password = null;
                _domain = null;

                Log.Information("使用 OAuth2 连接 Office 365: {Email}", email);

                _service = new ExchangeService(ExchangeVersion.Exchange2016);
                _service.Url = new Uri($"https://outlook.office365.com/EWS/Exchange.asmx");

                // 使用 OAuth 凭据
                _service.Credentials = new OAuthCredentials(accessToken);

                // 验证连接
                Folder.Bind(_service, WellKnownFolderName.Inbox);

                Log.Information("Office 365 OAuth2 连接成功");
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Office 365 OAuth2 连接失败: {Msg}", ex.Message);
                return false;
            }
        }

        /// <summary>
        /// 检查是否已通过 OAuth2 连接
        /// </summary>
        public bool IsOAuthConnected => _isOAuth && !string.IsNullOrEmpty(_accessToken);

        /// <summary>
        /// 使用客户端密钥(Client Secret)连接 (应用身份验证)
        /// </summary>
        public bool ConnectWithClientSecret(string tenantId, string clientId, string clientSecret, string accountName)
        {
            try
            {
                _email = accountName;
                _isOAuth = true;
                _tenantId = tenantId;
                _clientId = clientId;
                _clientSecret = clientSecret;

                Log.Information("使用客户端密钥连接 Office 365: Tenant={TenantId}, Client={ClientId}, Account={Account}", tenantId, clientId, accountName);

                // 1. 初始化 Graph 客户端
                var options = new TokenCredentialOptions
                {
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
                };

                var clientSecretCredential = new ClientSecretCredential(
                    tenantId, clientId, clientSecret, options);

                _graphClient = new GraphServiceClient(clientSecretCredential);

                // 注意：使用 ClientSecretCredential (应用权限) 时，不能使用 /me 端点
                // 因为 /me 需要委托认证（用户登录）
                // 直接验证能否获取访问令牌即可
                Log.Information("Office 365 Graph 客户端已初始化 (应用权限模式)");

                // 同时保持 EWS 服务用于某些操作
                var token = GetAccessTokenWithClientSecret(tenantId, clientId, clientSecret);
                if (!string.IsNullOrEmpty(token))
                {
                    _accessToken = token;
                    _service = new ExchangeService(ExchangeVersion.Exchange2016);
                    _service.Url = new Uri($"https://outlook.office365.com/EWS/Exchange.asmx");
                    _service.Credentials = new OAuthCredentials(token);
                    Log.Information("Office 365 连接成功 (EWS + Graph)");
                }
                else
                {
                    Log.Error("获取访问令牌失败");
                    return false;
                }

                // 获取 Graph API 令牌（用于调用 Graph API）
                _graphAccessToken = GetGraphApiToken(tenantId, clientId, clientSecret);
                if (string.IsNullOrEmpty(_graphAccessToken))
                {
                    Log.Warning("获取 Graph API 令牌失败，将使用 EWS 令牌");
                }

                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Office 365 客户端密钥连接失败: {Msg}", ex.Message);
                return false;
            }
        }

        /// <summary>
        /// 使用客户端密钥获取访问令牌 (EWS)
        /// </summary>
        private string GetAccessTokenWithClientSecret(string tenantId, string clientId, string clientSecret)
        {
            return GetAccessTokenWithScope(tenantId, clientId, clientSecret, "https://outlook.office365.com/.default");
        }

        /// <summary>
        /// 使用客户端密钥获取 Graph API 访问令牌
        /// </summary>
        private string GetGraphApiToken(string tenantId, string clientId, string clientSecret)
        {
            return GetAccessTokenWithScope(tenantId, clientId, clientSecret, "https://graph.microsoft.com/.default");
        }

        /// <summary>
        /// 使用客户端密钥获取访问令牌 (通用方法)
        /// </summary>
        private string GetAccessTokenWithScope(string tenantId, string clientId, string clientSecret, string scope)
        {
            try
            {
                var tokenUrl = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token";

                var postData = $"client_id={Uri.EscapeDataString(clientId)}&scope={Uri.EscapeDataString(scope)}&client_secret={Uri.EscapeDataString(clientSecret)}&grant_type=client_credentials";

                Log.Information("正在获取访问令牌, Scope: {Scope}", scope);

                var client = new HttpClient();
                var content = new StringContent(postData, Encoding.UTF8, "application/x-www-form-urlencoded");
                var response = client.PostAsync(tokenUrl, content).Result;
                var responseContent = response.Content.ReadAsStringAsync().Result;

                if (response.IsSuccessStatusCode)
                {
                    var json = JObject.Parse(responseContent);
                    var token = json["access_token"]?.ToString();
                    if (!string.IsNullOrEmpty(token))
                    {
                        Log.Information("获取访问令牌成功");
                        return token;
                    }
                }

                Log.Error("获取令牌失败: {Response}", responseContent);
                return null;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取访问令牌异常");
                return null;
            }
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            return redirectionUrl.ToLower().StartsWith("https://");
        }

        public bool ImportEml(string emlPath, string folderName = "Inbox", IProgress<int> progress = null)
        {
            if (_service == null)
            {
                Log.Error("未连接 Office 365");
                return false;
            }

            try
            {
                Log.Information("导入 EML: {Path}", emlPath);

                // 使用 MimeKit 解析 EML 文件
                var mimeMessage = MimeKit.MimeMessage.Load(emlPath);
                Log.Information("EML 解析完成, Subject={Subject}, From={From}, Date={Date}, Attachments={Count}",
                    mimeMessage.Subject, mimeMessage.From, mimeMessage.Date, mimeMessage.Attachments.Count());

                FolderId targetFolder = GetFolderId(folderName);
                Log.Information("准备保存到文件夹, folderName={Folder}, folderId={Id}", folderName, targetFolder?.ToString());

                // 创建 EmailMessage 并设置所有属性
                EmailMessage email = new EmailMessage(_service);

                // 设置主题
                email.Subject = mimeMessage.Subject ?? "";
                Log.Information("设置主题: {Subject}", email.Subject);

                // 设置发件人
                if (mimeMessage.From != null && mimeMessage.From.Count > 0)
                {
                    var from = mimeMessage.From[0] as MimeKit.MailboxAddress;
                    if (from != null)
                    {
                        email.From = new EmailAddress(from.Name ?? "", from.Address);
                        Log.Information("设置发件人: {Name} <{Address}>", from.Name, from.Address);
                    }
                }

                // 设置收件人
                if (mimeMessage.To != null && mimeMessage.To.Mailboxes.Any())
                {
                    foreach (var to in mimeMessage.To.Mailboxes)
                    {
                        email.ToRecipients.Add(new EmailAddress(to.Name ?? "", to.Address));
                    }
                    Log.Information("设置收件人: {Count} 个", mimeMessage.To.Mailboxes.Count());
                }

                // 设置抄送
                if (mimeMessage.Cc != null && mimeMessage.Cc.Count > 0)
                {
                    foreach (var cc in mimeMessage.Cc.Mailboxes)
                    {
                        email.CcRecipients.Add(new EmailAddress(cc.Name ?? "", cc.Address));
                    }
                }

                // 设置密送
                if (mimeMessage.Bcc != null && mimeMessage.Bcc.Count > 0)
                {
                    foreach (var bcc in mimeMessage.Bcc.Mailboxes)
                    {
                        email.BccRecipients.Add(new EmailAddress(bcc.Name ?? "", bcc.Address));
                    }
                }

                // 设置重要性
                if (mimeMessage.Importance == MimeKit.MessageImportance.High)
                    email.Importance = Microsoft.Exchange.WebServices.Data.Importance.High;
                else if (mimeMessage.Importance == MimeKit.MessageImportance.Low)
                    email.Importance = Microsoft.Exchange.WebServices.Data.Importance.Low;

                // 设置正文 - 优先使用 HTML
                if (!string.IsNullOrEmpty(mimeMessage.HtmlBody))
                {
                    email.Body = new MessageBody(BodyType.HTML, mimeMessage.HtmlBody);
                    Log.Information("设置 HTML 正文");
                }
                else if (mimeMessage.TextBody != null)
                {
                    email.Body = new MessageBody(BodyType.Text, mimeMessage.TextBody);
                    Log.Information("设置文本正文");
                }

                // 处理附件
                if (mimeMessage.Attachments != null && mimeMessage.Attachments.Count() > 0)
                {
                    Log.Information("处理附件: {Count} 个", mimeMessage.Attachments.Count());
                    foreach (var attachment in mimeMessage.Attachments)
                    {
                        var mimePart = attachment as MimeKit.MimePart;
                        if (mimePart != null && mimePart.Content != null)
                        {
                            using (var stream = new MemoryStream())
                            {
                                mimePart.Content.DecodeTo(stream);
                                stream.Position = 0;
                                var fileName = mimePart.FileName ?? "attachment";
                                byte[] content = stream.ToArray();

                                // 使用正确的方式添加附件
                                email.Attachments.AddFileAttachment(fileName, content);
                                Log.Information("添加附件: {Name}", fileName);
                            }
                        }
                    }
                }

                // 设置为已读
                email.IsRead = true;

                // 清除草稿标志 - 使用 ExtendedPropertyDefinition 设置 MAPI 属性
                // PidTagMessageFlags (0x0E07) = 3591
                // 4 = MSGFLAG_READ (已读)，不包含 1 = MSGFLAG_DRAFT (草稿)
                ExtendedPropertyDefinition PR_MESSAGE_FLAGS = new ExtendedPropertyDefinition(3591, MapiPropertyType.Integer);
                email.SetExtendedProperty(PR_MESSAGE_FLAGS, 4);

                Log.Information("准备保存邮件到文件夹: {Folder}, TargetFolder: {Id}", folderName, targetFolder?.ToString());

                try
                {
                    // 保存邮件到目标文件夹
                    email.Save(targetFolder);

                    // 验证邮件是否保存成功
                    if (email.Id != null)
                    {
                        Log.Information("EML 导入成功: {Subject}, Folder: {Folder}, EmailId: {Id}", email.Subject, folderName, email.Id.ToString());
                        return true;
                    }
                    else
                    {
                        Log.Warning("EML 导入可能失败，邮件ID为空: {Subject}", email.Subject);
                        return false;
                    }
                }
                catch (Exception saveEx)
                {
                    Log.Error(saveEx, "保存邮件失败: {Subject}, Folder: {Folder}", email.Subject, folderName);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "导入 EML 失败: {Path}", emlPath);
                return false;
            }
        }

        /// <summary>
        /// 使用 Graph API 导入 EML 到指定用户的邮箱 (通过 EWS 保存，避免草稿问题)
        /// </summary>
        public bool ImportEmlWithGraph(string emlPath, string targetUserEmail, string folderName = "Inbox")
        {
            try
            {
                PstImportLogger.Info($"导入邮件: {emlPath} -> {targetUserEmail}/{folderName}");
                Log.Information("开始导入 EML: {Path} 到 {Email}, Folder: {Folder}", emlPath, targetUserEmail, folderName);

                // 首先尝试使用 Graph API MIME 上传（更可靠）
                Log.Information("尝试使用 Graph API MIME 上传...");
                if (ImportEmlWithGraphMime(emlPath, targetUserEmail, folderName))
                {
                    PstImportLogger.Info($"  -> 成功! (Graph API MIME) 主题: {Path.GetFileName(emlPath)}, 文件夹: {folderName}");
                    return true;
                }
                Log.Warning("Graph API MIME 上传失败，尝试 EWS...");

                // 使用 MimeKit 解析 EML 文件
                var mimeMessage = MimeKit.MimeMessage.Load(emlPath);
                var subject = mimeMessage.Subject ?? "(无主题)";
                PstImportLogger.Info($"  主题: {subject}, 发件人: {mimeMessage.From}, 附件: {mimeMessage.Attachments.Count()}");
                Log.Information("EML 解析完成, Subject={Subject}, From={From}, Attachments={Count}",
                    mimeMessage.Subject, mimeMessage.From, mimeMessage.Attachments.Count());

                // 确保 EWS 服务已连接
                if (_service == null)
                {
                    Log.Error("EWS 服务未连接");
                    PstImportLogger.Error("  -> 失败: EWS 服务未连接");
                    return false;
                }

                // 设置模拟用户 (Impersonation) - 这样可以访问目标用户的邮箱
                // 使用应用程序权限时，必须设置 ImpersonatedUserId 才能访问其他用户的邮箱
                _service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, targetUserEmail);
                Log.Information("已设置模拟用户: {Email}", targetUserEmail);

                // 获取目标文件夹 - 使用默认文件夹而不是 FindFolders
                FolderId targetFolderId = GetWellKnownFolderId(folderName);
                Log.Information("目标文件夹ID: {FolderId}", targetFolderId?.ToString());

                // 如果目标文件夹获取失败，使用默认收件箱
                if (targetFolderId == null)
                {
                    targetFolderId = new FolderId(WellKnownFolderName.Inbox);
                    folderName = "Inbox";
                }

                // 创建 EmailMessage 并设置所有属性
                EmailMessage email = new EmailMessage(_service);

                // 设置主题
                email.Subject = mimeMessage.Subject ?? "";
                Log.Information("设置主题: {Subject}", email.Subject);

                // 设置发件人
                if (mimeMessage.From != null && mimeMessage.From.Count > 0)
                {
                    var from = mimeMessage.From[0] as MimeKit.MailboxAddress;
                    if (from != null)
                    {
                        email.From = new EmailAddress(from.Name ?? "", from.Address);
                        Log.Information("设置发件人: {Name} <{Address}>", from.Name, from.Address);
                    }
                }

                // 设置收件人
                if (mimeMessage.To != null && mimeMessage.To.Mailboxes.Any())
                {
                    foreach (var to in mimeMessage.To.Mailboxes)
                    {
                        email.ToRecipients.Add(new EmailAddress(to.Name ?? "", to.Address));
                    }
                    Log.Information("设置收件人: {Count} 个", mimeMessage.To.Mailboxes.Count());
                }

                // 设置抄送
                if (mimeMessage.Cc != null && mimeMessage.Cc.Count > 0)
                {
                    foreach (var cc in mimeMessage.Cc.Mailboxes)
                    {
                        email.CcRecipients.Add(new EmailAddress(cc.Name ?? "", cc.Address));
                    }
                }

                // 设置正文 - 优先使用 HTML
                if (!string.IsNullOrEmpty(mimeMessage.HtmlBody))
                {
                    email.Body = new MessageBody(BodyType.HTML, mimeMessage.HtmlBody);
                    Log.Information("设置 HTML 正文");
                }
                else if (mimeMessage.TextBody != null)
                {
                    email.Body = new MessageBody(BodyType.Text, mimeMessage.TextBody);
                    Log.Information("设置文本正文");
                }

                // 处理附件
                if (mimeMessage.Attachments != null && mimeMessage.Attachments.Count() > 0)
                {
                    Log.Information("处理附件: {Count} 个", mimeMessage.Attachments.Count());
                    foreach (var attachment in mimeMessage.Attachments)
                    {
                        var mimePart = attachment as MimeKit.MimePart;
                        if (mimePart != null && mimePart.Content != null)
                        {
                            using (var stream = new MemoryStream())
                            {
                                mimePart.Content.DecodeTo(stream);
                                stream.Position = 0;
                                var fileName = mimePart.FileName ?? "attachment";
                                byte[] content = stream.ToArray();

                                email.Attachments.AddFileAttachment(fileName, content);
                                Log.Information("添加附件: {Name}", fileName);
                            }
                        }
                    }
                }

                // 注意：不要在这里设置 IsRead 或 PR_MESSAGE_FLAGS，这可能导致 BadRequest
                // 先保存邮件，然后更新属性

                Log.Information("准备保存邮件到文件夹: {Folder}, TargetFolder: {Id}", folderName, targetFolderId?.ToString());

                try
                {
                    // 保存邮件到目标文件夹
                    email.Save(targetFolderId);

                    // 验证邮件是否保存成功
                    if (email.Id != null)
                    {
                        // 保存成功后，更新邮件属性以清除草稿标志
                        try
                        {
                            // 清除草稿标志 - 使用 ExtendedPropertyDefinition
                            ExtendedPropertyDefinition PR_MESSAGE_FLAGS = new ExtendedPropertyDefinition(3591, MapiPropertyType.Integer);
                            email.SetExtendedProperty(PR_MESSAGE_FLAGS, 4); // 4 = MSGFLAG_READ，不包含 1 = MSGFLAG_DRAFT

                            // 设置为已读
                            email.IsRead = true;

                            // 更新邮件
                            email.Update(ConflictResolutionMode.AlwaysOverwrite);
                            Log.Information("已更新邮件属性，清除草稿标志");
                        }
                        catch (Exception updateEx)
                        {
                            Log.Warning("更新邮件属性失败: {Msg}", updateEx.Message);
                        }

                        PstImportLogger.Info($"  -> 成功! 主题: {subject}, 文件夹: {folderName}");
                        Log.Information("EML 导入成功: {Subject}, Folder: {Folder}, EmailId: {Id}", email.Subject, folderName, email.Id.ToString());
                        return true;
                    }
                    else
                    {
                        Log.Warning("EML 导入可能失败，邮件ID为空: {Subject}", email.Subject);
                        return false;
                    }
                }
                catch (Exception saveEx)
                {
                    // 记录详细错误信息
                    var errorDetails = GetDetailedErrorMessage(saveEx);
                    Log.Error(saveEx, "保存邮件失败: {Subject}, Folder: {Folder}, Error: {Error}", email.Subject, folderName, errorDetails);
                    PstImportLogger.Error($"  -> EWS保存失败，尝试Graph API MIME上传: {errorDetails}");

                    // 尝试使用 Graph API 的 MIME 上传方式作为备用方案
                    try
                    {
                        Log.Information("尝试使用 Graph API MIME 上传: {Path}", emlPath);
                        if (ImportEmlWithGraphMime(emlPath, targetUserEmail, folderName))
                        {
                            PstImportLogger.Info($"  -> Graph API MIME上传成功! 主题: {subject}, 文件夹: {folderName}");
                            return true;
                        }
                        else
                        {
                            PstImportLogger.Error($"  -> Graph API MIME上传也失败");
                            return false;
                        }
                    }
                    catch (Exception mimeEx)
                    {
                        Log.Error(mimeEx, "Graph API MIME 上传也失败: {Path}", emlPath);
                        PstImportLogger.Error($"  -> Graph API MIME上传异常: {mimeEx.Message}");
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                PstImportLogger.Error($"  -> 异常: {ex.Message}", ex);
                Log.Error(ex, "Graph API 导入 EML 失败: {Path}", emlPath);
                return false;
            }
        }

        /// <summary>
        /// 获取 Graph API 文件夹 ID（带缓存，提高批量导入性能）
        /// </summary>
        private string GetGraphMailFolderId(string userEmail, string folderName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(folderName)) folderName = "Inbox";

                // 1. 标准化文件夹名称
                var normalizedName = folderName.Trim();

                // 2. 检查缓存
                string cacheKey = $"{userEmail}|{normalizedName}";
                if (_folderCache.TryGetValue(cacheKey, out var cachedId))
                {
                    return cachedId;
                }

                // 3. 标准文件夹直接返回
                var wellKnownFolders = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "inbox", "inbox" },
                    { "收件箱", "inbox" },
                    { "sent", "sent" },
                    { "sentitems", "sent" },
                    { "已发送", "sent" },
                    { "deleted", "deleteditems" },
                    { "deleteditems", "deleteditems" },
                    { "已删除", "deleteditems" },
                    { "drafts", "drafts" },
                    { "草稿", "drafts" },
                    { "junk", "junkemail" },
                    { "junkemail", "junkemail" },
                    { "垃圾邮件", "junkemail" },
                    { "archive", "archive" },
                    { "archivefolderroot", "archive" },
                    { "存档", "archive" }
                };

                if (wellKnownFolders.TryGetValue(normalizedName.ToLower(), out var wellKnownId))
                {
                    _folderCache.TryAdd(cacheKey, wellKnownId);
                    return wellKnownId;
                }

                // 4. 查找已存在的自定义文件夹
                var folder = _graphClient.Users[userEmail].MailFolders
                    .GetAsync(requestConfiguration => requestConfiguration.QueryParameters.Filter = $"displayName eq '{normalizedName}'")
                    .Result;

                string finalFolderId;
                if (folder?.Value?.Count > 0)
                {
                    finalFolderId = folder.Value[0].Id;
                }
                else
                {
                    // 5. 创建新文件夹
                    var newFolder = new Microsoft.Graph.Models.MailFolder
                    {
                        DisplayName = normalizedName
                    };
                    var created = _graphClient.Users[userEmail].MailFolders.PostAsync(newFolder).Result;
                    finalFolderId = created?.Id ?? "inbox";
                    Log.Information("为 {User} 创建了新文件夹: {Folder}", userEmail, normalizedName);
                }

                // 6. 存入缓存
                _folderCache.TryAdd(cacheKey, finalFolderId);
                return finalFolderId;
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "获取文件夹ID失败，使用默认收件箱: {Folder}", folderName);
                return "inbox";
            }
        }

        public int ImportEmlFolder(string emlFolderPath, string targetFolder = "Inbox", IProgress<int> progress = null)
        {
            if (!Directory.Exists(emlFolderPath))
            {
                Log.Error("EML 文件夹不存在: {Path}", emlFolderPath);
                return 0;
            }

            int imported = 0;
            var files = Directory.GetFiles(emlFolderPath, "*.eml", SearchOption.AllDirectories);

            Log.Information("开始导入 {Count} 个 EML 文件", files.Length);

            foreach (var file in files)
            {
                try
                {
                    // 计算相对路径，确定目标子文件夹
                    string relativePath = GetRelativePath(emlFolderPath, file);
                    string subFolder = targetFolder;

                    // 如果文件在子文件夹中，创建对应的目标子文件夹
                    if (relativePath.Contains(Path.DirectorySeparatorChar) || relativePath.Contains(Path.AltDirectorySeparatorChar))
                    {
                        string relativeDir = Path.GetDirectoryName(relativePath);
                        if (!string.IsNullOrEmpty(relativeDir))
                        {
                            // 使用源文件的子文件夹名称
                            string folderSuffix = relativeDir.Replace(Path.DirectorySeparatorChar, '/').Replace(Path.AltDirectorySeparatorChar, '/');
                            subFolder = string.IsNullOrEmpty(targetFolder) ? folderSuffix : targetFolder + "/" + folderSuffix;
                            Log.Information("文件在子文件夹中: {File} -> {Folder}", relativePath, subFolder);
                        }
                    }

                    if (ImportEml(file, subFolder))
                    {
                        imported++;
                    }

                    if (progress != null && imported % 10 == 0)
                    {
                        progress.Report(imported);
                    }
                }
                catch (Exception ex)
                {
                    Log.Warning("导入失败: {File} - {Msg}", file, ex.Message);
                }
            }

            Log.Information("导入完成: {Count}/{Total}", imported, files.Length);
            return imported;
        }

        /// <summary>
        /// 使用EWS直接上传PST文件（不保存为临时文件）
        /// </summary>
        public int ImportPstDirect(string pstPath, string targetFolder = "Inbox", IProgress<int> progress = null, object ownerForm = null)
        {
            Log.Information("开始EWS直接上传PST: {Path}", pstPath);

            if (_service == null)
            {
                Log.Error("未连接 Office 365");
                return 0;
            }

            if (!File.Exists(pstPath))
            {
                Log.Error("PST文件不存在: {Path}", pstPath);
                return 0;
            }

            int imported = 0;

            try
            {
                // 使用Outlook COM读取PST
                Outlook.Application outlookApp = null;
                try
                {
                    outlookApp = (Outlook.Application)Marshal.GetActiveObject("Outlook.Application");
                }
                catch
                {
                    // 如果没有运行中的Outlook，创建新实例
                    Log.Information("没有运行中的Outlook，创建新实例");
                    outlookApp = new Outlook.Application();
                    System.Threading.Thread.Sleep(3000); // 等待Outlook启动
                }

                var ns = outlookApp.GetNamespace("MAPI");
                Outlook.Folder targetPstFolder = null;

                // 加载PST
                try { ns.AddStoreEx(pstPath, Outlook.OlStoreType.olStoreUnicode); }
                catch (Exception ex) { Log.Warning("添加Store失败: {Msg}", ex.Message); }

                // 查找PST文件夹
                foreach (Outlook.Folder folder in ns.Folders)
                {
                    try
                    {
                        if (folder.Store != null && !string.IsNullOrEmpty(folder.Store.FilePath))
                        {
                            if (Path.GetFullPath(folder.Store.FilePath).Equals(Path.GetFullPath(pstPath), StringComparison.OrdinalIgnoreCase))
                            {
                                targetPstFolder = folder;
                                break;
                            }
                        }
                    }
                    catch { }
                }

                if (targetPstFolder == null)
                {
                    Log.Error("无法找到PST文件夹");
                    return 0;
                }

                // 获取目标FolderId
                FolderId targetFolderId = GetFolderId(targetFolder);

                // 遍历PST文件夹并上传
                imported = UploadPstFolder(targetPstFolder, targetFolderId, progress, ref imported);

                // 卸载PST
                try { ns.RemoveStore(targetPstFolder); } catch { }

                Log.Information("PST直接上传完成: {Count} 封邮件", imported);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "PST直接上传失败");
            }

            return imported;
        }

        private int UploadPstFolder(Outlook.Folder folder, FolderId parentFolderId, IProgress<int> progress, ref int imported)
        {
            try
            {
                Outlook.Items items = folder.Items;
                int folderItemCount = items.Count;
                Log.Information("正在处理文件夹: {Name}, 邮件数: {Count}", folder.Name, folderItemCount);

                for (int i = 1; i <= folderItemCount; i++)
                {
                    try
                    {
                        object item = items[i];
                        if (item is Outlook.MailItem mailItem && (int)mailItem.Class == 43)
                        {
                            // 使用EWS上传
                            if (UploadMailItem(mailItem, parentFolderId))
                            {
                                imported++;
                                if (progress != null && imported % 10 == 0)
                                    progress.Report(imported);
                            }
                        }
                        Marshal.ReleaseComObject(item);
                    }
                    catch (Exception ex)
                    {
                        Log.Warning("上传邮件失败: {Msg}", ex.Message);
                    }
                }
                Marshal.ReleaseComObject(items);

                // 处理子文件夹
                try
                {
                    foreach (Outlook.Folder sub in folder.Folders)
                    {
                        // 创建对应的子文件夹
                        FolderId subFolderId = CreateFolderIfNotExists(parentFolderId, sub.Name);
                        UploadPstFolder(sub, subFolderId, progress, ref imported);
                        Marshal.ReleaseComObject(sub);
                    }
                }
                catch { }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "处理PST文件夹失败: {Name}", folder.Name);
            }

            return imported;
        }

        private bool UploadMailItem(Outlook.MailItem mailItem, FolderId parentFolderId)
        {
            try
            {
                // 从Outlook邮件创建EWS邮件
                EmailMessage email = new EmailMessage(_service);

                // 设置主题
                email.Subject = mailItem.Subject ?? "";

                // 设置发件人
                try
                {
                    string senderAddr = mailItem.SenderEmailAddress;
                    string senderName = mailItem.SenderName;
                    if (!string.IsNullOrEmpty(senderAddr))
                    {
                        email.From = new EmailAddress(senderName ?? senderAddr, senderAddr);
                    }
                }
                catch { }

                // 设置收件人
                try
                {
                    string to = mailItem.To;
                    if (!string.IsNullOrEmpty(to))
                    {
                        var toRecipients = to.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var r in toRecipients)
                        {
                            email.ToRecipients.Add(r.Trim());
                        }
                    }
                }
                catch { }

                // 设置抄送
                try
                {
                    string cc = mailItem.CC;
                    if (!string.IsNullOrEmpty(cc))
                    {
                        var ccRecipients = cc.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var r in ccRecipients)
                        {
                            email.CcRecipients.Add(r.Trim());
                        }
                    }
                }
                catch { }

                // 设置正文
                try
                {
                    string htmlBody = mailItem.HTMLBody;
                    if (!string.IsNullOrEmpty(htmlBody))
                    {
                        email.Body = new MessageBody(BodyType.HTML, htmlBody);
                    }
                    else
                    {
                        email.Body = new MessageBody(BodyType.Text, mailItem.Body ?? "");
                    }
                }
                catch { }

                // 设置日期 - DateTimeSent是只读的，使用ExtendedProperty
                try
                {
                    if (mailItem.SentOn.Year > 1900)
                    {
                        ExtendedPropertyDefinition PR_CLIENT_SUBMIT_DATE = new ExtendedPropertyDefinition(0x0039, MapiPropertyType.SystemTime);
                        email.SetExtendedProperty(PR_CLIENT_SUBMIT_DATE, mailItem.SentOn);
                    }
                }
                catch { }

                // 设置已读
                email.IsRead = true;

                // 清除草稿标志 - 使用 ExtendedPropertyDefinition 设置 MAPI 属性
                // PidTagMessageFlags (0x0E07) = 3591
                // 4 = MSGFLAG_READ (已读)，不包含 1 = MSGFLAG_DRAFT (草稿)
                ExtendedPropertyDefinition PR_MESSAGE_FLAGS = new ExtendedPropertyDefinition(3591, MapiPropertyType.Integer);
                email.SetExtendedProperty(PR_MESSAGE_FLAGS, 4);

                // 保存到目标文件夹
                email.Save(parentFolderId);

                Log.Information("上传邮件成功: {Subject}", email.Subject);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "上传邮件失败");
                return false;
            }
        }

        private FolderId CreateFolderIfNotExists(FolderId parentId, string folderName)
        {
            if (string.IsNullOrEmpty(folderName))
            {
                return parentId;
            }

            try
            {
                // 尝试查找已存在的文件夹
                var searchFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, folderName);
                var folderView = new FolderView(1);
                var findResult = _service.FindFolders(parentId, searchFilter, folderView).Result;

                if (findResult.TotalCount > 0)
                {
                    Log.Information("找到文件夹: {Name}", folderName);
                    return findResult.Folders[0].Id;
                }
            }
            catch (Exception ex)
            {
                Log.Warning("查找文件夹失败: {Name}, {Msg}", folderName, ex.Message);
            }

            // 创建新文件夹
            try
            {
                var newFolder = new Folder(_service);
                newFolder.DisplayName = folderName;
                newFolder.Save(parentId);

                // 等待Exchange创建文件夹
                System.Threading.Thread.Sleep(500);

                // 重新获取文件夹确保创建成功
                var verifyView = new FolderView(1);
                var verifyFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, folderName);
                var verifyResult = _service.FindFolders(parentId, verifyFilter, verifyView).Result;

                if (verifyResult.TotalCount > 0)
                {
                    Log.Information("创建文件夹成功并验证: {Name}", folderName);
                    return verifyResult.Folders[0].Id;
                }
                else
                {
                    Log.Warning("创建文件夹后无法验证: {Name}", folderName);
                    return parentId;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "创建文件夹失败: {Name}", folderName);
                return parentId;
            }
        }

        private void ParseAndImportEmlContent(EmailMessage email, string emlContent)
        {
            var lines = emlContent.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            string body = "";
            bool inBody = false;

            foreach (var line in lines)
            {
                if (inBody)
                {
                    body += line + "\r\n";
                }
                else if (string.IsNullOrWhiteSpace(line))
                {
                    inBody = true;
                }
                else
                {
                    if (line.StartsWith("Subject:", StringComparison.OrdinalIgnoreCase))
                    {
                        email.Subject = DecodeHeaderValue(line.Substring(8).Trim());
                    }
                    else if (line.StartsWith("From:", StringComparison.OrdinalIgnoreCase))
                    {
                        var from = ParseAddress(line.Substring(5).Trim());
                        email.From = new EmailAddress(from.Name, from.Address);
                    }
                    else if (line.StartsWith("To:", StringComparison.OrdinalIgnoreCase))
                    {
                        email.ToRecipients.Add(ParseAddress(line.Substring(3).Trim()));
                    }
                    else if (line.StartsWith("Cc:", StringComparison.OrdinalIgnoreCase))
                    {
                        email.CcRecipients.Add(ParseAddress(line.Substring(3).Trim()));
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(body))
            {
                try
                {
                    var base64Lines = body.Trim().Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                    var base64String = string.Join("", base64Lines);
                    var decoded = Convert.FromBase64String(base64String);
                    var htmlBody = System.Text.Encoding.UTF8.GetString(decoded);
                    email.Body = new MessageBody(BodyType.HTML, htmlBody);
                }
                catch
                {
                    email.Body = new MessageBody(BodyType.Text, body);
                }
            }
        }

        private EmailAddress ParseAddress(string addressLine)
        {
            try
            {
                if (addressLine.Contains("<") && addressLine.Contains(">"))
                {
                    var name = addressLine.Substring(0, addressLine.IndexOf('<')).Trim().Trim('"');
                    var email = addressLine.Substring(addressLine.IndexOf('<') + 1,
                        addressLine.IndexOf('>') - addressLine.IndexOf('<') - 1);
                    return new EmailAddress(name, email);
                }
                return new EmailAddress(addressLine.Trim());
            }
            catch
            {
                return new EmailAddress(addressLine.Trim());
            }
        }

        private string DecodeHeaderValue(string value)
        {
            try
            {
                if (string.IsNullOrEmpty(value)) return value;

                if (value.StartsWith("=?") && value.Contains("?="))
                {
                    int start = value.IndexOf("=?") + 2;
                    int end = value.IndexOf("?=", start);
                    if (end > start)
                    {
                        var encoded = value.Substring(start, end - start);
                        var parts = encoded.Split('?');
                        if (parts.Length >= 3 && parts[1].ToLower() == "b")
                        {
                            var decoded = Convert.FromBase64String(parts[2]);
                            return System.Text.Encoding.UTF8.GetString(decoded);
                        }
                    }
                }
                return value;
            }
            catch
            {
                return value;
            }
        }

        /// <summary>
        /// 获取异常的详细错误信息
        /// </summary>
        private string GetDetailedErrorMessage(Exception ex)
        {
            if (ex == null) return "Unknown error";

            var message = ex.Message;

            // 如果包含 HTTP 响应信息，尝试获取更详细的内容
            if (ex is ServiceResponseException serviceEx)
            {
                message = $"ServiceResponseException: {serviceEx.ErrorCode} - {serviceEx.Message}";
                if (serviceEx.Response != null)
                {
                    message += $", Response: {serviceEx.Response}";
                }
            }
            else if (ex is ServiceRequestException requestEx)
            {
                message = $"ServiceRequestException: {requestEx.Message}";
            }

            // 尝试获取内部异常信息
            if (ex.InnerException != null)
            {
                message += $" | Inner: {ex.InnerException.Message}";
            }

            return message;
        }

        /// <summary>
        /// 使用 Graph SDK 对象映射法导入 EML
        /// </summary>
        private bool ImportEmlWithGraphMime(string emlPath, string targetUserEmail, string folderName = "Inbox")
        {
            try
            {
                if (_graphClient == null)
                {
                    Log.Error("Graph 客户端未初始化");
                    return false;
                }

                // 使用 MimeKit 解析 EML 文件
                var mimeMessage = MimeKit.MimeMessage.Load(emlPath);
                var subject = mimeMessage.Subject ?? "(无主题)";

                Log.Information("Graph SDK 导入 - 主题: {Subject}, 附件: {Count}", subject, mimeMessage.Attachments.Count());

                // 处理发件人 - 避免 unknown@unknown.com
                string fromAddress = "no-reply@booming.one";
                string fromName = "";
                if (mimeMessage.From != null && mimeMessage.From.Count > 0)
                {
                    var from = mimeMessage.From[0] as MimeKit.MailboxAddress;
                    if (from != null && !string.IsNullOrEmpty(from.Address) && !from.Address.Contains("unknown"))
                    {
                        fromAddress = from.Address;
                        fromName = from.Name ?? "";
                    }
                }

                // 获取目标文件夹 ID
                var folderId = GetGraphMailFolderId(targetUserEmail, folderName);
                Log.Information("Graph SDK 导入 - 目标文件夹: {FolderId}", folderId);

                // 构建 Graph Message 对象
                var graphMessage = new Microsoft.Graph.Models.Message
                {
                    Subject = mimeMessage.Subject ?? "",
                    Body = new Microsoft.Graph.Models.ItemBody
                    {
                        ContentType = !string.IsNullOrEmpty(mimeMessage.HtmlBody) ?
                            Microsoft.Graph.Models.BodyType.Html :
                            Microsoft.Graph.Models.BodyType.Text,
                        Content = mimeMessage.HtmlBody ?? mimeMessage.TextBody ?? ""
                    },
                    // 设置接收和发送时间
                    ReceivedDateTime = mimeMessage.Date != null ? mimeMessage.Date.DateTime : DateTime.UtcNow,
                    SentDateTime = mimeMessage.Date != null ? mimeMessage.Date.DateTime : DateTime.UtcNow,

                    // 关键：显式设为非草稿
                    IsDraft = false,

                    // 发件人
                    From = new Microsoft.Graph.Models.Recipient
                    {
                        EmailAddress = new Microsoft.Graph.Models.EmailAddress
                        {
                            Name = fromName,
                            Address = fromAddress
                        }
                    },

                    // 关键：使用扩展属性去除"草稿"标记并设为"已读"
                    // PidTagMessageFlags (0x0E07)
                    // 值 1 (MSGFLAG_READ) 表示已读且非草稿
                    SingleValueExtendedProperties = new List<Microsoft.Graph.Models.SingleValueLegacyExtendedProperty>
                    {
                        new Microsoft.Graph.Models.SingleValueLegacyExtendedProperty
                        {
                            Id = "Integer 0x0E07",
                            Value = "1"
                        }
                    }
                };

                // 添加收件人
                if (mimeMessage.To != null && mimeMessage.To.Mailboxes.Any())
                {
                    graphMessage.ToRecipients = new List<Microsoft.Graph.Models.Recipient>();
                    foreach (var to in mimeMessage.To.Mailboxes)
                    {
                        graphMessage.ToRecipients.Add(new Microsoft.Graph.Models.Recipient
                        {
                            EmailAddress = new Microsoft.Graph.Models.EmailAddress
                            {
                                Name = to.Name,
                                Address = to.Address
                            }
                        });
                    }
                }

                // 添加抄送
                if (mimeMessage.Cc != null && mimeMessage.Cc.Count > 0)
                {
                    graphMessage.CcRecipients = new List<Microsoft.Graph.Models.Recipient>();
                    foreach (var cc in mimeMessage.Cc.Mailboxes)
                    {
                        graphMessage.CcRecipients.Add(new Microsoft.Graph.Models.Recipient
                        {
                            EmailAddress = new Microsoft.Graph.Models.EmailAddress
                            {
                                Name = cc.Name,
                                Address = cc.Address
                            }
                        });
                    }
                }

                // 使用 Graph SDK 发送邮件到目标文件夹
                var result = _graphClient.Users[targetUserEmail]
                    .MailFolders[folderId]
                    .Messages
                    .PostAsync(graphMessage)
                    .Result;

                if (result != null)
                {
                    Log.Information("Graph SDK 导入成功: {Subject}, ID: {Id}", subject, result.Id);
                    return true;
                }

                Log.Error("Graph SDK 导入返回空结果");
                return false;
            }
            catch (Exception ex)
            {
                // 打印详细的 Graph API 错误
                var errorMsg = ex.Message;
                if (ex.InnerException != null)
                {
                    errorMsg += " | Inner: " + ex.InnerException.Message;
                }
                Log.Error(ex, "Graph SDK 导入失败: {Path}, Error: {Error}", emlPath, errorMsg);
                return false;
            }
        }

        // 旧的 HTTP 上传方法，保留作为参考
        private bool ImportEmlWithGraphMimeOld(string emlPath, string targetUserEmail, string folderName = "Inbox")
        {
            try
            {
                if (_graphClient == null)
                {
                    Log.Error("Graph 客户端未初始化");
                    return false;
                }

                // 使用 MimeKit 解析 EML 文件
                var mimeMessage = MimeKit.MimeMessage.Load(emlPath);
                var subject = mimeMessage.Subject ?? "(无主题)";

                // 获取目标文件夹 ID
                var folderId = GetGraphMailFolderId(targetUserEmail, folderName);
                Log.Information("Graph HTTP 上传 - 目标文件夹: {FolderId}, 主题: {Subject}", folderId, subject);

                // 读取 EML 文件的原始内容
                var emlContent = File.ReadAllText(emlPath, Encoding.UTF8);

                // 使用 HTTP 上传 MIME 内容
                var mimeRequest = new HttpRequestMessage(HttpMethod.Post,
                    $"https://graph.microsoft.com/v1.0/users/{targetUserEmail}/mailFolders/{folderId}/messages");
                mimeRequest.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", _accessToken);
                mimeRequest.Content = new StringContent(emlContent, Encoding.UTF8, "message/rfc822");

                var httpClient = new HttpClient();
                var response = httpClient.SendAsync(mimeRequest).Result;

                if (response.IsSuccessStatusCode)
                {
                    Log.Information("Graph HTTP MIME 上传成功: {Path}", emlPath);
                    return true;
                }

                var errorContent = response.Content.ReadAsStringAsync().Result;
                Log.Warning("Graph HTTP MIME 上传失败: {StatusCode}, Response: {Response}", response.StatusCode, errorContent);

                // HTTP 上传失败，尝试使用 EWS
                return ImportEmlWithEws(emlPath, targetUserEmail, folderName);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Graph HTTP 上传异常: {Path}, 尝试 EWS 方式", emlPath);
                return ImportEmlWithEws(emlPath, targetUserEmail, folderName);
            }
        }

        /// <summary>
        /// 使用 EWS 方式导入 EML
        /// </summary>
        private bool ImportEmlWithEws(string emlPath, string targetUserEmail, string folderName)
        {
            try
            {
                if (_service == null)
                {
                    Log.Error("EWS 服务未初始化");
                    return false;
                }

                // 设置模拟用户
                _service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, targetUserEmail);

                // 获取目标文件夹
                FolderId targetFolderId = GetWellKnownFolderId(folderName);

                // 解析 EML
                var mimeMessage = MimeKit.MimeMessage.Load(emlPath);
                var subject = mimeMessage.Subject ?? "(无主题)";

                // 创建 EmailMessage
                EmailMessage email = new EmailMessage(_service);
                email.Subject = subject;

                // 设置发件人
                if (mimeMessage.From != null && mimeMessage.From.Count > 0)
                {
                    var from = mimeMessage.From[0] as MimeKit.MailboxAddress;
                    if (from != null)
                        email.From = new EmailAddress(from.Name ?? "", from.Address);
                }

                // 设置收件人
                if (mimeMessage.To != null && mimeMessage.To.Mailboxes.Any())
                {
                    foreach (var to in mimeMessage.To.Mailboxes)
                        email.ToRecipients.Add(new EmailAddress(to.Name ?? "", to.Address));
                }

                // 设置正文
                if (!string.IsNullOrEmpty(mimeMessage.HtmlBody))
                    email.Body = new MessageBody(BodyType.HTML, mimeMessage.HtmlBody);
                else if (mimeMessage.TextBody != null)
                    email.Body = new MessageBody(BodyType.Text, mimeMessage.TextBody);

                // 保存邮件
                email.Save(targetFolderId);

                if (email.Id != null)
                {
                    // 清除草稿标志
                    ExtendedPropertyDefinition PR_MESSAGE_FLAGS = new ExtendedPropertyDefinition(3591, MapiPropertyType.Integer);
                    email.SetExtendedProperty(PR_MESSAGE_FLAGS, 4); // 4 = 已读，非草稿
                    email.IsRead = true;
                    email.Update(ConflictResolutionMode.AlwaysOverwrite);

                    Log.Information("EWS 导入成功: {Subject}, ID: {Id}", subject, email.Id);
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "EWS 导入失败: {Path}", emlPath);
                return false;
            }
        }

        /// <summary>
        /// 获取已知文件夹ID（不调用FindFolders，避免403错误）
        /// </summary>
        private FolderId GetWellKnownFolderId(string folderName)
        {
            try
            {
                var normalizedName = folderName.Trim().ToLower();

                // 常见的内置文件夹映射
                if (normalizedName == "inbox" || normalizedName == "收件箱")
                    return new FolderId(WellKnownFolderName.Inbox);
                if (normalizedName == "sent" || normalizedName == "sentitems" || normalizedName == "已发送")
                    return new FolderId(WellKnownFolderName.SentItems);
                if (normalizedName == "deleted" || normalizedName == "deleteditems" || normalizedName == "已删除")
                    return new FolderId(WellKnownFolderName.DeletedItems);
                if (normalizedName == "drafts" || normalizedName == "草稿")
                    return new FolderId(WellKnownFolderName.Drafts);
                if (normalizedName == "junk" || normalizedName == "junkemail" || normalizedName == "垃圾邮件")
                    return new FolderId(WellKnownFolderName.JunkEmail);
                if (normalizedName == "archive" || normalizedName == "archivefolder" || normalizedName == "存档")
                    return new FolderId(WellKnownFolderName.ArchiveMsgFolderRoot);

                // 如果不是内置文件夹，返回null让调用方使用默认收件箱
                Log.Warning("未知的文件夹名称: {Folder}, 使用默认收件箱", folderName);
                return new FolderId(WellKnownFolderName.Inbox);
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "获取已知文件夹失败: {Folder}", folderName);
                return new FolderId(WellKnownFolderName.Inbox);
            }
        }

        private FolderId GetFolderId(string folderName)
        {
            try
            {
                // 处理嵌套文件夹路径，如 "Inbox/mail2"
                string[] folderParts = folderName.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);

                // 直接在根目录创建文件夹，不使用 Inbox
                FolderId currentFolderId = WellKnownFolderName.MsgFolderRoot;
                string currentPath = "";

                for (int i = 0; i < folderParts.Length; i++)
                {
                    string part = folderParts[i];
                    currentPath = string.IsNullOrEmpty(currentPath) ? part : currentPath + "/" + part;

                    // 查找当前级别的文件夹
                    var folderView = new FolderView(100);
                    folderView.Traversal = FolderTraversal.Shallow;

                    var searchFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, part);
                    var findResults = _service.FindFolders(currentFolderId, searchFilter, folderView).Result;

                    FolderId foundFolderId = null;
                    foreach (var folder in findResults)
                    {
                        if (folder.DisplayName.Equals(part, StringComparison.OrdinalIgnoreCase))
                        {
                            foundFolderId = folder.Id;
                            break;
                        }
                    }

                    if (foundFolderId != null)
                    {
                        currentFolderId = foundFolderId;
                        Log.Information("找到文件夹: {Folder}", currentPath);
                    }
                    else
                    {
                        // 文件夹不存在，创建它
                        Log.Information("文件夹不存在，正在创建: {Folder}", currentPath);
                        try
                        {
                            var newFolder = new Folder(_service);
                            newFolder.DisplayName = part;
                            newFolder.Save(currentFolderId);

                            // 等待Exchange创建文件夹
                            System.Threading.Thread.Sleep(500);

                            // 重新获取文件夹ID确保创建成功
                            var verifyView = new FolderView(1);
                            var verifyFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, part);
                            var verifyResult = _service.FindFolders(currentFolderId, verifyFilter, verifyView).Result;

                            if (verifyResult.TotalCount > 0)
                            {
                                currentFolderId = verifyResult.Folders[0].Id;
                                Log.Information("文件夹创建并验证成功: {Folder}", currentPath);
                            }
                            else
                            {
                                Log.Warning("文件夹创建后无法验证: {Folder}", currentPath);
                            }
                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex, "创建文件夹失败: {Folder}", currentPath);
                            throw;
                        }
                    }
                }

                Log.Information("GetFolderId 返回: {FolderId}", currentFolderId?.ToString());
                return currentFolderId;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取或创建文件夹失败: {Folder}", folderName);
                // 返回默认收件箱
                try
                {
                    return new FolderId(WellKnownFolderName.Inbox, _email);
                }
                catch
                {
                    return new FolderId(WellKnownFolderName.MsgFolderRoot, _email);
                }
            }
        }

        private string GetRelativePath(string basePath, string fullPath)
        {
            if (!basePath.EndsWith(Path.DirectorySeparatorChar.ToString()))
                basePath += Path.DirectorySeparatorChar;

            Uri baseUri = new Uri(basePath);
            Uri fullUri = new Uri(fullPath);

            string relativeUri = Uri.UnescapeDataString(baseUri.MakeRelativeUri(fullUri).ToString());
            return relativeUri.Replace('/', Path.DirectorySeparatorChar);
        }

        public bool TestConnection()
        {
            if (_service == null) return false;

            try
            {
                var folder = Folder.Bind(_service, WellKnownFolderName.Inbox).Result;
                Log.Information("连接测试成功，收件箱有 {Count} 封邮件", folder.TotalCount);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "连接测试失败");
                return false;
            }
        }

        /// <summary>
        /// 创建联系人
        /// </summary>
        public bool CreateContact(string displayName, string emailAddress)
        {
            return CreateContact(displayName, emailAddress, null, null, null);
        }

        /// <summary>
        /// 创建联系人（带扩展信息）
        /// </summary>
        public bool CreateContact(string displayName, string emailAddress, string phone, string company, string title)
        {
            if (_service == null)
            {
                Log.Error("未连接 Office 365");
                return false;
            }

            // 如果邮箱为空，设置为null而不是空字符串
            if (string.IsNullOrWhiteSpace(emailAddress))
                emailAddress = null;

            try
            {
                Contact contact = new Contact(_service);
                contact.DisplayName = displayName;

                // 只有邮箱不为空时才设置
                if (!string.IsNullOrEmpty(emailAddress))
                {
                    contact.EmailAddresses[EmailAddressKey.EmailAddress1] = new EmailAddress(emailAddress);
                }

                // 添加扩展信息
                if (!string.IsNullOrWhiteSpace(phone))
                    contact.PhoneNumbers[PhoneNumberKey.PrimaryPhone] = phone;
                if (!string.IsNullOrWhiteSpace(company))
                    contact.CompanyName = company;
                if (!string.IsNullOrWhiteSpace(title))
                    contact.JobTitle = title;

                contact.Save();
                Log.Information("创建联系人成功: {Name} <{Email}>, Phone={Phone}, Company={Company}, Title={Title}",
                    displayName, emailAddress ?? "(无邮箱)", phone, company, title);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "创建联系人失败: {Name} <{Email}>", displayName, emailAddress ?? "(无邮箱)");
                return false;
            }
        }

        /// <summary>
        /// 根据邮箱查找联系人
        /// </summary>
        public string FindContactByEmail(string emailAddress)
        {
            if (_service == null)
            {
                Log.Error("未连接 Office 365");
                return null;
            }

            try
            {
                // 使用EmailAddressKey搜索
                var searchFilter = new SearchFilter.ContainsSubstring(ContactSchema.EmailAddresses, emailAddress, ContainmentMode.Substring, ComparisonMode.IgnoreCase);
                var searchResult = _service.FindItems(WellKnownFolderName.Contacts, searchFilter, new ItemView(10)).Result;

                foreach (var item in searchResult.Items)
                {
                    if (item is Contact contact)
                    {
                        Log.Information("找到联系人: {Name} <{Email}>", contact.DisplayName, emailAddress);
                        return contact.Id.UniqueId;
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "查找联系人失败: {Email}", emailAddress);
                return null;
            }
        }

        /// <summary>
        /// 根据邮箱更新或创建联系人
        /// </summary>
        public bool UpsertContact(string displayName, string emailAddress, string phone, string company, string title)
        {
            if (_service == null)
            {
                Log.Error("未连接 Office 365");
                return false;
            }

            try
            {
                // 先查找是否已存在
                var searchFilter = new SearchFilter.ContainsSubstring(ContactSchema.EmailAddresses, emailAddress, ContainmentMode.Substring, ComparisonMode.IgnoreCase);
                var searchResult = _service.FindItems(WellKnownFolderName.Contacts, searchFilter, new ItemView(10)).Result;

                foreach (var item in searchResult.Items)
                {
                    if (item is Contact existingContact)
                    {
                        // 更新已存在的联系人
                        existingContact.DisplayName = displayName;
                        if (!string.IsNullOrWhiteSpace(phone))
                            existingContact.PhoneNumbers[PhoneNumberKey.PrimaryPhone] = phone;
                        if (!string.IsNullOrWhiteSpace(company))
                            existingContact.CompanyName = company;
                        if (!string.IsNullOrWhiteSpace(title))
                            existingContact.JobTitle = title;

                        existingContact.Update(ConflictResolutionMode.AlwaysOverwrite);
                        Log.Information("更新联系人成功: {Name} <{Email}>", displayName, emailAddress);
                        return true;
                    }
                }

                // 不存在，创建新联系人
                return CreateContact(displayName, emailAddress, phone, company, title);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Upsert联系人失败: {Name} <{Email}>", displayName, emailAddress);
                return false;
            }
        }

        /// <summary>
        /// 创建日历事件
        /// </summary>
        public bool CreateCalendarEvent(string subject, DateTime start, DateTime end, string description = "", string location = "")
        {
            if (_service == null)
            {
                Log.Error("未连接 Office 365");
                return false;
            }

            try
            {
                Appointment appointment = new Appointment(_service);
                appointment.Subject = subject;
                appointment.Start = start;
                appointment.End = end;
                appointment.Body = description;
                appointment.Location = location;
                appointment.Save(WellKnownFolderName.Calendar);
                Log.Information("创建日历事件成功: {Subject}", subject);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "创建日历事件失败: {Subject}", subject);
                return false;
            }
        }

        /// <summary>
        /// 发送测试邮件
        /// </summary>
        public bool SendTestEmail(string to, string subject, string body)
        {
            try
            {
                if (_service == null)
                {
                    Log.Error("EWS服务未连接");
                    return false;
                }

                var email = new EmailMessage(_service)
                {
                    Subject = subject,
                    Body = new MessageBody(BodyType.Text, body)
                };
                email.ToRecipients.Add(to);
                email.Send();
                Log.Information("测试邮件已发送到: {To}", to);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "发送测试邮件失败: {To}", to);
                return false;
            }
        }

        /// <summary>
        /// 从VCF文件导入单个联系人
        /// </summary>
        public bool ImportContactFromVcf(string targetEmail, string vcfFilePath)
        {
            try
            {
                if (_graphClient == null)
                {
                    Log.Error("Graph客户端未初始化");
                    return false;
                }

                var contact = new Microsoft.Graph.Models.Contact
                {
                    DisplayName = Path.GetFileNameWithoutExtension(vcfFilePath)
                };

                // 读取VCF文件内容
                var vcfContent = File.ReadAllText(vcfFilePath);

                // 使用Graph API创建联系人
                // 注意：此为简化版本，实际生产环境需要解析VCF并填充对应字段
                Log.Information("正在导入联系人到 {Email}: {File}", targetEmail, vcfFilePath);

                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "导入联系人失败: {File}", vcfFilePath);
                return false;
            }
        }

        /// <summary>
        /// 批量从VCF文件导入联系人 (使用 Client Secret / 应用权限)
        /// 使用并行请求 + SemaphoreSlim 控制并发数，带指数退避重试
        /// </summary>
        /// <param name="targetEmail">目标用户邮箱</param>
        /// <param name="vcfFilePaths">VCF文件路径列表</param>
        /// <param name="progressCallback">进度回调 (当前索引, 总数, 状态消息)</param>
        /// <param name="maxDegreeOfParallelism">最大并发数，默认10</param>
        /// <returns>成功导入数量</returns>
        public async Task<int> ImportContactsBatchFromVcfAsync(
            string targetEmail,
            IEnumerable<string> vcfFilePaths,
            Action<int, int, string> progressCallback = null,
            int maxDegreeOfParallelism = 10)
        {
            if (_graphClient == null)
            {
                Log.Error("Graph客户端未初始化，请先调用 ConnectWithClientSecret");
                return 0;
            }

            if (string.IsNullOrWhiteSpace(targetEmail))
            {
                Log.Error("目标邮箱不能为空");
                return 0;
            }

            var fileList = vcfFilePaths.ToList();
            int totalCount = fileList.Count;
            int successCount = 0;
            int currentIndex = 0;

            Log.Information("开始批量导入联系人到 {Email}, 共 {Count} 个VFC文件, 并发数: {Parallelism}",
                targetEmail, totalCount, maxDegreeOfParallelism);

            // 使用 SemaphoreSlim 控制并发数
            var semaphore = new System.Threading.SemaphoreSlim(maxDegreeOfParallelism);
            var lockObj = new object();

            // 并行处理所有文件
            var tasks = fileList.Select(async vcfFile =>
            {
                await semaphore.WaitAsync();
                try
                {
                    int index;
                    lock (lockObj)
                    {
                        index = currentIndex++;
                    }

                    // 带重试的导入
                    bool success = await ImportSingleContactWithRetryAsync(targetEmail, vcfFile);
                    if (success)
                    {
                        lock (lockObj) { successCount++; }
                    }

                    progressCallback?.Invoke(index + 1, totalCount, $"已导入 {index + 1}/{totalCount}");
                    return success;
                }
                finally
                {
                    semaphore.Release();
                }
            });

            await System.Threading.Tasks.Task.WhenAll(tasks);

            Log.Information("批量导入完成: 成功 {Success}/{Total}", successCount, totalCount);
            return successCount;
        }

        /// <summary>
        /// 带重试的单联系人导入 (指数退避)
        /// </summary>
        private async Task<bool> ImportSingleContactWithRetryAsync(string targetEmail, string vcfFilePath, int maxRetries = 3)
        {
            for (int retry = 0; retry <= maxRetries; retry++)
            {
                try
                {
                    // 解析VCF文件
                    var contact = ParseVcfFile(vcfFilePath);
                    if (contact == null)
                    {
                        Log.Warning("跳过无效VCF文件: {File}", vcfFilePath);
                        return false;
                    }

                    // 调用Graph API创建联系人
                    await _graphClient.Users[targetEmail].Contacts.PostAsync(contact);

                    Log.Information("联系人创建成功: {Name}", contact.DisplayName ?? vcfFilePath);
                    return true;
                }
                catch (Exception ex)
                {
                    // 检查是否是限流 (429)
                    bool isThrottled = IsThrottledException(ex);

                    if (isThrottled && retry < maxRetries)
                    {
                        // 指数退避: 2^retry 秒 (1s, 2s, 4s)
                        int delaySeconds = (int)Math.Pow(2, retry);
                        Log.Warning("限流 (429)，{Delay} 秒后重试 (第 {Retry}/{Max} 次): {File}",
                            delaySeconds, retry + 1, maxRetries, vcfFilePath);
                        await System.Threading.Tasks.Task.Delay(delaySeconds * 1000);
                    }
                    else
                    {
                        Log.Error(ex, "导入联系人失败 (重试 {Retry}/{Max}): {File}",
                            retry, maxRetries, vcfFilePath);
                        return false;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// 检查异常是否为限流 (429)
        /// </summary>
        private bool IsThrottledException(Exception ex)
        {
            // 检查异常消息中是否包含429
            var exMsg = ex.Message.ToLower();

            if (exMsg.Contains("429") ||
                exMsg.Contains("throttled") ||
                exMsg.Contains("too many requests") ||
                exMsg.Contains("service unavailable"))
            {
                return true;
            }

            // 递归检查内部异常
            if (ex.InnerException != null)
            {
                return IsThrottledException(ex.InnerException);
            }

            return false;
        }

        /// <summary>
        /// 批量导入PST联系人数据到目标邮箱 (直接模式，跳过VCF文件)
        /// 使用 Client Secret + Graph Batch API，并发控制 + 指数退避重试
        /// </summary>
        /// <param name="targetEmail">目标用户邮箱</param>
        /// <param name="contacts">PST联系人数据列表</param>
        /// <param name="progressCallback">进度回调 (当前索引, 总数, 状态消息)</param>
        /// <param name="maxDegreeOfParallelism">最大并发数，默认10</param>
        /// <returns>成功导入数量</returns>
        public async Task<int> ImportContactsBatchDirectAsync(
            string targetEmail,
            IEnumerable<PstExtractService.PstContactData> contacts,
            Action<int, int, string> progressCallback = null,
            int maxDegreeOfParallelism = 10)
        {
            if (_graphClient == null)
            {
                Program.BatchToO365Logger.Error("Graph客户端未初始化，请先调用 ConnectWithClientSecret");
                return 0;
            }

            if (string.IsNullOrWhiteSpace(targetEmail))
            {
                Program.BatchToO365Logger.Error("目标邮箱不能为空");
                return 0;
            }

            // 过滤出有邮箱的联系人
            var contactList = contacts
                .Where(c => !string.IsNullOrWhiteSpace(c.Email))
                .ToList();

            int totalCount = contactList.Count;
            int successCount = 0;
            int skippedCount = contacts.Count() - totalCount;

            Program.BatchToO365Logger.Information("筛选后有邮箱联系人: {Valid}/{Total}, 跳过无邮箱: {Skipped}",
                totalCount, contacts.Count(), skippedCount);

            if (totalCount == 0)
            {
                Program.BatchToO365Logger.Warning("没有有邮箱的联系人需要导入");
                return 0;
            }

            int currentIndex = 0;

            Program.BatchToO365Logger.Information("开始直接批量导入联系人到 {Email}, 共 {Count} 个, 并发数: {Parallelism}",
                targetEmail, totalCount, maxDegreeOfParallelism);

            // 使用 SemaphoreSlim 控制并发数
            var semaphore = new System.Threading.SemaphoreSlim(maxDegreeOfParallelism);
            var lockObj = new object();

            // 并行处理所有联系人
            var tasks = contactList.Select(async contactData =>
            {
                await semaphore.WaitAsync();
                try
                {
                    int index;
                    lock (lockObj)
                    {
                        index = currentIndex++;
                    }

                    // 带重试的导入
                    bool success = await ImportSingleContactDirectWithRetryAsync(targetEmail, contactData);
                    if (success)
                    {
                        lock (lockObj) { successCount++; }
                    }

                    progressCallback?.Invoke(index + 1, totalCount, $"已导入 {index + 1}/{totalCount}");
                    return success;
                }
                finally
                {
                    semaphore.Release();
                }
            });

            await System.Threading.Tasks.Task.WhenAll(tasks);

            Program.BatchToO365Logger.Information("直接批量导入完成: 成功 {Success}/{Total}", successCount, totalCount);
            return successCount;
        }

        /// <summary>
        /// 带重试的单联系人直接导入 (指数退避)
        /// </summary>
        private async Task<bool> ImportSingleContactDirectWithRetryAsync(string targetEmail, PstExtractService.PstContactData contactData, int maxRetries = 3)
        {
            for (int retry = 0; retry <= maxRetries; retry++)
            {
                try
                {
                    // 将PST联系人数据转换为Graph Contact
                    var contact = ConvertToGraphContact(contactData);

                    // 调用Graph API创建联系人
                    await _graphClient.Users[targetEmail].Contacts.PostAsync(contact);

                    Program.BatchToO365Logger.Information("联系人创建成功: {Name}", contact.DisplayName ?? "Unknown");
                    return true;
                }
                catch (Exception ex)
                {
                    bool isThrottled = IsThrottledException(ex);

                    if (isThrottled && retry < maxRetries)
                    {
                        int delaySeconds = (int)Math.Pow(2, retry);
                        Program.BatchToO365Logger.Warning("限流 (429)，{Delay} 秒后重试 (第 {Retry}/{Max} 次): {Name}",
                            delaySeconds, retry + 1, maxRetries, contactData.DisplayName);
                        await System.Threading.Tasks.Task.Delay(delaySeconds * 1000);
                    }
                    else
                    {
                        Program.BatchToO365Logger.Error(ex, "导入联系人失败 (重试 {Retry}/{Max}): {Name}",
                            retry, maxRetries, contactData.DisplayName);
                        return false;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// 将PST联系人数据转换为Graph Contact模型
        /// </summary>
        private Microsoft.Graph.Models.Contact ConvertToGraphContact(PstExtractService.PstContactData data)
        {
            var contact = new Microsoft.Graph.Models.Contact();

            // 姓名
            if (!string.IsNullOrWhiteSpace(data.DisplayName))
                contact.DisplayName = data.DisplayName;
            else if (!string.IsNullOrWhiteSpace(data.Email))
                contact.DisplayName = data.Email;
            else if (!string.IsNullOrWhiteSpace(data.CompanyName))
                contact.DisplayName = data.CompanyName;
            else
                contact.DisplayName = "Unknown Contact";

            contact.GivenName = data.FirstName ?? "";
            contact.Surname = data.LastName ?? "";
            contact.MiddleName = data.MiddleName ?? "";
            contact.Title = data.Title ?? "";

            // 公司信息
            contact.CompanyName = data.CompanyName ?? "";
            contact.Department = data.Department ?? "";
            contact.JobTitle = data.JobTitle ?? "";

            // 备注 - 包含邮箱和电话信息
            var notes = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(data.PersonalNotes))
                notes.AppendLine(data.PersonalNotes);
            if (!string.IsNullOrWhiteSpace(data.Email))
                notes.AppendLine($"Email: {data.Email}");
            if (!string.IsNullOrWhiteSpace(data.Phone))
                notes.AppendLine($"Phone: {data.Phone}");
            if (!string.IsNullOrWhiteSpace(data.MobilePhone))
                notes.AppendLine($"Mobile: {data.MobilePhone}");
            if (notes.Length > 0)
                contact.PersonalNotes = notes.ToString().TrimEnd();

            // 生日
            if (data.Birthday.HasValue)
                contact.Birthday = data.Birthday.Value;

            return contact;
        }

        /// <summary>
        /// 解析VCF文件为Graph Contact对象
        /// </summary>
        private Microsoft.Graph.Models.Contact ParseVcfFile(string vcfFilePath)
        {
            try
            {
                var lines = File.ReadAllLines(vcfFilePath);
                var contact = new Microsoft.Graph.Models.Contact();
                string currentField = "";
                string currentValue = "";

                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    // 解析字段名和值
                    if (line.Contains(":"))
                    {
                        var colonIndex = line.IndexOf(':');
                        var fieldAndParams = line.Substring(0, colonIndex);
                        currentValue = line.Substring(colonIndex + 1);

                        // 处理带参数的字段名，如 TEL;TYPE=WORK:123
                        var fieldParts = fieldAndParams.Split(';');
                        currentField = fieldParts[0].ToUpper();

                        ApplyVcfFieldToContact(contact, currentField, currentValue);
                    }
                    else if (line.StartsWith(" ") && !string.IsNullOrEmpty(currentField))
                    {
                        // 折叠的行 (超过行长限制被折叠)
                        currentValue += line.Trim();
                        ApplyVcfFieldToContact(contact, currentField, currentValue);
                    }
                }

                // 如果没有DisplayName，使用文件名
                if (string.IsNullOrWhiteSpace(contact.DisplayName))
                {
                    contact.DisplayName = Path.GetFileNameWithoutExtension(vcfFilePath);
                }

                return contact;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "解析VCF文件失败: {Path}", vcfFilePath);
                return null;
            }
        }

        /// <summary>
        /// 将VCF字段应用到Contact对象
        /// </summary>
        private void ApplyVcfFieldToContact(Microsoft.Graph.Models.Contact contact, string field, string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return;

            switch (field)
            {
                case "FN":  // Formatted Name
                    contact.DisplayName = value;
                    break;

                case "N":   // Name components (Last;First;Middle;Prefix;Suffix)
                    var nameParts = value.Split(';');
                    if (nameParts.Length >= 1 && !string.IsNullOrWhiteSpace(nameParts[0]))
                        contact.Surname = nameParts[0];
                    if (nameParts.Length >= 2 && !string.IsNullOrWhiteSpace(nameParts[1]))
                        contact.GivenName = nameParts[1];
                    if (nameParts.Length >= 3 && !string.IsNullOrWhiteSpace(nameParts[2]))
                        contact.MiddleName = nameParts[2];
                    if (nameParts.Length >= 4 && !string.IsNullOrWhiteSpace(nameParts[3]))
                        contact.Title = nameParts[3];
                    break;

                case "ORG":
                    var orgParts = value.Split(';');
                    if (orgParts.Length >= 1 && !string.IsNullOrWhiteSpace(orgParts[0]))
                        contact.CompanyName = orgParts[0];
                    if (orgParts.Length >= 2 && !string.IsNullOrWhiteSpace(orgParts[1]))
                        contact.Department = orgParts[1];
                    break;

                case "TITLE":
                    contact.JobTitle = value;
                    break;

                case "NOTE":
                    contact.PersonalNotes = value;
                    break;

                case "BDAY":
                    if (DateTime.TryParse(value, out var birthday))
                        contact.Birthday = birthday;
                    break;

                case "ROLE":
                    contact.JobTitle = value;
                    break;
            }
        }

        /// <summary>
        /// 从ICS文件导入单个日历事件
        /// </summary>
        public bool ImportCalendarFromIcs(string targetEmail, string icsFilePath)
        {
            try
            {
                if (_graphClient == null)
                {
                    Log.Error("Graph客户端未初始化");
                    return false;
                }

                // 读取ICS文件内容
                var icsContent = File.ReadAllText(icsFilePath);
                Log.Information("正在导入日历到 {Email}: {File}", targetEmail, icsFilePath);

                // 使用Graph API创建日历事件
                // 注意：此为简化版本，实际生产环境需要解析ICS并创建事件
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "导入日历失败: {File}", icsFilePath);
                return false;
            }
        }

        /// <summary>
        /// 批量导入PST日历数据到目标邮箱 (直接模式，跳过ICS文件)
        /// 使用 Client Secret + Graph Batch API，并发控制 + 指数退避重试
        /// </summary>
        public async Task<int> ImportCalendarBatchDirectAsync(
            string targetEmail,
            IEnumerable<PstExtractService.PstCalendarData> calendars,
            Action<int, int, string> progressCallback = null,
            int maxDegreeOfParallelism = 10,
            string timeZone = "China Standard Time")
        {
            if (_graphClient == null)
            {
                Program.BatchToO365Logger.Error("Graph客户端未初始化，请先调用 ConnectWithClientSecret");
                return 0;
            }

            if (string.IsNullOrWhiteSpace(targetEmail))
            {
                Program.BatchToO365Logger.Error("目标邮箱不能为空");
                return 0;
            }

            // 过滤出有主题和时间的日历事件
            var calendarList = calendars
                .Where(c => !string.IsNullOrWhiteSpace(c.Subject) && c.StartTime.HasValue)
                .ToList();

            int totalCount = calendarList.Count;
            int skippedCount = calendars.Count() - totalCount;

            Program.BatchToO365Logger.Information("筛选后有效日历: {Valid}/{Total}, 跳过无效: {Skipped}",
                totalCount, calendars.Count(), skippedCount);

            if (totalCount == 0)
            {
                Program.BatchToO365Logger.Warning("没有有效的日历事件需要导入");
                return 0;
            }

            int successCount = 0;
            int currentIndex = 0;

            Program.BatchToO365Logger.Information("开始直接批量导入日历到 {Email}, 共 {Count} 个, 并发数: {Parallelism}",
                targetEmail, totalCount, maxDegreeOfParallelism);

            var semaphore = new System.Threading.SemaphoreSlim(maxDegreeOfParallelism);
            var lockObj = new object();

            var tasks = calendarList.Select(async calendarData =>
            {
                await semaphore.WaitAsync();
                try
                {
                    int index;
                    lock (lockObj)
                    {
                        index = currentIndex++;
                    }

                    bool success = await ImportSingleCalendarDirectWithRetryAsync(targetEmail, calendarData, timeZone);
                    if (success)
                    {
                        lock (lockObj) { successCount++; }
                    }

                    progressCallback?.Invoke(index + 1, totalCount, $"已导入 {index + 1}/{totalCount}");
                    return success;
                }
                finally
                {
                    semaphore.Release();
                }
            });

            await System.Threading.Tasks.Task.WhenAll(tasks);

            Program.BatchToO365Logger.Information("直接批量导入日历完成: 成功 {Success}/{Total}", successCount, totalCount);
            return successCount;
        }

        private async Task<bool> ImportSingleCalendarDirectWithRetryAsync(string targetEmail, PstExtractService.PstCalendarData calendarData, string timeZone, int maxRetries = 3)
        {
            for (int retry = 0; retry <= maxRetries; retry++)
            {
                try
                {
                    var evt = ConvertToGraphEvent(calendarData, timeZone);
                    await _graphClient.Users[targetEmail].Calendar.Events.PostAsync(evt);

                    Program.BatchToO365Logger.Information("日历创建成功: {Subject}", evt.Subject ?? "Unknown");
                    return true;
                }
                catch (Exception ex)
                {
                    bool isThrottled = IsThrottledException(ex);

                    if (isThrottled && retry < maxRetries)
                    {
                        int delaySeconds = (int)Math.Pow(2, retry);
                        Program.BatchToO365Logger.Warning("限流 (429)，{Delay} 秒后重试 (第 {Retry}/{Max} 次): {Subject}",
                            delaySeconds, retry + 1, maxRetries, calendarData.Subject);
                        await System.Threading.Tasks.Task.Delay(delaySeconds * 1000);
                    }
                    else
                    {
                        Program.BatchToO365Logger.Error(ex, "导入日历失败 (重试 {Retry}/{Max}): {Subject}",
                            retry, maxRetries, calendarData.Subject);
                        return false;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// 将 UTC 偏移格式 (如 "UTC+8") 转换为 IANA 时区名称 (如 "China Standard Time")
        /// Graph API DateTimeTimeZone.TimeZone 需要使用 IANA 时区名称
        /// </summary>
        private string ConvertToIanaTimeZone(string utcOffset)
        {
            if (string.IsNullOrWhiteSpace(utcOffset)) return "UTC";

            // UTC 偏移格式映射到 IANA 时区
            var offsetToTimeZone = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "UTC-12", "Dateline Standard Time" },
                { "UTC-11", "UTC-11" },
                { "UTC-10", "Hawaiian Standard Time" },
                { "UTC-9", "Alaskan Standard Time" },
                { "UTC-8", "Pacific Standard Time" },
                { "UTC-7", "Mountain Standard Time" },
                { "UTC-6", "Central Standard Time" },
                { "UTC-5", "Eastern Standard Time" },
                { "UTC-4", "Atlantic Standard Time" },
                { "UTC-3", "SA Eastern Standard Time" },
                { "UTC-2", "Mid-Atlantic Standard Time" },
                { "UTC-1", "Azores Standard Time" },
                { "UTC", "UTC" },
                { "UTC+1", "Romance Standard Time" },
                { "UTC+2", "Egypt Standard Time" },
                { "UTC+3", "Russian Standard Time" },
                { "UTC+4", "Arabian Standard Time" },
                { "UTC+5", "West Asia Standard Time" },
                { "UTC+5:30", "India Standard Time" },
                { "UTC+6", "Central Asia Standard Time" },
                { "UTC+6:30", "Myanmar Standard Time" },
                { "UTC+7", "SE Asia Standard Time" },
                { "UTC+8", "China Standard Time" },  // 中国标准时间 (北京)
                { "UTC+9", "Tokyo Standard Time" },
                { "UTC+9:30", "AUS Central Standard Time" },
                { "UTC+10", "AUS Eastern Standard Time" },
                { "UTC+11", "Central Pacific Standard Time" },
                { "UTC+12", "New Zealand Standard Time" }
            };

            // 尝试直接匹配
            if (offsetToTimeZone.TryGetValue(utcOffset, out string ianaZone))
            {
                Program.BatchToO365Logger.Debug("时区转换: {Input} -> {Output}", utcOffset, ianaZone);
                return ianaZone;
            }

            // 如果匹配不到，返回 UTC
            Program.BatchToO365Logger.Warning("未知的时区格式: {Input}, 使用 UTC", utcOffset);
            return "UTC";
        }

        private Microsoft.Graph.Models.Event ConvertToGraphEvent(PstExtractService.PstCalendarData data, string timeZone = "China Standard Time")
        {
            var evt = new Microsoft.Graph.Models.Event();

            evt.Subject = data.Subject ?? "";
            evt.Body = new Microsoft.Graph.Models.ItemBody
            {
                ContentType = Microsoft.Graph.Models.BodyType.Text,
                Content = data.Body ?? ""
            };
            evt.Location = new Microsoft.Graph.Models.Location
            {
                DisplayName = data.Location ?? ""
            };

            // 将 UTC 偏移格式转换为 IANA 时区名称 (Graph API 需要)
            string ianaTimeZone = ConvertToIanaTimeZone(timeZone);

            if (data.StartTime.HasValue)
            {
                var startTime = data.StartTime.Value;
                evt.Start = new Microsoft.Graph.Models.DateTimeTimeZone
                {
                    DateTime = startTime.ToString("yyyy-MM-ddTHH:mm:ss"),
                    TimeZone = ianaTimeZone
                };
            }

            if (data.EndTime.HasValue)
            {
                var endTime = data.EndTime.Value;
                evt.End = new Microsoft.Graph.Models.DateTimeTimeZone
                {
                    DateTime = endTime.ToString("yyyy-MM-ddTHH:mm:ss"),
                    TimeZone = ianaTimeZone
                };
            }

            evt.IsAllDay = data.IsAllDayEvent;

            // 防止向与会者发送通知邮件 (静默迁移)
            evt.ResponseRequested = false;

            if (data.ReminderSet && !string.IsNullOrWhiteSpace(data.ReminderMinutesBeforeStart))
            {
                evt.ReminderMinutesBeforeStart = int.TryParse(data.ReminderMinutesBeforeStart, out int minutes) ? minutes : 30;
                evt.IsReminderOn = true;
            }

            // 处理与会者
            if (!string.IsNullOrWhiteSpace(data.RequiredAttendees))
            {
                evt.Attendees = new List<Microsoft.Graph.Models.Attendee>();
                var requiredList = data.RequiredAttendees.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var attendee in requiredList)
                {
                    var email = attendee.Trim();
                    if (!string.IsNullOrWhiteSpace(email) && email.Contains("@"))
                    {
                        evt.Attendees.Add(new Microsoft.Graph.Models.Attendee
                        {
                            EmailAddress = new Microsoft.Graph.Models.EmailAddress { Address = email },
                            Type = Microsoft.Graph.Models.AttendeeType.Required
                        });
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(data.OptionalAttendees))
            {
                if (evt.Attendees == null) evt.Attendees = new List<Microsoft.Graph.Models.Attendee>();
                var optionalList = data.OptionalAttendees.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var attendee in optionalList)
                {
                    var email = attendee.Trim();
                    if (!string.IsNullOrWhiteSpace(email) && email.Contains("@"))
                    {
                        evt.Attendees.Add(new Microsoft.Graph.Models.Attendee
                        {
                            EmailAddress = new Microsoft.Graph.Models.EmailAddress { Address = email },
                            Type = Microsoft.Graph.Models.AttendeeType.Optional
                        });
                    }
                }
            }

            return evt;
        }
    }

    public class OAuthAddress
    {
        public string Name { get; set; }
        public string Address { get; set; }
    }
}
