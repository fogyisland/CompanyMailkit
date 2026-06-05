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
using MailConverter.Services.Contacts;
using MailConverter.Services.Calendars;

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

        /// <summary>
        /// 日志回调 (用于将导入过程中的日志消息传递到UI的日志框)
        /// UI 可以设置此回调以接收 PstImportLogger 的同步输出
        /// </summary>
        public Action<string> LogCallback { get; set; }

        // 文件夹缓存：Key = "user@domain.com|FolderName", Value = FolderId
        private static ConcurrentDictionary<string, string> _folderCache = new ConcurrentDictionary<string, string>();
        private string _clientSecret;

        /// <summary>
        /// 输出日志: 写入 PstImportLogger 并触发 LogCallback
        /// </summary>
        private void EmitLog(string message)
        {
            PstImportLogger.Info(message);
            LogCallback?.Invoke(message);
        }

        /// <summary>
        /// 输出错误日志: 写入 PstImportLogger 并触发 LogCallback
        /// </summary>
        private void EmitLogError(string message, Exception ex = null)
        {
            PstImportLogger.Error(message, ex);
            LogCallback?.Invoke("[错误] " + message);
        }

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
                _graphAccessToken = accessToken;
                _isOAuth = true;
                _password = null;
                _domain = null;

                Log.Information("使用 OAuth2 连接 Office 365: {Email}", email);

                _service = new ExchangeService(ExchangeVersion.Exchange2016);
                _service.Url = new Uri($"https://outlook.office365.com/EWS/Exchange.asmx");

                // 使用 OAuth 凭据
                _service.Credentials = new OAuthCredentials(accessToken);

                // 验证 EWS 连接
                Folder.Bind(_service, WellKnownFolderName.Inbox);

                // 初始化 Graph 客户端 (使用同一 OAuth 访问令牌，aud=graph.microsoft.com)
                try
                {
                    var credential = new OAuthAccessTokenCredential(accessToken);
                    _graphClient = new GraphServiceClient(credential, new[] { "https://graph.microsoft.com/.default" });
                    Log.Information("Graph 客户端初始化成功 (OAuth)");
                }
                catch (Exception gEx)
                {
                    Log.Warning(gEx, "Graph 客户端初始化失败: {Msg}", gEx.Message);
                }

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
        /// 使用 Graph API 列出指定用户的邮件文件夹 (使用 access token 创建临时客户端)
        /// </summary>
        public List<string> ListGraphMailFolders(string accessToken, string userEmail)
        {
            var folderNames = new List<string> { "Inbox", "Sent Items", "Drafts", "Deleted Items", "Junk Email", "Archive" };
            try
            {
                if (string.IsNullOrEmpty(accessToken))
                    throw new ArgumentException("访问令牌为空");

                var credential = new OAuthAccessTokenCredential(accessToken);
                using (var client = new GraphServiceClient(credential, new[] { "https://graph.microsoft.com/.default" }))
                {
                    var result = client.Users[userEmail].MailFolders.GetAsync().Result;
                    if (result?.Value != null)
                    {
                        foreach (var folder in result.Value)
                        {
                            if (!string.IsNullOrEmpty(folder.DisplayName) && !folderNames.Contains(folder.DisplayName))
                                folderNames.Add(folder.DisplayName);
                        }
                    }
                }
                Log.Information("Graph API 文件夹列表获取成功: {Count} 个文件夹 (用户: {User})", folderNames.Count, userEmail);
                return folderNames;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Graph API 列文件夹失败: {Msg}", ex.Message);
                throw;
            }
        }

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
                EmitLog($"导入邮件: {emlPath} -> {targetUserEmail}/{folderName}");
                Log.Information("开始导入 EML: {Path} 到 {Email}, Folder: {Folder}", emlPath, targetUserEmail, folderName);

                // 首先尝试使用 Graph API MIME 上传（更可靠）
                Log.Information("尝试使用 Graph API MIME 上传...");
                if (ImportEmlWithGraphMime(emlPath, targetUserEmail, folderName))
                {
                    EmitLog($"  -> 成功! (Graph API MIME) 主题: {Path.GetFileName(emlPath)}, 文件夹: {folderName}");
                    return true;
                }
                Log.Warning("Graph API MIME 上传失败，尝试 EWS...");

                // 使用 MimeKit 解析 EML 文件
                var mimeMessage = MimeKit.MimeMessage.Load(emlPath);
                var subject = mimeMessage.Subject ?? "(无主题)";
                EmitLog($"  主题: {subject}, 发件人: {mimeMessage.From}, 附件: {mimeMessage.Attachments.Count()}");
                Log.Information("EML 解析完成, Subject={Subject}, From={From}, Attachments={Count}",
                    mimeMessage.Subject, mimeMessage.From, mimeMessage.Attachments.Count());

                // 确保 EWS 服务已连接
                if (_service == null)
                {
                    Log.Error("EWS 服务未连接");
                    EmitLogError("  -> 失败: EWS 服务未连接");
                    return false;
                }

                // 设置模拟用户 (Impersonation)
                // 重要：OAuth 委托认证下，如果目标邮箱是认证用户自己的邮箱，不需要 Impersonation
                // 只有应用程序权限才需要 Impersonation 访问其他用户的邮箱
                if (_isOAuth && targetUserEmail == _email)
                {
                    // 委托认证 + 自己的邮箱：不需要 Impersonation
                    _service.ImpersonatedUserId = null;
                    Log.Information("OAuth 委托认证自己的邮箱，跳过 Impersonation");
                }
                else
                {
                    // 应用权限或不同用户：需要 Impersonation
                    _service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, targetUserEmail);
                    Log.Information("已设置模拟用户: {Email}", targetUserEmail);
                }

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

                // 探测 EWS 写入权限: 尝试绑定目标文件夹 (Bind 失败时会抛出详细错误)
                try
                {
                    var probeFolder = Folder.Bind(_service, targetFolderId).Result;
                    EmitLog($"  -> EWS 目标文件夹验证: {probeFolder.DisplayName} (TotalCount={probeFolder.TotalCount})");
                }
                catch (Exception bindEx)
                {
                    EmitLogError($"  -> EWS 绑定文件夹失败 (可能缺少 EWS.AccessAsUser.All 写入权限): {bindEx.Message}");
                    // 继续尝试保存，让 EWS 抛出更具体的错误
                }

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

                        EmitLog($"  -> 成功! 主题: {subject}, 文件夹: {folderName}");
                        Log.Information("EML 导入成功: {Subject}, Folder: {Folder}, EmailId: {Id}", email.Subject, folderName, email.Id.ToString());
                        return true;
                    }
                    else
                    {
                        Log.Warning("EML 导入可能失败，邮件ID为空: {Subject}", email.Subject);
                        EmitLogError($"  -> 失败: email.Save() 返回但邮件ID为空 (主题: {email.Subject}, 文件夹: {folderName})");
                        return false;
                    }
                }
                catch (Exception saveEx)
                {
                    // 记录详细错误信息
                    var errorDetails = GetDetailedErrorMessage(saveEx);
                    Log.Error(saveEx, "保存邮件失败: {Subject}, Folder: {Folder}, Error: {Error}", email.Subject, folderName, errorDetails);
                    EmitLogError($"  -> EWS保存失败，尝试Graph API MIME上传: {errorDetails}");

                    // 尝试使用 Graph API 的 MIME 上传方式作为备用方案
                    try
                    {
                        Log.Information("尝试使用 Graph API MIME 上传: {Path}", emlPath);
                        if (ImportEmlWithGraphMime(emlPath, targetUserEmail, folderName))
                        {
                            EmitLog($"  -> Graph API MIME上传成功! 主题: {subject}, 文件夹: {folderName}");
                            return true;
                        }
                        else
                        {
                            EmitLogError("  -> Graph API MIME上传也失败");
                            return false;
                        }
                    }
                    catch (Exception mimeEx)
                    {
                        Log.Error(mimeEx, "Graph API MIME 上传也失败: {Path}", emlPath);
                        EmitLogError($"  -> Graph API MIME上传异常: {mimeEx.Message}");
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                EmitLogError($"  -> 异常: {ex.Message}", ex);
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

                // 支持嵌套路径: "MyImport/Mail1" -> 依次在 根/MyImport 下找 Mail1
                var segments = folderName.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);
                if (segments.Length == 0) segments = new[] { "Inbox" };

                string currentParentId = null;  // null = 邮箱根目录 (MailFolders 顶层)
                string currentPath = "";
                string finalId = null;

                for (int i = 0; i < segments.Length; i++)
                {
                    string segment = segments[i].Trim();
                    currentPath = string.IsNullOrEmpty(currentPath) ? segment : currentPath + "/" + segment;

                    // 1. 缓存优先 (整路径作为 key)
                    string cacheKey = $"{userEmail}|{currentPath}";
                    if (_folderCache.TryGetValue(cacheKey, out var cachedId))
                    {
                        finalId = cachedId;
                        currentParentId = cachedId;
                        continue;
                    }

                    // 2. 第一个段如果是已知文件夹 (inbox/sent/...), 走 well-known id
                    if (i == 0)
                    {
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
                        if (wellKnownFolders.TryGetValue(segment.ToLower(), out var wellKnownId))
                        {
                            _folderCache.TryAdd(cacheKey, wellKnownId);
                            finalId = wellKnownId;
                            currentParentId = wellKnownId;
                            continue;
                        }
                    }

                    // 3. 在当前父目录下查找/创建
                    string foundId = null;
                    var escapedName = EscapeODataString(segment);

                    if (currentParentId == null)
                    {
                        // 根目录: 使用 /mailFolders 顶层
                        var folder = _graphClient.Users[userEmail].MailFolders
                            .GetAsync(rc => rc.QueryParameters.Filter = $"displayName eq '{escapedName}'")
                            .GetAwaiter().GetResult();

                        if (folder?.Value?.Count > 0)
                        {
                            foundId = folder.Value[0].Id;
                        }
                        else
                        {
                            var newFolder = new Microsoft.Graph.Models.MailFolder { DisplayName = segment };
                            var created = _graphClient.Users[userEmail].MailFolders.PostAsync(newFolder).GetAwaiter().GetResult();
                            foundId = created?.Id;
                            Log.Information("为 {User} 在根目录创建文件夹: {Folder}", userEmail, segment);
                        }
                    }
                    else
                    {
                        // 子目录: 使用 /mailFolders/{parentId}/childFolders
                        var folder = _graphClient.Users[userEmail].MailFolders[currentParentId].ChildFolders
                            .GetAsync(rc => rc.QueryParameters.Filter = $"displayName eq '{escapedName}'")
                            .GetAwaiter().GetResult();

                        if (folder?.Value?.Count > 0)
                        {
                            foundId = folder.Value[0].Id;
                        }
                        else
                        {
                            var newFolder = new Microsoft.Graph.Models.MailFolder { DisplayName = segment };
                            var created = _graphClient.Users[userEmail].MailFolders[currentParentId].ChildFolders
                                .PostAsync(newFolder).GetAwaiter().GetResult();
                            foundId = created?.Id;
                            Log.Information("为 {User} 在 {Parent} 下创建子文件夹: {Folder}", userEmail, currentParentId, segment);
                        }
                    }

                    if (string.IsNullOrEmpty(foundId))
                    {
                        Log.Warning("无法解析文件夹段: {Segment}, 路径: {Path}, 回退到 inbox", segment, currentPath);
                        return "inbox";
                    }

                    _folderCache.TryAdd(cacheKey, foundId);
                    currentParentId = foundId;
                    finalId = foundId;
                }

                return finalId ?? "inbox";
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "获取文件夹ID失败，使用默认收件箱: {Folder}", folderName);
                return "inbox";
            }
        }

        /// <summary>
        /// 转义 OData filter 字符串中的单引号
        /// </summary>
        private static string EscapeODataString(string value)
        {
            return (value ?? "").Replace("'", "''");
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

            // 目标邮箱：使用 _email (登录时设置的)
            string targetEmail = _email ?? "";
            if (string.IsNullOrEmpty(targetEmail))
            {
                Log.Error("ImportEmlFolder: 目标邮箱为空 (_email 未设置)");
            }

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

                    // 优先使用 Graph API (OAuth 委托认证下 token aud=graph.microsoft.com，EWS 会 401)
                    if (ImportEmlWithGraph(file, targetEmail, subFolder))
                    {
                        imported++;
                    }
                    else if (ImportEml(file, subFolder))
                    {
                        // 回退到 EWS (用于 password 模式)
                        imported++;
                    }

                    if (progress != null && files.Length > 0)
                    {
                        int percentage = (int)(imported * 100L / files.Length);
                        progress.Report(percentage);
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
        /// 使用 Graph API 导入 EML (调用 Python 脚本 graph_deliver.py)
        /// Python 脚本负责: 解析 EML (容忍畸形头), 查找/创建文件夹, MIME 上传 + Message 对象 fallback
        /// </summary>
        private bool ImportEmlWithGraphMime(string emlPath, string targetUserEmail, string folderName = "Inbox")
        {
            try
            {
                if (string.IsNullOrEmpty(_accessToken))
                {
                    Log.Error("Graph 访问令牌未初始化");
                    EmitLogError("  -> Graph 访问令牌未初始化");
                    return false;
                }

                // 优先使用专门的 Graph Token (Scope: https://graph.microsoft.com/.default)
                // 回退到 EWS Token 仅作为兜底
                var tokenForGraph = !string.IsNullOrEmpty(_graphAccessToken) ? _graphAccessToken : _accessToken;

                // 找到 Python 可执行文件 (复用 Program.cs 的嵌入式 Python)
                var pythonExe = Program.GetPythonExecutable();
                if (string.IsNullOrEmpty(pythonExe) || !File.Exists(pythonExe))
                {
                    Log.Error("Python 环境未找到, 无法执行 Graph 投递");
                    EmitLogError("  -> Python 环境未找到, 无法执行 Graph 投递");
                    return ImportEmlWithEws(emlPath, targetUserEmail, folderName);
                }

                // 找到 graph_deliver.py 脚本
                var scriptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "script", "graph_deliver.py");
                if (!File.Exists(scriptPath))
                {
                    Log.Error("graph_deliver.py 脚本不存在: {Path}", scriptPath);
                    EmitLogError($"  -> graph_deliver.py 脚本不存在: {scriptPath}");
                    return ImportEmlWithEws(emlPath, targetUserEmail, folderName);
                }

                // 构造命令行参数:
                // graph_deliver.py <token> <user_email> <eml_path> [target_folder]
                // 总是传第4个参数, 空字符串代表"邮箱根目录"
                var args = $"\"{scriptPath}\" \"{tokenForGraph}\" \"{targetUserEmail}\" \"{emlPath}\" \"{folderName ?? ""}\"";

                Log.Information("调用 graph_deliver.py: 文件夹={Folder}, EML={Path}", folderName, emlPath);

                var startInfo = Program.CreatePythonStartInfo(pythonExe, args);
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardError = true;
                startInfo.UseShellExecute = false;
                startInfo.CreateNoWindow = true;

                using (var process = System.Diagnostics.Process.Start(startInfo))
                {
                    if (process == null)
                    {
                        Log.Error("启动 Python 进程失败");
                        EmitLogError("  -> 启动 Python 进程失败");
                        return ImportEmlWithEws(emlPath, targetUserEmail, folderName);
                    }

                    // 异步读取输出避免死锁
                    var stdoutTask = process.StandardOutput.ReadToEndAsync();
                    var stderrTask = process.StandardError.ReadToEndAsync();

                    if (!process.WaitForExit(60000)) // 60秒超时
                    {
                        try { process.Kill(); } catch { }
                        Log.Warning("graph_deliver.py 执行超时 (60s), 终止进程");
                        EmitLogError("  -> graph_deliver.py 执行超时");
                        return ImportEmlWithEws(emlPath, targetUserEmail, folderName);
                    }

                    var stdout = stdoutTask.GetAwaiter().GetResult();
                    var stderr = stderrTask.GetAwaiter().GetResult();

                    // 把 Python 的详细日志透传到 UI
                    if (!string.IsNullOrEmpty(stdout))
                    {
                        foreach (var line in stdout.Split('\n'))
                        {
                            var trimmed = line.TrimEnd('\r');
                            if (string.IsNullOrEmpty(trimmed)) continue;
                            if (trimmed.StartsWith("RESULT:"))
                            {
                                // 结果行单独记录
                                Log.Information("[graph_deliver] {Line}", trimmed);
                            }
                            else
                            {
                                EmitLog(trimmed);
                                Log.Information("[graph_deliver] {Line}", trimmed);
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(stderr))
                    {
                        foreach (var line in stderr.Split('\n'))
                        {
                            var trimmed = line.TrimEnd('\r');
                            if (string.IsNullOrEmpty(trimmed)) continue;
                            Log.Warning("[graph_deliver stderr] {Line}", trimmed);
                        }
                    }

                    if (process.ExitCode == 0)
                    {
                        Log.Information("Graph 投递成功: {Path}", emlPath);
                        return true;
                    }

                    Log.Warning("Graph 投递失败, ExitCode={Code}, Path: {Path}", process.ExitCode, emlPath);
                    EmitLogError($"  -> Graph 投递失败 (ExitCode={process.ExitCode})");
                    // Python 失败时尝试 EWS
                    return ImportEmlWithEws(emlPath, targetUserEmail, folderName);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "调用 graph_deliver.py 异常: {Path}, 尝试 EWS 方式", emlPath);
                EmitLogError($"  -> 调用 graph_deliver.py 异常: {ex.Message}");
                return ImportEmlWithEws(emlPath, targetUserEmail, folderName);
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
                EmitLogError($"  -> EWS 导入异常: {ex.Message}");
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
        /// 使用 Graph API (OAuth token 的 aud=graph.microsoft.com，EWS 在此 token 下会"成功"
        /// 但联系人实际未保存到 O365，必须走 Graph 路径)
        /// </summary>
        public bool CreateContact(string displayName, string emailAddress, string phone, string company, string title)
        {
            if (_graphClient == null)
            {
                Log.Error("Graph 客户端未初始化，请先调用 ConnectWithOAuth");
                return false;
            }
            if (string.IsNullOrWhiteSpace(_email))
            {
                Log.Error("未设置目标邮箱 (_email 为空)");
                return false;
            }

            try
            {
                var contact = new Microsoft.Graph.Models.Contact
                {
                    DisplayName = !string.IsNullOrWhiteSpace(displayName) ? displayName : emailAddress
                };

                // 邮箱
                if (!string.IsNullOrWhiteSpace(emailAddress))
                {
                    contact.EmailAddresses = new List<Microsoft.Graph.Models.EmailAddress>
                    {
                        new Microsoft.Graph.Models.EmailAddress
                        {
                            Name = !string.IsNullOrWhiteSpace(displayName) ? displayName : emailAddress,
                            Address = emailAddress
                        }
                    };
                }

                // 公司/职位
                if (!string.IsNullOrWhiteSpace(company))
                    contact.CompanyName = company;
                if (!string.IsNullOrWhiteSpace(title))
                    contact.JobTitle = title;

                // 电话 - 放入 BusinessPhones 列表
                if (!string.IsNullOrWhiteSpace(phone))
                {
                    contact.BusinessPhones = new List<string> { phone };
                }

                // 同步调用 Graph API
                var created = _graphClient.Users[_email].Contacts.PostAsync(contact).GetAwaiter().GetResult();

                if (created != null && !string.IsNullOrEmpty(created.Id))
                {
                    Log.Information("创建联系人成功 (Graph): {Name} <{Email}>, Id={Id}, Phone={Phone}, Company={Company}, Title={Title}",
                        displayName, emailAddress ?? "(无邮箱)", created.Id, phone, company, title);
                    return true;
                }
                else
                {
                    Log.Warning("Graph 创建联系人返回空 Id: {Name} <{Email}>", displayName, emailAddress ?? "(无邮箱)");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "创建联系人失败 (Graph): {Name} <{Email}>", displayName, emailAddress ?? "(无邮箱)");
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
            if (_graphClient == null && _service == null)
            {
                Log.Error("未连接 Office 365 (Graph 和 EWS 都未初始化)");
                return false;
            }

            // 优先走 Microsoft Graph(用户 OAuth 登录拿的 token 只声明了 Graph 范围的 scope,
            // 用它去 EWS 会被服务端 aud 校验静默拒绝,Save() 返回无异常但实际不写入日历)。
            // EWS 路径仅作为 Graph 不可用时的回退。
            if (_graphClient != null)
            {
                try
                {
                    var graphEvent = new Microsoft.Graph.Models.Event
                    {
                        Subject = subject,
                        Body = new Microsoft.Graph.Models.ItemBody
                        {
                            ContentType = Microsoft.Graph.Models.BodyType.Text,
                            Content = description ?? ""
                        },
                        Start = new Microsoft.Graph.Models.DateTimeTimeZone
                        {
                            // start 来自 CSV 解析,Kind=Unspecified,wall-clock 即用户本地时间
                            // 这里强制按"东八区"序列化,O365 会按用户邮箱时区显示一致
                            DateTime = start.ToString("yyyy-MM-ddTHH:mm:ss"),
                            TimeZone = "China Standard Time"
                        },
                        End = new Microsoft.Graph.Models.DateTimeTimeZone
                        {
                            DateTime = end.ToString("yyyy-MM-ddTHH:mm:ss"),
                            TimeZone = "China Standard Time"
                        },
                        Location = string.IsNullOrEmpty(location)
                            ? null
                            : new Microsoft.Graph.Models.Location { DisplayName = location }
                    };

                    var created = _graphClient.Me.Events.PostAsync(graphEvent).GetAwaiter().GetResult();
                    if (created != null && !string.IsNullOrEmpty(created.Id))
                    {
                        Log.Information("创建日历事件成功 (Graph): {Subject} id={Id}", subject, created.Id);
                        return true;
                    }
                    Log.Warning("创建日历事件 (Graph): {Subject} 返回为空,可能未创建", subject);
                    return false;
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "创建日历事件 (Graph) 失败: {Subject}", subject);
                    return false;
                }
            }

            // 回退:EWS 路径(仅在 Graph 不可用时)
            try
            {
                Appointment appointment = new Appointment(_service);
                appointment.Subject = subject;
                appointment.Start = start;
                appointment.End = end;
                appointment.Body = description;
                appointment.Location = location;
                appointment.Save(WellKnownFolderName.Calendar);
                Log.Information("创建日历事件成功 (EWS 回退): {Subject}", subject);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "创建日历事件 (EWS) 失败: {Subject}", subject);
                return false;
            }
        }

        /// <summary>
        /// 创建后立即按 Subject + Start 反查,确认事件真的写入了日历文件夹。
        /// EWS 偶尔会出现 Save() 不抛异常但服务端丢消息的情况,加这一层能尽早暴露问题。
        /// </summary>
        private bool VerifyAppointmentSaved(string subject, DateTime start)
        {
            try
            {
                var calendar = Folder.Bind(_service, WellKnownFolderName.Calendar).Result;
                Log.Debug("VerifyAppointmentSaved: 已绑定日历文件夹, TotalCount={Count}", calendar.TotalCount);

                var startWindow = start.AddMinutes(-2);
                var endWindow = start.AddMinutes(2);
                var filter = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                    new SearchFilter.IsEqualTo(AppointmentSchema.Subject, subject),
                    new SearchFilter.IsGreaterThanOrEqualTo(AppointmentSchema.Start, startWindow),
                    new SearchFilter.IsLessThanOrEqualTo(AppointmentSchema.Start, endWindow));
                var view = new ItemView(5);
                var result = calendar.FindItems(filter, view).Result;
                Log.Debug("VerifyAppointmentSaved: Subject='{Subject}' 查询命中 {Count} 条", subject, result?.Count() ?? 0);
                return result != null && result.Count() > 0;
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "VerifyAppointmentSaved 查询异常: Subject='{Subject}'", subject);
                return false; // 查询失败不阻塞主流程,按"未校验"对待
            }
        }

        /// <summary>
        /// 诊断:列出最近 N 个日历事件的 Subject + Start,用于排查"Save 成功但查不到"问题
        /// </summary>
        public List<(string Subject, DateTime Start)> ListRecentCalendarEvents(int count = 10)
        {
            var result = new List<(string, DateTime)>();
            try
            {
                var calendar = Folder.Bind(_service, WellKnownFolderName.Calendar).Result;
                var view = new ItemView(count);
                view.OrderBy.Add(AppointmentSchema.Start, SortDirection.Descending);
                var items = calendar.FindItems(view).Result;
                foreach (var item in items)
                {
                    if (item is Appointment appt)
                    {
                        result.Add((appt.Subject ?? "(无主题)", appt.Start));
                    }
                }
                Log.Information("ListRecentCalendarEvents: 日历文件夹共 {Total} 项, 最近 {Count} 条:", calendar.TotalCount, result.Count);
                foreach (var (s, st) in result)
                {
                    Log.Information("  - {Start:yyyy-MM-dd HH:mm} | {Subject}", st, s);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "ListRecentCalendarEvents 失败");
            }
            return result;
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

        // ========== 联系人 facade (转发到 ContactSyncService) ==========

        /// <summary>
        /// 批量导入PST联系人数据到目标邮箱 (直接模式，跳过VCF文件)
        /// </summary>
        public async Task<int> ImportContactsBatchDirectAsync(
            string targetEmail,
            IEnumerable<ContactData> contacts,
            Action<int, int, string> progressCallback = null,
            int maxDegreeOfParallelism = 10)
        {
            return await ContactSyncService.ImportContactsBatchDirectAsync(
                _graphClient, targetEmail, contacts, progressCallback, maxDegreeOfParallelism);
        }

        // ========== 日历 facade (转发到 CalendarSyncService) ==========

        /// <summary>
        /// 批量导入PST日历数据到目标邮箱 (直接模式，跳过ICS文件)
        /// </summary>
        public async Task<int> ImportCalendarBatchDirectAsync(
            string targetEmail,
            IEnumerable<CalendarData> calendars,
            Action<int, int, string> progressCallback = null,
            int maxDegreeOfParallelism = 10,
            string timeZone = "China Standard Time")
        {
            return await CalendarSyncService.ImportCalendarBatchDirectAsync(
                _graphClient, targetEmail, calendars, progressCallback, maxDegreeOfParallelism, timeZone);
        }
    }

    public class OAuthAddress
    {
        public string Name { get; set; }
        public string Address { get; set; }
    }

    /// <summary>
    /// 使用现有 OAuth 访问令牌的 Graph 凭据 (基于 Azure.Core.TokenCredential)
    /// (避免使用 InteractiveBrowserCredential / UsernamePasswordCredential 触发额外的用户交互)
    /// </summary>
    public class OAuthAccessTokenCredential : Azure.Core.TokenCredential
    {
        private readonly string _accessToken;
        private readonly DateTimeOffset _expiresOn;

        public OAuthAccessTokenCredential(string accessToken, DateTimeOffset? expiresOn = null)
        {
            _accessToken = accessToken;
            _expiresOn = expiresOn ?? DateTimeOffset.UtcNow.AddHours(1);
        }

        public override Azure.Core.AccessToken GetToken(Azure.Core.TokenRequestContext requestContext, System.Threading.CancellationToken cancellationToken = default)
        {
            return new Azure.Core.AccessToken(_accessToken, _expiresOn);
        }

        public override System.Threading.Tasks.ValueTask<Azure.Core.AccessToken> GetTokenAsync(Azure.Core.TokenRequestContext requestContext, System.Threading.CancellationToken cancellationToken = default)
        {
            return new System.Threading.Tasks.ValueTask<Azure.Core.AccessToken>(new Azure.Core.AccessToken(_accessToken, _expiresOn));
        }
    }
}
