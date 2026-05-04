using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Serilog;

namespace MailConverter
{
    public class CardDavService
    {
        private HttpClient _httpClient;
        private string _baseUrl;
        private string _username;
        private string _password;
        private List<CardDavContact> _googleContacts;

        public CardDavService()
        {
            _httpClient = new HttpClient();
            // 支持TLS 1.2
            ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
        }

        public bool Connect(string serverUrl, string username, string password)
        {
            try
            {
                // 保留URL末尾斜杠，不使用TrimEnd
                _baseUrl = serverUrl.Trim();
                _username = username;
                _password = password;

                Log.Information("CardDAV开始连接: {Url}, User: {User}", _baseUrl, username);

                // 强制TLS 1.2
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                ServicePointManager.DefaultConnectionLimit = 10;
                ServicePointManager.Expect100Continue = false;

                // 设置Basic认证
                var authValue = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{username}:{password}"));

                // 使用自定义Handler配置TLS
                var handler = new HttpClientHandler();
                handler.ServerCertificateCustomValidationCallback = delegate { return true; };
                handler.SslProtocols = System.Security.Authentication.SslProtocols.Tls12;
                // 禁用自动重定向，手动处理以保留Header
                handler.AllowAutoRedirect = false;
                handler.UseCookies = false;

                _httpClient = new HttpClient(handler);
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", authValue);

                // 测试连接 - 发送PROPFIND请求
                var testResult = TestConnection().Result;
                if (testResult)
                {
                    // 如果是Google Gmail，尝试进行资源发现
                    if (_baseUrl.Contains("googleapis.com/carddav"))
                    {
                        var discoveredUrl = DiscoverGoogleAddressbook();
                        if (!string.IsNullOrEmpty(discoveredUrl))
                        {
                            _baseUrl = discoveredUrl;
                            Log.Information("CardDAV Google地址簿发现成功: {Url}", _baseUrl);
                        }
                    }
                    Log.Information("CardDAV连接成功: {Url}", _baseUrl);
                    return true;
                }
                Log.Warning("CardDAV连接失败: {Url}", _baseUrl);
                return false;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "CardDAV连接异常: {Url}", serverUrl);
                return false;
            }
        }

        private async Task<bool> TestConnection()
        {
            try
            {
                Log.Information("CardDAV测试连接: {Url}", _baseUrl);

                // 使用 TryAddWithoutValidation 添加 User-Agent
                _httpClient.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "MyCsharpApp/1.0");

                // PROPFIND 请求需要带 XML Body (Google要求查找addressbook-home-set)
                string propfindXml = @"<?xml version='1.0' encoding='utf-8' ?>
<d:propfind xmlns:d='DAV:' xmlns:c='urn:ietf:params:xml:ns:carddav'>
  <d:prop>
    <c:addressbook-home-set />
  </d:prop>
</d:propfind>";

                // 直接发送 PROPFIND 请求，Depth 为 0
                var request = new HttpRequestMessage(new HttpMethod("PROPFIND"), _baseUrl);
                request.Content = new StringContent(propfindXml, Encoding.UTF8, "text/xml");
                request.Headers.Add("Depth", "0");

                var response = await _httpClient.SendAsync(request);
                var statusCode = (int)response.StatusCode;
                Log.Information("CardDAV PROPFIND响应: {Status} ({Code})", response.StatusCode, statusCode);

                // 207=MultiStatus成功, 404=NotFound(URL错误), 401=认证失败, 403=禁止访问, 400=请求错误
                if (statusCode == 207 || statusCode == 404 || statusCode == 401 || statusCode == 403 || statusCode == 400)
                {
                    Log.Information("CardDAV连接验证通过（服务器有响应）: {Status}", response.StatusCode);
                    return true;
                }

                return response.IsSuccessStatusCode;
            }
            catch (Exception ex)
            {
                // 提取详细错误信息
                string detailedMessage = ex.Message;
                if (ex.InnerException != null)
                {
                    detailedMessage += " | 内部异常: " + ex.InnerException.Message;
                    if (ex.InnerException.InnerException != null)
                    {
                        detailedMessage += " | 根因: " + ex.InnerException.InnerException.Message;
                    }
                }
                Log.Error(ex, "CardDAV连接测试异常: {Url}, 详情: {Detail}", _baseUrl, detailedMessage);
                return false;
            }
        }

        // Google CardDAV 资源发现
        private string DiscoverGoogleAddressbook()
        {
            try
            {
                Log.Information("开始Google CardDAV资源发现...");

                // 发送PROPFIND到principals获取addressbook-home-set
                string propfindXml = @"<?xml version='1.0' encoding='utf-8' ?>
<d:propfind xmlns:d='DAV:' xmlns:c='urn:ietf:params:xml:ns:carddav'>
  <d:prop>
    <c:addressbook-home-set />
  </d:prop>
</d:propfind>";

                var request = new HttpRequestMessage(new HttpMethod("PROPFIND"), _baseUrl);
                request.Content = new StringContent(propfindXml, Encoding.UTF8, "text/xml");
                request.Headers.TryAddWithoutValidation("User-Agent", "MyCsharpApp/1.0");
                request.Headers.Add("Depth", "0");

                var response = _httpClient.SendAsync(request).Result;
                var statusCode = (int)response.StatusCode;
                Log.Information("Google资源发现响应: {Status} ({Code})", response.StatusCode, statusCode);

                if (statusCode == 207)
                {
                    string xmlResult = response.Content.ReadAsStringAsync().Result;
                    Log.Information("资源发现响应内容: {Content}", xmlResult.Substring(0, Math.Min(800, xmlResult.Length)));

                    // 解析XML获取addressbook-home-set URL
                    // 通常格式: <addressbook-home-set><d:href>/carddav/v1/principals/xxx/lists/</d:href></addressbook-home-set>
                    if (xmlResult.Contains("addressbook-home-set") || xmlResult.Contains("addressbook"))
                    {
                        // 尝试提取URL
                        var startTag = xmlResult.IndexOf("<d:href>");
                        if (startTag >= 0)
                        {
                            var endTag = xmlResult.IndexOf("</d:href>", startTag);
                            if (endTag > startTag)
                            {
                                var href = xmlResult.Substring(startTag + 8, endTag - startTag - 8);
                                // 确保URL以/结尾
                                if (!href.EndsWith("/"))
                                    href += "/";
                                // 确保包含lists/路径
                                if (!href.Contains("/lists/"))
                                {
                                    // 如果href是 /carddav/v1/principals/xxx/ 需要加上lists/default
                                    href = href.TrimEnd('/') + "/lists/default/";
                                }
                                else
                                {
                                    // 如果href已经包含lists/，只需加上default
                                    href = href.TrimEnd('/') + "/default/";
                                }
                                // 完整URL
                                var addressbookUrl = "https://www.googleapis.com" + href;
                                Log.Information("发现地址簿URL: {Url}", addressbookUrl);

                                // 现在获取联系人列表
                                var contacts = GetGoogleContacts(addressbookUrl);
                                if (contacts != null && contacts.Count > 0)
                                {
                                    Log.Information("成功获取 {Count} 个联系人", contacts.Count);
                                    // 保存联系人到成员变量供后续使用
                                    _googleContacts = contacts;
                                }

                                return addressbookUrl;
                            }
                        }
                    }
                }

                Log.Warning("未能在响应中找到addressbook-home-set");
                return null;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Google CardDAV资源发现异常");
                return null;
            }
        }

        // 获取Google联系人列表
        private List<CardDavContact> GetGoogleContacts(string addressbookUrl)
        {
            try
            {
                Log.Information("开始获取Google联系人列表: {Url}", addressbookUrl);

                // 先只获取联系人列表（href和etag），不包含address-data
                string propfindXml = @"<?xml version='1.0' encoding='utf-8' ?>
<d:propfind xmlns:d='DAV:'>
  <d:prop>
    <d:getetag />
  </d:prop>
</d:propfind>";

                var request = new HttpRequestMessage(new HttpMethod("PROPFIND"), addressbookUrl);
                request.Content = new StringContent(propfindXml, Encoding.UTF8, "text/xml");
                request.Headers.TryAddWithoutValidation("User-Agent", "MyCsharpApp/1.0");
                request.Headers.Add("Depth", "1");

                var response = _httpClient.SendAsync(request).Result;
                var statusCode = (int)response.StatusCode;
                Log.Information("Google联系人列表响应: {Status} ({Code})", response.StatusCode, statusCode);

                var contacts = new List<CardDavContact>();

                if (statusCode == 207)
                {
                    string xmlResult = response.Content.ReadAsStringAsync().Result;
                    Log.Information("联系人列表响应长度: {Length}", xmlResult.Length);
                    // 打印响应前2000个字符用于调试
                    Log.Information("PROPFIND响应内容前2000字符: {Content}", xmlResult.Substring(0, Math.Min(2000, xmlResult.Length)));

                    // 使用XDocument解析XML，处理命名空间
                    try
                    {
                        var xdoc = XDocument.Parse(xmlResult);

                        // 显式定义命名空间
                        XNamespace d = "DAV:";
                        XNamespace card = "urn:ietf:params:xml:ns:carddav";

                        // 查找所有的 <d:response> 节点
                        var responses = xdoc.Descendants(d + "response");
                        Log.Information("发现 {Count} 个response节点", responses.Count());

                        // 调试：打印所有节点名
                        foreach (var element in xdoc.Descendants().Take(5))
                        {
                            Log.Information("调试节点: {Name}", element.Name);
                        }

                        int count = 0;
                        foreach (var resp in responses)
                        {
                            // 查找 <d:href>
                            var href = resp.Element(d + "href")?.Value;

                            // 查找 <card:address-data> 获取vCard内容
                            var addressData = resp.Descendants(card + "address-data").FirstOrDefault()?.Value;

                            // 过滤：排除文件夹本身（以/结尾的href）
                            if (!string.IsNullOrEmpty(href) && !href.EndsWith("/"))
                            {
                                count++;
                                var fullUrl = "https://www.googleapis.com" + href;
                                contacts.Add(new CardDavContact { Url = fullUrl, VCardData = addressData });
                                Log.Information("发现联系人: {Url}, 包含vCard: {HasData}", fullUrl, !string.IsNullOrEmpty(addressData));
                            }
                        }

                        Log.Information("使用XDocument解析到 {Count} 个联系人", count);
                    }
                    catch (Exception parseEx)
                    {
                        Log.Warning("XDocument解析失败，尝试字符串匹配: {Error}", parseEx.Message);
                        // 回退到字符串匹配
                        int startIndex = 0;
                        while (true)
                        {
                            var hrefStart = xmlResult.IndexOf("<d:href>", startIndex);
                            if (hrefStart < 0) break;

                            var hrefEnd = xmlResult.IndexOf("</d:href>", hrefStart);
                            if (hrefEnd < 0) break;

                            var href = xmlResult.Substring(hrefStart + 8, hrefEnd - hrefStart - 8);

                            // 只处理.vcf文件
                            if (href.EndsWith(".vcf"))
                            {
                                var fullUrl = "https://www.googleapis.com" + href;
                                contacts.Add(new CardDavContact { Url = fullUrl });
                                Log.Information("发现联系人: {Url}", fullUrl);
                            }

                            startIndex = hrefEnd + 8;
                        }
                    }

                    Log.Information("共发现 {Count} 个联系人", contacts.Count);
                }

                return contacts;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取Google联系人列表异常");
                return null;
            }
        }

        public async Task<List<CardDavContact>> GetContactsAsync()
        {
            // 如果已经在发现阶段获取了Google联系人，直接返回
            if (_googleContacts != null && _googleContacts.Count > 0)
            {
                Log.Information("使用已获取的Google联系人: {Count}", _googleContacts.Count);
                return _googleContacts;
            }

            var contacts = new List<CardDavContact>();

            try
            {
                // QQ邮箱CardDAV PROPFIND请求
                string propfindXml = @"<?xml version='1.0' encoding='utf-8' ?>
                    <d:propfind xmlns:d='DAV:' xmlns:c='urn:ietf:params:xml:ns:carddav'>
                        <d:prop>
                            <d:getetag />
                        </d:prop>
                    </d:propfind>";

                var request = new HttpRequestMessage(new HttpMethod("PROPFIND"), _baseUrl);
                request.Content = new StringContent(propfindXml, Encoding.UTF8, "text/xml");
                request.Headers.TryAddWithoutValidation("User-Agent", "MyCsharpApp/1.0");
                request.Headers.Add("Depth", "1");

                Log.Information("发送PROPFIND请求到: {Url}", _baseUrl);
                var response = await _httpClient.SendAsync(request);
                Log.Information("PROPFIND响应: {Status}", response.StatusCode);

                if (response.IsSuccessStatusCode)
                {
                    string xmlResult = await response.Content.ReadAsStringAsync();
                    Log.Information("响应内容长度: {Length}", xmlResult.Length);
                    contacts = ParseContactUrls(xmlResult);
                }
                else
                {
                    string errorContent = await response.Content.ReadAsStringAsync();
                    Log.Warning("PROPFIND请求失败: {Status}, 内容: {Error}", response.StatusCode, errorContent);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取联系人列表失败");
            }

            return contacts;
        }

        private List<CardDavContact> ParseContactUrls(string xml)
        {
            var contacts = new List<CardDavContact>();
            try
            {
                XNamespace d = "DAV:";
                var doc = XDocument.Parse(xml);
                var hrefs = doc.Descendants(d + "href");

                foreach (var href in hrefs)
                {
                    var url = href.Value;
                    // 过滤掉目录，只保留.vcf文件
                    if (url.EndsWith(".vcf") || url.EndsWith(".vcf/"))
                    {
                        var cleanUrl = url.TrimEnd('/');
                        contacts.Add(new CardDavContact { Url = cleanUrl });
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "解析CardDAV响应失败");
            }
            return contacts;
        }

        // 获取联系人列表（包含ETag用于增量同步）
        public async Task<List<CardDavContact>> GetContactsWithETag()
        {
            var contacts = new List<CardDavContact>();

            try
            {
                Log.Information("开始获取联系人列表(带ETag): {Url}", _baseUrl);

                // PROPFIND请求获取href和etag
                string propfindXml = @"<?xml version='1.0' encoding='utf-8' ?>
<d:propfind xmlns:d='DAV:'>
  <d:prop>
    <d:getetag />
  </d:prop>
</d:propfind>";

                var request = new HttpRequestMessage(new HttpMethod("PROPFIND"), _baseUrl);
                request.Content = new StringContent(propfindXml, Encoding.UTF8, "text/xml");
                request.Headers.TryAddWithoutValidation("User-Agent", "MyCsharpApp/1.0");
                request.Headers.Add("Depth", "1");

                var response = await _httpClient.SendAsync(request);
                var statusCode = (int)response.StatusCode;
                Log.Information("联系人列表响应: {Status} ({Code})", response.StatusCode, statusCode);

                if (statusCode == 207)
                {
                    string xmlResult = await response.Content.ReadAsStringAsync();
                    var xdoc = XDocument.Parse(xmlResult);
                    XNamespace d = "DAV:";

                    var responses = xdoc.Descendants(d + "response");
                    Log.Information("发现 {Count} 个response节点", responses.Count());

                    int count = 0;
                    foreach (var resp in responses)
                    {
                        var href = resp.Element(d + "href")?.Value;
                        var etag = resp.Descendants(d + "getetag").FirstOrDefault()?.Value;

                        // 排除文件夹本身
                        if (!string.IsNullOrEmpty(href) && !href.EndsWith("/"))
                        {
                            var fullUrl = "https://www.googleapis.com" + href;
                            var contact = new CardDavContact { Url = fullUrl, ETag = etag };

                            // 获取vCard内容以提取姓名和邮箱
                            try
                            {
                                var vcfContent = await GetVCardContentAsync(fullUrl);

                                // 检查是否返回 502 错误
                                if (vcfContent == "502")
                                {
                                    Log.Warning("遇到 Bad Gateway，跳过此联系人: {Url}", fullUrl);
                                    contacts.Add(contact);
                                    count++;
                                    continue;
                                }

                                // 打印原始vCard内容用于调试
                                Log.Information("vCard原始内容({Length}字节): {Content}", vcfContent?.Length ?? 0, vcfContent?.Substring(0, Math.Min(500, vcfContent?.Length ?? 0)));

                                // 显示所有有内容的联系人（长度>=100字节），即使没有邮箱
                                if (!string.IsNullOrEmpty(vcfContent) && vcfContent.Length >= 100)
                                {
                                    var parsed = ParseVCard(vcfContent);
                                    if (parsed.Count > 0)
                                    {
                                        contact.Name = parsed[0].Name;
                                        contact.Email = parsed[0].Email;
                                        contact.Phone = parsed[0].Phone;
                                        Log.Information("解析结果: Name={Name}, Email={Email}, Phone={Phone}", contact.Name, contact.Email, contact.Phone);
                                        // 如果没有邮箱，显示"无邮箱"
                                        if (string.IsNullOrWhiteSpace(contact.Email))
                                        {
                                            contact.Email = "(无邮箱)";
                                        }
                                    }
                                }
                                else
                                {
                                    contact.Name = "(空联系人)";
                                    contact.Email = "(无邮箱)";
                                }

                                // 添加小延迟避免请求过快
                                await Task.Delay(50);
                            }
                            catch { }

                            contacts.Add(contact);
                            count++;
                            if (count % 10 == 0)
                                Log.Information("已处理 {Count}/{Total} 个联系人", count, responses.Count());
                        }
                    }

                    Log.Information("获取到 {Count} 个联系人ETag", contacts.Count);
                }

                return contacts;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取联系人ETag失败");
                return null;
            }
        }

        public async Task<string> GetVCardContentAsync(string contactUrl)
        {
            try
            {
                var request = new HttpRequestMessage(HttpMethod.Get, contactUrl);
                request.Headers.TryAddWithoutValidation("User-Agent", "MyCsharpApp/1.0");
                var response = await _httpClient.SendAsync(request);

                // 处理 502 错误
                if ((int)response.StatusCode == 502)
                {
                    Log.Warning("获取vCard遇到 Bad Gateway: {Url}", contactUrl);
                    return "502";
                }

                if (response.IsSuccessStatusCode)
                {
                    var content = await response.Content.ReadAsStringAsync();
                    Log.Information("成功获取vCard: {Url}, 长度: {Length}", contactUrl, content.Length);
                    return content;
                }
                else
                {
                    Log.Warning("获取vCard失败: {Url}, 状态: {Status}", contactUrl, response.StatusCode);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取vCard失败: {Url}", contactUrl);
            }
            return null;
        }

        public List<CardDavContact> ParseVCard(string vcfContent)
        {
            var contacts = new List<CardDavContact>();
            var lines = vcfContent.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

            string currentName = "";
            string currentEmail = "";
            string currentPhone = "";
            bool inVCard = false;

            foreach (var line in lines)
            {
                if (line.StartsWith("BEGIN:VCARD", StringComparison.OrdinalIgnoreCase))
                {
                    inVCard = true;
                    currentName = "";
                    currentEmail = "";
                    currentPhone = "";
                }
                else if (line.StartsWith("END:VCARD", StringComparison.OrdinalIgnoreCase))
                {
                    // 只要有姓名或电话或邮箱，就添加到联系人列表
                    if (!string.IsNullOrWhiteSpace(currentName) || !string.IsNullOrWhiteSpace(currentPhone) || !string.IsNullOrWhiteSpace(currentEmail))
                    {
                        contacts.Add(new CardDavContact
                        {
                            Name = currentName,
                            Email = currentEmail ?? "",
                            Phone = currentPhone ?? ""
                        });
                        Log.Information($"发现联系人: {currentName} <{currentEmail}> 电话:{currentPhone}");
                    }
                    inVCard = false;
                }
                else if (inVCard)
                {
                    // FN: 姓名 - 精确匹配
                    if (line.StartsWith("FN:", StringComparison.OrdinalIgnoreCase))
                        currentName = line.Substring(3).Trim();
                    // N: 姓名（备用格式，优先级低于FN）
                    else if (line.StartsWith("N:", StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(currentName))
                    {
                        var parts = line.Substring(2).Trim().Split(';');
                        if (parts.Length >= 2)
                            currentName = (parts[1] + " " + parts[0]).Trim(); // 姓+名
                    }
                    // TEL;TYPE=CELL,PREF: 电话 - 优先匹配
                    else if (line.StartsWith("TEL;TYPE=CELL,PREF:", StringComparison.OrdinalIgnoreCase))
                    {
                        var colonIndex = line.IndexOf(':');
                        if (colonIndex > 0)
                            currentPhone = line.Substring(colonIndex + 1).Trim();
                    }
                    // TEL;TYPE=CELL: 电话
                    else if (line.StartsWith("TEL;TYPE=CELL:", StringComparison.OrdinalIgnoreCase))
                    {
                        var colonIndex = line.IndexOf(':');
                        if (colonIndex > 0)
                            currentPhone = line.Substring(colonIndex + 1).Trim();
                    }
                    // TEL: 电话 - 最后备选
                    else if (line.StartsWith("TEL:", StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(currentPhone))
                    {
                        var colonIndex = line.IndexOf(':');
                        if (colonIndex > 0)
                            currentPhone = line.Substring(colonIndex + 1).Trim();
                    }
                    // EMAIL;TYPE=OTHER: 邮箱
                    else if (line.StartsWith("EMAIL;TYPE=OTHER:", StringComparison.OrdinalIgnoreCase))
                    {
                        var colonIndex = line.IndexOf(':');
                        if (colonIndex > 0)
                            currentEmail = line.Substring(colonIndex + 1).Trim();
                    }
                    // EMAIL;TYPE=HOME,PREF: 邮箱
                    else if (line.StartsWith("EMAIL;TYPE=HOME,PREF:", StringComparison.OrdinalIgnoreCase))
                    {
                        var colonIndex = line.IndexOf(':');
                        if (colonIndex > 0)
                            currentEmail = line.Substring(colonIndex + 1).Trim();
                    }
                    // EMAIL: 邮箱 - 最后备选
                    else if (line.StartsWith("EMAIL:", StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(currentEmail))
                    {
                        var colonIndex = line.IndexOf(':');
                        if (colonIndex > 0)
                            currentEmail = line.Substring(colonIndex + 1).Trim();
                    }
                }
            }

            return contacts;
        }
    }

    public class CardDavContact
    {
        public string Url { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public string VCardData { get; set; }
        public string ETag { get; set; }
    }
}