using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text.Json;
using System.Threading.Tasks;

namespace MailConverter.Services
{
    public class ExchangeRbacManager
    {
        private readonly string _appId;
        private readonly string _serviceId;
        private readonly string _tenantId;
        private readonly string _adminEmail;
        private readonly IntPtr _windowHandle;
        private readonly string _logPath;
        private string _clientSecret;

        public ExchangeRbacManager(string appId, string serviceId, string tenantId, string adminEmail, IntPtr windowHandle, string clientSecret = null)
        {
            _appId = appId;
            _serviceId = serviceId;
            _tenantId = tenantId;
            _adminEmail = adminEmail;
            _windowHandle = windowHandle;
            _clientSecret = clientSecret;
            _logPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs", "O365Online", "ServicePrincipal", $"ServicePrincipal_{DateTime.Now:yyyyMMdd}.log");
        }

        private void Log(string message)
        {
            try
            {
                var dir = System.IO.Path.GetDirectoryName(_logPath);
                if (!System.IO.Directory.Exists(dir))
                    System.IO.Directory.CreateDirectory(dir);
                System.IO.File.AppendAllText(_logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}\n");
            }
            catch { }
        }

        // 使用客户端密钥获取应用程序 Token (应用程序权限)
        private string GetClientSecretToken()
        {
            Log("使用客户端密钥获取应用程序 Token...");
            try
            {
                var tokenUrl = $"https://login.microsoftonline.com/{_tenantId}/oauth2/v2.0/token";
                var scope = "https://outlook.office365.com/.default";

                var postData = $"client_id={Uri.EscapeDataString(_appId)}&scope={Uri.EscapeDataString(scope)}&client_secret={Uri.EscapeDataString(_clientSecret)}&grant_type=client_credentials";

                var client = new System.Net.Http.HttpClient();
                var content = new System.Net.Http.StringContent(postData, System.Text.Encoding.UTF8, "application/x-www-form-urlencoded");
                var response = client.PostAsync(tokenUrl, content).Result;
                var responseContent = response.Content.ReadAsStringAsync().Result;

                if (response.IsSuccessStatusCode)
                {
                    using (var json = JsonDocument.Parse(responseContent))
                    {
                    var token = json.RootElement.GetProperty("access_token").GetString();
                    Log($"客户端密钥获取 Token 成功, Token长度: {token?.Length ?? 0}");

                    // 解析 Token 的 roles claim 用于调试
                    if (!string.IsNullOrEmpty(token))
                    {
                        try
                        {
                            var parts = token.Split('.');
                            if (parts.Length >= 2)
                            {
                                var payload = parts[1];
                                while (payload.Length % 4 != 0) payload += "=";
                                var jsonPayload = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(payload));
                                var jwt = JsonSerializer.Deserialize<JsonElement>(jsonPayload);
                                var aud = jwt.GetProperty("aud").GetString();
                                var roles = jwt.TryGetProperty("roles", out var r) ? r.ToString() : "";
                                Log($"Token aud: {aud}, roles: {roles}");
                            }
                        }
                        catch (Exception ex) { Log($"Token解析失败: {ex.Message}"); }
                    }

                    return token;
                    }
                }

                Log($"获取令牌失败: {responseContent}");
                return null;
            }
            catch (Exception ex)
            {
                Log($"获取客户端密钥 Token 异常: {ex.Message}");
                return null;
            }
        }

        // 主线程调用 - 执行 OAuth2 登录并返回 Token (委托权限)
        public async Task<(bool Success, string Message, string AccessToken)> LoginAndGetTokenAsync()
        {
            Log("开始 OAuth2 登录...");
            try
            {
                var app = Microsoft.Identity.Client.PublicClientApplicationBuilder.Create(_appId)
                    .WithAuthority(Microsoft.Identity.Client.AzureCloudInstance.AzurePublic, _tenantId)
                    .WithDefaultRedirectUri()
                    .Build();

                // 强制请求 Exchange 资源的 Token
                string[] scopes = new string[] {
                    "https://outlook.office365.com/.default"
                };

                // 获取已缓存的账户
                var accounts = await app.GetAccountsAsync().ConfigureAwait(false);
                var account = accounts.FirstOrDefault();

                Microsoft.Identity.Client.AuthenticationResult authResult;

                if (account != null)
                {
                    // 尝试使用缓存的 Token
                    try
                    {
                        authResult = await app.AcquireTokenSilent(scopes, account)
                            .WithTenantId(_tenantId)
                            .ExecuteAsync()
                            .ConfigureAwait(false);
                    }
                    catch
                    {
                        // 缓存失败，弹窗获取
                        authResult = await app.AcquireTokenInteractive(scopes)
                            .WithParentActivityOrWindow(_windowHandle)
                            .ExecuteAsync()
                            .ConfigureAwait(false);
                    }
                }
                else
                {
                    // 没有缓存，强制弹窗
                    authResult = await app.AcquireTokenInteractive(scopes)
                        .WithParentActivityOrWindow(_windowHandle)
                        .ExecuteAsync()
                        .ConfigureAwait(false);
                }

                Log($"OAuth2 登录成功, Token长度: {authResult.AccessToken?.Length ?? 0}");

                // 解析 Token 的 aud claim 用于调试
                try
                {
                    var parts = authResult.AccessToken.Split('.');
                    if (parts.Length >= 2)
                    {
                        var payload = parts[1];
                        while (payload.Length % 4 != 0) payload += "=";
                        var json = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(payload));
                        var jwt = JsonSerializer.Deserialize<JsonElement>(json);
                        var aud = jwt.GetProperty("aud").GetString();
                        var scp = jwt.TryGetProperty("scp", out var s) ? s.GetString() : "";
                        Log($"Token aud: {aud}, scopes: {scp}");
                    }
                }
                catch (Exception ex) { Log($"Token解析失败: {ex.Message}"); }

                return (true, "OAuth2 登录成功", authResult.AccessToken);
            }
            catch (Microsoft.Identity.Client.MsalException ex)
            {
                Log($"OAuth2 登录失败: {ex.Message}");
                return (false, $"OAuth2 登录失败: {ex.Message}", null);
            }
        }

        // 使用客户端密钥获取 Token
        public async Task<(bool Success, string Message, string AccessToken)> GetTokenAsync()
        {
            if (string.IsNullOrEmpty(_clientSecret))
            {
                return (false, "请填写 Client Secret", null);
            }

            var token = GetClientSecretToken();
            if (!string.IsNullOrEmpty(token))
            {
                return (true, "客户端密钥认证成功", token);
            }
            return (false, "客户端密钥获取 Token 失败", null);
        }

        // 获取 ServicePrincipal 的 Object ID (企业应用程序 ID)
        private string GetServicePrincipalId(string graphToken)
        {
            Log("获取 ServicePrincipal ID...");
            try
            {
                var client = new System.Net.Http.HttpClient();
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", graphToken);

                // 调用 Graph API 获取 ServicePrincipal
                var url = $"https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '{_appId}'";
                var response = client.GetAsync(url).Result;
                var content = response.Content.ReadAsStringAsync().Result;

                Log($"Graph API 响应: {content}");

                if (response.IsSuccessStatusCode)
                {
                    using (var json = JsonDocument.Parse(content))
                    {
                    var value = json.RootElement.GetProperty("value");
                    if (value.GetArrayLength() > 0)
                    {
                        var spId = value[0].GetProperty("id").GetString();
                        Log($"ServicePrincipal ID: {spId}");
                        return spId;
                    }
                    }
                }
                Log($"获取 ServicePrincipal 失败: {content}");
                return null;
            }
            catch (Exception ex)
            {
                Log($"获取 ServicePrincipal 异常: {ex.Message}");
                return null;
            }
        }

        // 后台执行 - 使用 PowerShell SDK 连接 Exchange Online
        public async Task<(bool Success, string Message)> EnsureExchangePermissionsAsync(string accessToken)
        {
            Log("开始执行权限配置...");

            // 从管理员邮箱提取组织域名（保留完整域名）
            string organization = null;
            if (!string.IsNullOrEmpty(_adminEmail) && _adminEmail.Contains("@"))
            {
                organization = _adminEmail.Split('@')[1]; // 例如 booming.com
            }
            else if (!string.IsNullOrEmpty(_tenantId))
            {
                if (_tenantId.Contains(".onmicrosoft.com"))
                {
                    organization = _tenantId; // 例如 xxx.onmicrosoft.com
                }
                else
                {
                    organization = _tenantId + ".onmicrosoft.com";
                }
            }
            Log($"组织域名: {organization}");

            return await Task.Run(() =>
            {
                try
                {
                    Log("创建 PowerShell Runspace...");
                    using (var runspace = RunspaceFactory.CreateRunspace())
                    {
                        runspace.Open();
                        Log("Runspace 已打开");

                        using (var ps = PowerShell.Create())
                        {
                            ps.Runspace = runspace;

                            // 1. 设置本地模块路径并检查 ExchangeOnlineManagement 模块
                            Log("检查 ExchangeOnlineManagement 模块...");
                            var localModulePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PowerShellModules", "ExchangeOnlineManagement", "3.9.2", "netFramework");
                            ps.AddScript($@"
                                $env:PSModulePath = '{localModulePath}' + ';' + $env:PSModulePath
                                if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {{
                                    throw 'ExchangeOnlineManagement 模块未安装'
                                }}
                                Import-Module '{Path.Combine(localModulePath, "ExchangeOnlineManagement.psm1")}' -DisableNameChecking -ErrorAction Stop
                            ");
                            ps.Invoke();
                            if (ps.HadErrors)
                            {
                                var error = ps.Streams.Error[0];
                                Log($"模块检查失败: {error.Exception?.Message}");
                                return (false, $"模块检查失败: {error.Exception?.Message ?? "未知错误"}");
                            }
                            Log("模块检查通过");
                            ps.Commands.Clear();

                            // 2. 使用 AccessToken 连接 Exchange Online
                            Log("连接 Exchange Online...");
                            ps.AddCommand("Connect-ExchangeOnline");
                            ps.AddParameter("AccessToken", accessToken);
                            ps.AddParameter("Organization", organization);
                            ps.AddParameter("ShowBanner", false);
                            ps.AddParameter("ErrorAction", "Stop");
                            ps.Invoke();

                            if (ps.HadErrors)
                            {
                                var error = ps.Streams.Error[0];
                                var errorDetails = error.Exception?.Message ?? "未知错误";
                                var errorRecord = error.ToString();
                                Log($"连接失败: {errorDetails}");
                                Log($"详细错误: {errorRecord}");

                                // 检查是否是权限问题
                                if (errorDetails.Contains("UnAuthorized") || errorDetails.Contains("Unauthorized") || errorDetails.Contains("403") || errorDetails.Contains("role assigned"))
                                {
                                    return (false, "连接失败: " + errorDetails + "\n\n解决方法:\n1. 在 Azure 门户中进入: Azure AD → 企业应用程序\n2. 找到你的应用 (App ID: " + _appId + ")\n3. 点击'用户和组' → '添加用户/组'\n4. 选择角色 → 搜索并选择 'Exchange 管理员' (Exchange Administrator)\n5. 分配后等待几分钟后重试");
                                }
                                return (false, "连接失败: " + errorDetails);
                            }
                            Log("连接成功!");
                            ps.Commands.Clear();

                            // 3. 获取 Graph API Token 用于查询 ServicePrincipal
                            var graphTokenUrl = $"https://login.microsoftonline.com/{_tenantId}/oauth2/v2.0/token";
                            var graphScope = "https://graph.microsoft.com/.default";
                            var graphPostData = $"client_id={Uri.EscapeDataString(_appId)}&scope={Uri.EscapeDataString(graphScope)}&client_secret={Uri.EscapeDataString(_clientSecret)}&grant_type=client_credentials";
                            var graphClient = new System.Net.Http.HttpClient();
                            var graphContent = new System.Net.Http.StringContent(graphPostData, System.Text.Encoding.UTF8, "application/x-www-form-urlencoded");
                            var graphResponse = graphClient.PostAsync(graphTokenUrl, graphContent).Result;
                            var graphResponseContent = graphResponse.Content.ReadAsStringAsync().Result;
                            string graphToken = null;
                            if (graphResponse.IsSuccessStatusCode)
                            {
                                using (var json = JsonDocument.Parse(graphResponseContent))
                                {
                                graphToken = json.RootElement.GetProperty("access_token").GetString();
                                Log($"Graph Token 获取成功, 长度: {graphToken?.Length ?? 0}");
                                }
                            }
                            else
                            {
                                Log($"Graph Token 获取失败: {graphResponseContent}");
                            }

                            // 4. 检查 ServicePrincipal 是否已存在
                            Log("检查 ServicePrincipal...");
                            ps.AddScript($"Get-ServicePrincipal -AppId '{_appId}' -ErrorAction SilentlyContinue");
                            var spResult = ps.Invoke();
                            ps.Commands.Clear();

                            bool spExists = spResult.Count > 0;
                            Log($"ServicePrincipal 存在: {spExists}");

                            if (!spExists)
                            {
                                try
                                {
                                    // 获取 ServicePrincipal 的 Object ID
                                    string serviceId = _serviceId;
                                    if (string.IsNullOrEmpty(serviceId) && !string.IsNullOrEmpty(graphToken))
                                    {
                                        serviceId = GetServicePrincipalId(graphToken);
                                    }

                                    if (string.IsNullOrEmpty(serviceId))
                                    {
                                        return (false, "无法获取 ServicePrincipal ID。请确认：\n1. 已在 Azure 门户为应用授予了 API 权限\n2. 已点击'代表管理员授予同意'");
                                    }

                                    Log($"使用 ServiceId: {serviceId} 创建 ServicePrincipal...");

                                    // 使用 -Confirm:$false 自动确认
                                    ps.AddScript($"New-ServicePrincipal -AppId '{_appId}' -ServiceId '{serviceId}' -Confirm:$false -ErrorAction Stop");
                                    ps.Invoke();

                                    if (ps.HadErrors)
                                    {
                                        // 检查是否是"已在使用中"错误，如果是则表示已创建成功
                                        var error = ps.Streams.Error[0];
                                        var errorMsg = error.Exception?.Message ?? "";
                                        Log($"创建返回错误: {errorMsg}");

                                        // 如果是已在使用中，说明已创建成功
                                        if (errorMsg.Contains("已在使用中") || errorMsg.Contains("already in use"))
                                        {
                                            Log("ServicePrincipal 已存在（通过错误信息确认）");
                                        }
                                        else
                                        {
                                            return (false, $"创建 ServicePrincipal 失败: {errorMsg}");
                                        }
                                    }
                                    else
                                    {
                                        Log("创建成功");
                                    }
                                }
                                catch (Exception spEx)
                                {
                                    // 如果创建 ServicePrincipal 时出错，检查是否是"已在使用中"
                                    if (spEx.Message.Contains("已在使用中") || spEx.Message.Contains("already in use"))
                                    {
                                        Log("ServicePrincipal 已存在（通过异常确认）");
                                    }
                                    else
                                    {
                                        Log($"ServicePrincipal 检查异常: {spEx.Message}");
                                    }
                                }
                            }
                            ps.Commands.Clear();

                            // 5. 检查是否已在角色组中 (改进检查逻辑)
                            Log("检查角色组成员...");
                            ps.AddScript($@"
                                $members = Get-RoleGroupMember -Identity 'Organization Management' -ErrorAction SilentlyContinue
                                if ($members) {{
                                    $exists = $members | Where-Object {{ $_.RawIdentity -like '*{_appId}*' -or $_.ExternalDirectoryObjectId -eq '{_appId}' }}
                                    if ($exists) {{ $true }} else {{ $false }}
                                }} else {{
                                    $false
                                }}
                            ");
                            var memberResult = ps.Invoke();
                            ps.Commands.Clear();

                            bool alreadyMember = memberResult.Count > 0 && (bool)memberResult[0].BaseObject;
                            Log($"已在角色组: {alreadyMember}");

                            if (!alreadyMember)
                            {
                                Log("添加到角色组...");
                                ps.AddCommand("Add-RoleGroupMember");
                                ps.AddParameter("Identity", "Organization Management");
                                ps.AddParameter("Member", _appId);
                                ps.AddParameter("ErrorAction", "Stop");
                                ps.Invoke();

                                if (ps.HadErrors)
                                {
                                    var error = ps.Streams.Error[0];
                                    string errorMsg = error.Exception?.Message ?? "";
                                    // 检查是否是"已是成员"相关的错误
                                    if (!errorMsg.Contains("already a member") &&
                                        !errorMsg.Contains("已是成员") &&
                                        !errorMsg.Contains("已经是组"))
                                    {
                                        Log($"添加角色组失败: {errorMsg}");
                                        return (false, $"添加到角色组失败: {errorMsg}");
                                    }
                                    Log("已在角色组中 (错误忽略)");
                                }
                                Log("添加成功");
                            }

                            // 7. 断开连接
                            Log("断开连接...");
                            ps.Commands.Clear();
                            ps.AddCommand("Disconnect-ExchangeOnline");
                            ps.AddParameter("Confirm", false);
                            ps.Invoke();

                            Log("执行完成!");
                            return (true, "Exchange RBAC 权限配置成功！\n1. OAuth2 登录成功\n2. 已连接 Exchange Online\n3. ServicePrincipal 已创建/验证\n4. 已添加到 Organization Management 角色组");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"执行异常: {ex.Message}");
                    return (false, $"执行异常: {ex.Message}");
                }
            });
        }
    }
}