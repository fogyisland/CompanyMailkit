using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using Serilog;

namespace MailConverter
{
    /// <summary>
    /// Azure 资源管理服务 - 自动创建 Storage Account
    /// 支持三种模式: PowerShell Az模块 / Azure CLI / 手动输入
    /// </summary>
    public class AzureResourceService
    {
        private string _subscriptionId;
        private string _storageAccountName;
        private string _containerName;
        private string _accountKey;
        private string _sasToken;
        private string _resourceGroup;

        public bool IsLoggedIn { get; private set; }
        public string SubscriptionId => _subscriptionId;
        public string StorageAccountName => _storageAccountName;
        public string ContainerName => _containerName;
        public string SasToken => _sasToken;
        public string AccountKey => _accountKey;

        /// <summary>
        /// 检查 Az PowerShell 模块是否可用
        /// </summary>
        public static bool IsAzModuleAvailable()
        {
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = "powershell",
                    Arguments = "-Command \"Get-Module -ListAvailable -Name Az.Storage -ErrorAction SilentlyContinue | Select-Object -First 1\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(startInfo))
                {
                    var output = process.StandardOutput.ReadToEnd();
                    process.WaitForExit();
                    return output.Contains("Az.Storage") || output.Contains("ModuleType");
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 检查 Azure CLI 是否可用
        /// </summary>
        public static bool IsAzureCliAvailable()
        {
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = "az",
                    Arguments = "--version",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(startInfo))
                {
                    var output = process.StandardOutput.ReadToEnd();
                    process.WaitForExit();
                    return output.Contains("azure-cli");
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 使用 Azure CLI 登录
        /// </summary>
        public bool LoginWithAzureCli(Action<string> progressCallback = null)
        {
            try
            {
                progressCallback?.Invoke("正在检查 Azure CLI 登录状态...");

                // 检查是否已登录
                var checkProcess = new ProcessStartInfo
                {
                    FileName = "az",
                    Arguments = "account show",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(checkProcess))
                {
                    var output = process.StandardOutput.ReadToEnd();
                    var error = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    if (process.ExitCode == 0 && !string.IsNullOrEmpty(output))
                    {
                        // 已登录，解析订阅 ID
                        var lines = output.Split('\n');
                        foreach (var line in lines)
                        {
                            if (line.Contains("\"id\""))
                            {
                                var parts = line.Split(':');
                                if (parts.Length >= 2)
                                {
                                    _subscriptionId = parts[1].Trim().Replace("\"", "").Replace(",", "");
                                }
                            }
                        }
                        IsLoggedIn = true;
                        Log.Information("Azure CLI 已登录, Subscription: {Sub}", _subscriptionId);
                        return true;
                    }
                }

                // 需要登录
                progressCallback?.Invoke("请在浏览器中完成 Azure 登录...");

                var loginProcess = new ProcessStartInfo
                {
                    FileName = "az",
                    Arguments = "login --use-device-code",
                    UseShellExecute = true,
                    CreateNoWindow = false
                };

                using (var process = Process.Start(loginProcess))
                {
                    process.WaitForExit();
                }

                // 再次检查
                return LoginWithAzureCli(progressCallback);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Azure CLI 登录失败");
                return false;
            }
        }

        /// <summary>
        /// 使用 Azure CLI 创建 Storage Account
        /// </summary>
        public bool CreateStorageAccountWithCli(string resourceGroup, string location = "southeastasia")
        {
            try
            {
                _resourceGroup = resourceGroup;
                _storageAccountName = $"pstmig{DateTime.Now:yyyyMMddHHmmss}".ToLower();

                var args = $"storage account create --name {_storageAccountName} --resource-group {resourceGroup} --location {location} --sku Standard_LRS --kind StorageV2";

                var startInfo = new ProcessStartInfo
                {
                    FileName = "az",
                    Arguments = args,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(startInfo))
                {
                    var output = process.StandardOutput.ReadToEnd();
                    var error = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    if (process.ExitCode != 0)
                    {
                        Log.Error("创建 Storage Account 失败: {Error}", error);
                        return false;
                    }

                    Log.Information("Storage Account 创建成功: {Name}", _storageAccountName);
                    return true;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "创建 Storage Account 失败");
                return false;
            }
        }

        /// <summary>
        /// 使用 Azure CLI 创建 Container
        /// </summary>
        public bool CreateContainerWithCli(string containerName)
        {
            try
            {
                _containerName = containerName ?? "pstfiles";

                var args = $"storage container create --name {_containerName} --account-name {_storageAccountName} --auth-mode login";

                var startInfo = new ProcessStartInfo
                {
                    FileName = "az",
                    Arguments = args,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(startInfo))
                {
                    var output = process.StandardOutput.ReadToEnd();
                    var error = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    Log.Information("Container 创建完成: {Result}", output.Contains("created") || output.Contains("already exists") ? "成功" : output);
                    return true;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "创建 Container 失败");
                return false;
            }
        }

        /// <summary>
        /// 使用 Azure CLI 获取 Account Key
        /// </summary>
        public string GetAccountKeyWithCli(string resourceGroup)
        {
            try
            {
                var args = $"storage account keys list --account-name {_storageAccountName} --resource-group {resourceGroup} --query '[0].value' -o tsv";

                var startInfo = new ProcessStartInfo
                {
                    FileName = "az",
                    Arguments = args,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(startInfo))
                {
                    var output = process.StandardOutput.ReadToEnd();
                    var error = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    if (process.ExitCode == 0 && !string.IsNullOrWhiteSpace(output))
                    {
                        _accountKey = output.Trim();
                        Log.Information("获取 Account Key 成功");
                        return _accountKey;
                    }

                    Log.Error("获取 Account Key 失败: {Error}", error);
                    return null;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取 Account Key 失败");
                return null;
            }
        }

        /// <summary>
        /// 使用 Azure CLI 生成 SAS Token
        /// </summary>
        public string GenerateSasTokenWithCli(int validityDays = 7)
        {
            try
            {
                var args = $"storage container generate-sas --name {_containerName} --account-name {_storageAccountName} --permissions rwl --expiry {(DateTime.UtcNow.AddDays(validityDays)):yyyy-MM-ddTHH:mm:ssZ} --output tsv";

                var startInfo = new ProcessStartInfo
                {
                    FileName = "az",
                    Arguments = args,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(startInfo))
                {
                    var output = process.StandardOutput.ReadToEnd();
                    var error = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    if (process.ExitCode == 0 && !string.IsNullOrWhiteSpace(output))
                    {
                        _sasToken = "?" + output.Trim();
                        Log.Information("SAS Token 生成成功，有效期 {Days} 天", validityDays);
                        return _sasToken;
                    }

                    Log.Error("生成 SAS Token 失败: {Error}", error);
                    return null;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "生成 SAS Token 失败");
                return null;
            }
        }

        /// <summary>
        /// 完整自动化流程 (使用 Azure CLI)
        /// </summary>
        public AzureStorageInfo SetupStorageAccount(
            string subscriptionId,
            string resourceGroup,
            string containerName,
            string location = "southeastasia",
            Action<string> progressCallback = null)
        {
            var result = new AzureStorageInfo();

            try
            {
                _subscriptionId = subscriptionId;

                // 1. 设置订阅
                progressCallback?.Invoke("正在设置订阅...");
                var subProcess = new ProcessStartInfo
                {
                    FileName = "az",
                    Arguments = $"account set --subscription {subscriptionId}",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(subProcess))
                {
                    process.WaitForExit();
                }

                // 2. 创建资源组 (如果不存在)
                progressCallback?.Invoke("正在检查资源组...");
                var rgProcess = new ProcessStartInfo
                {
                    FileName = "az",
                    Arguments = $"group create --name {resourceGroup} --location {location} --output none",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(rgProcess))
                {
                    process.WaitForExit();
                }

                // 3. 创建 Storage Account
                progressCallback?.Invoke("正在创建 Storage Account...");
                if (!CreateStorageAccountWithCli(resourceGroup, location))
                {
                    result.ErrorMessage = "创建 Storage Account 失败";
                    return result;
                }

                // 4. 创建 Container
                progressCallback?.Invoke("正在创建 Container...");
                if (!CreateContainerWithCli(containerName))
                {
                    result.ErrorMessage = "创建 Container 失败";
                    return result;
                }

                // 5. 获取 Account Key
                progressCallback?.Invoke("正在获取 Account Key...");
                var key = GetAccountKeyWithCli(resourceGroup);
                if (string.IsNullOrEmpty(key))
                {
                    result.ErrorMessage = "获取 Account Key 失败";
                    return result;
                }

                // 6. 生成 SAS Token
                progressCallback?.Invoke("正在生成 SAS Token...");
                var sasToken = GenerateSasTokenWithCli(7);
                if (string.IsNullOrEmpty(sasToken))
                {
                    result.ErrorMessage = "生成 SAS Token 失败";
                    return result;
                }

                result.Success = true;
                result.StorageAccount = _storageAccountName;
                result.Container = _containerName;
                result.AccountKey = _accountKey;
                result.SasToken = _sasToken;
                result.BlobEndpoint = $"https://{_storageAccountName}.blob.core.windows.net/{_containerName}";

                progressCallback?.Invoke($"完成！\nStorage Account: {_storageAccountName}\nContainer: {_containerName}");

                Log.Information("Azure Storage 自动配置完成: {Account}/{Container}", _storageAccountName, _containerName);
                return result;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "自动化配置失败: {Msg}", ex.Message);
                result.ErrorMessage = ex.Message;
                return result;
            }
        }

        /// <summary>
        /// 检查环境是否就绪
        /// </summary>
        public static string CheckEnvironment()
        {
            if (IsAzureCliAvailable())
                return "Azure CLI";
            if (IsAzModuleAvailable())
                return "PowerShell Az";
            return null;
        }
    }

    /// <summary>
    /// Azure Storage 配置信息
    /// </summary>
    public class AzureStorageInfo
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public string StorageAccount { get; set; }
        public string Container { get; set; }
        public string AccountKey { get; set; }
        public string SasToken { get; set; }
        public string BlobEndpoint { get; set; }
    }
}
