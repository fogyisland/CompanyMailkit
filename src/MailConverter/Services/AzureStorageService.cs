using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Serilog;
using Azure;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;

namespace MailConverter
{
    /// <summary>
    /// Azure Blob Storage 服务 - 使用 AzCopy 上传 PST 文件
    /// </summary>
    public class AzureStorageService
    {
        private string _sasUrl;

        public bool IsConnected => !string.IsNullOrEmpty(_sasUrl);
        public string SasUrl => _sasUrl;

        /// <summary>
        /// 获取 AzCopy 的完整路径 (优先使用程序目录下的版本)
        /// </summary>
        public static string GetAzCopyPath()
        {
            // 优先使用程序目录下的 azcopy.exe
            var appDir = AppDomain.CurrentDomain.BaseDirectory;
            var localAzCopy = Path.Combine(appDir, "azcopy.exe");

            if (File.Exists(localAzCopy))
            {
                return localAzCopy;
            }

            // 回退到 PATH 中的 azcopy
            return "azcopy";
        }

        /// <summary>
        /// 检查 AzCopy 是否可用
        /// </summary>
        public static bool IsAzCopyAvailable()
        {
            try
            {
                var azCopyPath = GetAzCopyPath();
                Program.PurviewLogger.Information("[PurView] 检查 AzCopy 路径: {Path}", azCopyPath);

                if (!File.Exists(azCopyPath))
                {
                    Program.PurviewLogger.Error("[PurView] AzCopy 文件不存在: {Path}", azCopyPath);
                    return false;
                }

                var startInfo = new ProcessStartInfo
                {
                    FileName = azCopyPath,
                    Arguments = "--version",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(startInfo))
                {
                    process.WaitForExit(5000); // 5秒超时
                    var available = process.ExitCode == 0;
                    Program.PurviewLogger.Information("[PurView] AzCopy 可用状态: {Available}", available);
                    return available;
                }
            }
            catch (Exception ex)
            {
                Program.PurviewLogger.Error(ex, "[PurView] 检查 AzCopy 失败: {Msg}", ex.Message);
                return false;
            }
        }

        /// <summary>
        /// 使用 SAS URL 连接
        /// </summary>
        public bool ConnectWithSasUrl(string sasUrl)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(sasUrl))
                {
                    Program.PurviewLogger.Error("SAS URL 不能为空");
                    return false;
                }

                _sasUrl = sasUrl.Trim();

                if (!_sasUrl.StartsWith("https://"))
                {
                    Program.PurviewLogger.Error("SAS URL 格式错误，应以 https:// 开头");
                    return false;
                }

                Program.PurviewLogger.Information("SAS URL 已设置");
                return true;
            }
            catch (Exception ex)
            {
                Program.PurviewLogger.Error(ex, "设置 SAS URL 失败: {Msg}", ex.Message);
                return false;
            }
        }

        /// <summary>
        /// 验证 SAS URL 是否可用 (使用 Azure SDK 尝试写入测试文件)
        /// </summary>
        public async Task<string> ValidateSasUrlAsync()
        {
            if (string.IsNullOrEmpty(_sasUrl))
                return "SAS URL 未设置";

            try
            {
                Program.PurviewLogger.Information("[PurView] 使用 Azure SDK 验证 SAS URL (写入测试)");

                var containerClient = new BlobContainerClient(new Uri(_sasUrl));

                // 生成测试文件名
                string testBlobName = $"sas-test-temp-{Guid.NewGuid()}.txt";
                BlobClient blobClient = containerClient.GetBlobClient(testBlobName);

                // 上传测试数据
                string content = "SAS Write Test";
                using (var ms = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(content)))
                {
                    await blobClient.UploadAsync(ms, overwrite: true);
                }

                Program.PurviewLogger.Information("[PurView] SAS URL 验证成功，写入权限正常");

                // 尝试删除测试文件 (如果 SAS 有删除权限)
                try
                {
                    await blobClient.DeleteIfExistsAsync();
                    Program.PurviewLogger.Information("[PurView] 测试文件已删除");
                }
                catch (Exception ex)
                {
                    Program.PurviewLogger.Information("[PurView] 测试文件删除失败 (可能无删除权限): {Msg}", ex.Message);
                }

                return "验证成功: SAS URL 可用 (写入权限正常)";
            }
            catch (RequestFailedException ex)
            {
                Program.PurviewLogger.Error("[PurView] SAS URL 验证失败 - 状态码: {Status}, 错误代码: {ErrorCode}, 信息: {Msg}",
                    ex.Status, ex.ErrorCode, ex.Message);

                string errorMessage;
                switch (ex.ErrorCode)
                {
                    case "AuthenticationFailed":
                        errorMessage = "验证失败: 认证失败 (密钥错误或签名被篡改)";
                        break;
                    case "AuthorizationPermissionMismatch":
                        errorMessage = "验证失败: 权限不足 (需要写入权限)";
                        break;
                    case "ExpiredError":
                        errorMessage = "验证失败: SAS 已过期";
                        break;
                    case "ContainerNotFound":
                    case "ResourceNotFound":
                        errorMessage = "验证失败: 容器不存在";
                        break;
                    default:
                        errorMessage = $"验证失败: {ex.ErrorCode} - {ex.Message}";
                        break;
                }
                return errorMessage;
            }
            catch (Exception ex)
            {
                Program.PurviewLogger.Error(ex, "[PurView] SAS URL 验证异常: {Msg}", ex.Message);
                return $"验证异常: {ex.Message}";
            }
        }

        /// <summary>
        /// 使用 AzCopy 上传单个 PST 文件 (支持进度报告)
        /// </summary>
        public async Task<bool> UploadWithAzCopyAsync(string localFilePath, string blobName = null, Action<int, string> progressCallback = null)
        {
            if (string.IsNullOrEmpty(_sasUrl))
            {
                Program.PurviewLogger.Error("未设置 SAS URL");
                return false;
            }

            if (!File.Exists(localFilePath))
            {
                Program.PurviewLogger.Error("文件不存在: {Path}", localFilePath);
                return false;
            }

            try
            {
                var fileName = blobName ?? Path.GetFileName(localFilePath);

                // 正确构造目标 URL: baseUrl/container/fileName?queryString
                var sasUrl = _sasUrl.Trim();
                var queryIndex = sasUrl.IndexOf('?');
                string baseUrl, queryString;
                if (queryIndex >= 0)
                {
                    baseUrl = sasUrl.Substring(0, queryIndex).TrimEnd('/');
                    queryString = sasUrl.Substring(queryIndex);
                }
                else
                {
                    baseUrl = sasUrl.TrimEnd('/');
                    queryString = "";
                }
                var destinationUrl = $"{baseUrl}/{fileName}{queryString}";

                Program.PurviewLogger.Information("[PurView] 开始上传: {File} -> {Dest}", localFilePath, destinationUrl);

                var args = $"copy \"{localFilePath}\" \"{destinationUrl}\" --blob-type BlockBlob --recursive";

                var startInfo = new ProcessStartInfo
                {
                    FileName = GetAzCopyPath(),
                    Arguments = args,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (var process = new Process())
                {
                    process.StartInfo = startInfo;
                    process.Start();

                    var outputLines = new List<string>();
                    var errorLines = new List<string>();

                    // 同步读取输出
                    string line;
                    while ((line = process.StandardOutput.ReadLine()) != null)
                    {
                        outputLines.Add(line);
                        // 解析进度行，如 "INFO: percent : 50%"
                        if (line.Contains("percent") && line.Contains("%"))
                        {
                            try
                            {
                                var parts = line.Split(':');
                                foreach (var part in parts)
                                {
                                    if (part.Trim().EndsWith("%"))
                                    {
                                        var percentStr = part.Trim().Replace("%", "");
                                        if (int.TryParse(percentStr, out int percent))
                                        {
                                            progressCallback?.Invoke(percent, $"{percent}%");
                                        }
                                        break;
                                    }
                                }
                            }
                            catch { }
                        }
                    }

                    while ((line = process.StandardError.ReadLine()) != null)
                    {
                        errorLines.Add(line);
                    }

                    process.WaitForExit();

                    if (process.ExitCode == 0)
                    {
                        Program.PurviewLogger.Information("[PurView] 上传成功: {File}", fileName);
                        progressCallback?.Invoke(100, "上传完成");
                        return true;
                    }
                    else
                    {
                        var allOutput = string.Join("\n", outputLines);
                        var allError = string.Join("\n", errorLines);
                        var fullError = string.IsNullOrEmpty(allError) ? allOutput : (allOutput + "\n" + allError);
                        Program.PurviewLogger.Error("[PurView] 上传失败, 退出码: {Code}, 输出: {Output}", process.ExitCode, fullError);
                        progressCallback?.Invoke(0, "上传失败");
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                Program.PurviewLogger.Error(ex, "上传文件失败: {Path}", localFilePath);
                return false;
            }
        }

        /// <summary>
        /// 检查 AzCopy 版本
        /// </summary>
        public static string GetAzCopyVersion()
        {
            try
            {
                var azCopyPath = GetAzCopyPath();
                Program.PurviewLogger.Information("[PurView] 获取 AzCopy 版本, 路径: {Path}", azCopyPath);

                var startInfo = new ProcessStartInfo
                {
                    FileName = azCopyPath,
                    Arguments = "--version",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(startInfo))
                {
                    process.WaitForExit(5000);
                    var output = process.StandardOutput.ReadToEnd();
                    var version = output.Trim();
                    Program.PurviewLogger.Information("[PurView] AzCopy 版本: {Version}", version);
                    return version;
                }
            }
            catch (Exception ex)
            {
                Program.PurviewLogger.Error(ex, "[PurView] 获取 AzCopy 版本失败: {Msg}", ex.Message);
                return "未安装: " + ex.Message;
            }
        }
    }
}
