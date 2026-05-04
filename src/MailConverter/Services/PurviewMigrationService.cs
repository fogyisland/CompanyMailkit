using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Azure.Core;
using Azure.Identity;
using Serilog;

namespace MailConverter
{
    /// <summary>
    /// Microsoft Purview PST 迁移服务
    /// 通过 Azure Blob Storage 上传 PST，Microsoft 自动导入到目标邮箱
    /// </summary>
    public class PurviewMigrationService
    {
        private string _tenantId;
        private string _clientId;
        private string _clientSecret;
        private string _accessToken;
        private HttpClient _httpClient;

        public bool IsConnected => !string.IsNullOrEmpty(_accessToken);
        public string TenantId => _tenantId;

        /// <summary>
        /// 使用客户端密钥连接
        /// </summary>
        public bool Connect(string tenantId, string clientId, string clientSecret)
        {
            try
            {
                _tenantId = tenantId;
                _clientId = clientId;
                _clientSecret = clientSecret;
                _httpClient = new HttpClient();

                // 获取访问令牌
                var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
                var context = new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" });
                var tokenResult = credential.GetTokenAsync(context).Result;
                _accessToken = tokenResult.Token;

                Log.Information("PurviewMigrationService 连接成功: Tenant={TenantId}", tenantId);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "PurviewMigrationService 连接失败: {Msg}", ex.Message);
                return false;
            }
        }

        /// <summary>
        /// 获取迁移批次列表
        /// </summary>
        public async Task<List<MigrationBatchInfo>> GetMigrationBatchesAsync()
        {
            var batches = new List<MigrationBatchInfo>();

            try
            {
                // 使用 Exchange Online PowerShell 方式获取
                // 由于 Graph API 对 PST 迁移批次支持有限，这里简化处理
                Log.Information("获取迁移批次列表 (简化模式)");
                return batches;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "获取迁移批次列表失败: {Msg}", ex.Message);
            }

            return batches;
        }

        /// <summary>
        /// 创建 PST 迁移端点 (通过 PowerShell)
        /// 返回端点创建所需的参数信息
        /// </summary>
        public PSTEndpointInfo PreparePstEndpoint(string storageAccount, string container, string sasToken)
        {
            return new PSTEndpointInfo
            {
                EndpointName = $"PSTEndpoint_{DateTime.Now:yyyyMMddHHmmss}",
                StorageAccount = storageAccount,
                Container = container,
                SasToken = sasToken,
                AzureBlobPath = $"https://{storageAccount}.blob.core.windows.net/{container}"
            };
        }

        /// <summary>
        /// 生成迁移 CSV 内容
        /// </summary>
        public static string GenerateMigrationCsv(List<Tuple<string, string>> mappings)
        {
            var sb = new StringBuilder();
            sb.AppendLine("filename,mailbox,isArchive");

            foreach (var mapping in mappings)
            {
                // filename: PST 文件名 (user@domain.pst)
                // mailbox: 目标邮箱地址
                // isArchive: 是否导入到存档
                sb.AppendLine($"{mapping.Item1},{mapping.Item2},False");
            }

            return sb.ToString();
        }

        /// <summary>
        /// 获取 PowerShell 命令用于创建迁移
        /// </summary>
        public static string GenerateMigrationCommands(PSTEndpointInfo endpoint, string batchName, string csvContent)
        {
            var sb = new StringBuilder();

            // 保存 CSV 到临时文件
            var csvPath = Path.Combine(Path.GetTempPath(), $"migration_{DateTime.Now:yyyyMMddHHmmss}.csv");
            File.WriteAllText(csvPath, csvContent);

            sb.AppendLine($"# 创建 PST 迁移端点");
            sb.AppendLine($"$password = ConvertTo-SecureString '{endpoint.SasToken}' -AsPlainText -Force");
            sb.AppendLine($"$creds = New-Object System.Management.Automation.PSCredential('{endpoint.StorageAccount}', $password)");
            sb.AppendLine($"New-MigrationEndpoint -Name '{endpoint.EndpointName}' -PstImport -ExchangeServer 'N/A' -Credentials $creds -SourceMailboxLegacyDN 'N/A' -TargetMailboxLegacyDN 'N/A'");
            sb.AppendLine();
            sb.AppendLine($"# 创建迁移批次");
            sb.AppendLine($"$csvData = Get-Content '{csvPath}' -Raw -Encoding Byte");
            sb.AppendLine($"New-MigrationBatch -Name '{batchName}' -SourceEndpoint '{endpoint.EndpointName}' -CSVData $csvData -AutoStart");
            sb.AppendLine();
            sb.AppendLine($"# 查看状态");
            sb.AppendLine($"Get-MigrationBatch -Identity '{batchName}'");

            return sb.ToString();
        }
    }

    /// <summary>
    /// PST 迁移端点信息
    /// </summary>
    public class PSTEndpointInfo
    {
        public string EndpointName { get; set; }
        public string StorageAccount { get; set; }
        public string Container { get; set; }
        public string SasToken { get; set; }
        public string AzureBlobPath { get; set; }
    }

    /// <summary>
    /// 迁移批次信息
    /// </summary>
    public class MigrationBatchInfo
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Status { get; set; }
        public int TotalCount { get; set; }
        public int SyncedCount { get; set; }
        public int FailedCount { get; set; }
        public DateTime? CreatedDateTime { get; set; }

        public string StatusText => $"{Status} ({SyncedCount}/{TotalCount}, Failed: {FailedCount})";
    }
}
