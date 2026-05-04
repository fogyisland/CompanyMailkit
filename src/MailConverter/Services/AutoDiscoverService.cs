using System;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using MailKit.Net.Imap;
using MailKit.Security;

namespace MailConverter
{
    /// <summary>
    /// IMAP账户配置
    /// </summary>
    public class ImapAccountConfig
    {
        public string Email { get; set; }
        public string Password { get; set; }
        public string Host { get; set; }
        public int Port { get; set; } = 993;
        public bool UseSsl { get; set; } = true;
    }

    /// <summary>
    /// AutoDiscover服务 - 自动发现IMAP/SMTP配置
    /// </summary>
    public class AutoDiscoverService
    {
        private static bool _certificateValidationPassed = false;

        static AutoDiscoverService()
        {
            // 忽略SSL证书验证（仅用于本地工具）
            ServicePointManager.ServerCertificateValidationCallback =
                (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) => true;
        }

        /// <summary>
        /// 自动发现IMAP配置
        /// </summary>
        public async Task<ImapAccountConfig> AutoDiscoverAsync(string email, string password)
        {
            // 提取域名
            int atIndex = email.IndexOf('@');
            string domain = atIndex > 0 ? email.Substring(atIndex + 1) : email;

            // 尝试常见的IMAP服务器配置
            var configs = GetCommonImapConfigs(domain);

            foreach (var config in configs)
            {
                try
                {
                    Log($"尝试连接: {config.Host}:{config.Port}");

                    using (var client = new ImapClient())
                    {
                        await client.ConnectAsync(config.Host, config.Port, config.UseSsl);

                        // 尝试认证
                        try
                        {
                            await client.AuthenticateAsync(email, password);
                            await client.DisconnectAsync(true);

                            Log($"成功! IMAP: {config.Host}:{config.Port}");
                            return config;
                        }
                        catch (AuthenticationException)
                        {
                            // 连接成功但认证失败，说明服务器配置正确
                            await client.DisconnectAsync(true);
                            Log($"服务器配置正确: {config.Host}:{config.Port}，但密码错误");
                            return config;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"尝试 {config.Host} 失败: {ex.Message}");
                    continue;
                }
            }

            // 如果自动发现失败，返回默认配置
            return new ImapAccountConfig
            {
                Email = email,
                Password = password,
                Host = $"imap.{domain}",
                Port = 993,
                UseSsl = true
            };
        }

        /// <summary>
        /// 测试IMAP连接
        /// </summary>
        public async Task<(bool Success, string Message)> TestConnectionAsync(string host, int port, bool useSsl, string email, string password)
        {
            try
            {
                Log($"测试连接: {host}:{port}");

                using (var client = new ImapClient())
                {
                    await client.ConnectAsync(host, port, useSsl);
                    await client.AuthenticateAsync(email, password);

                    await client.Inbox.OpenAsync(MailKit.FolderAccess.ReadOnly);
                    int count = client.Inbox.Count;

                    await client.DisconnectAsync(true);

                    return (true, $"连接成功! 收件箱有 {count} 封邮件");
                }
            }
            catch (Exception ex)
            {
                Log($"测试连接失败: {ex.Message}");
                return (false, $"连接失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取常见的IMAP服务器配置
        /// </summary>
        private System.Collections.Generic.List<ImapAccountConfig> GetCommonImapConfigs(string domain)
        {
            var configs = new System.Collections.Generic.List<ImapAccountConfig>();

            // 常见邮件服务商的IMAP配置
            switch (domain.ToLower())
            {
                case "gmail.com":
                case "googlemail.com":
                    configs.Add(new ImapAccountConfig { Host = "imap.gmail.com", Port = 993, UseSsl = true });
                    break;

                case "outlook.com":
                case "hotmail.com":
                case "live.com":
                    configs.Add(new ImapAccountConfig { Host = "outlook.office365.com", Port = 993, UseSsl = true });
                    break;

                case "qq.com":
                    configs.Add(new ImapAccountConfig { Host = "imap.qq.com", Port = 993, UseSsl = true });
                    break;

                case "163.com":
                    configs.Add(new ImapAccountConfig { Host = "imap.163.com", Port = 993, UseSsl = true });
                    break;

                case "126.com":
                    configs.Add(new ImapAccountConfig { Host = "imap.126.com", Port = 993, UseSsl = true });
                    break;

                case "sina.com":
                    configs.Add(new ImapAccountConfig { Host = "imap.sina.com", Port = 993, UseSsl = true });
                    break;

                case "aliyun.com":
                case "mail.aliyun.com":
                    configs.Add(new ImapAccountConfig { Host = "imap.aliyun.com", Port = 993, UseSsl = true });
                    break;

                case "foxmail.com":
                    configs.Add(new ImapAccountConfig { Host = "imap.foxmail.com", Port = 993, UseSsl = true });
                    break;

                default:
                    // 尝试常见的IMAP服务器
                    configs.Add(new ImapAccountConfig { Host = $"imap.{domain}", Port = 993, UseSsl = true });
                    configs.Add(new ImapAccountConfig { Host = $"imap.{domain}", Port = 143, UseSsl = false });
                    configs.Add(new ImapAccountConfig { Host = $"mail.{domain}", Port = 993, UseSsl = true });
                    configs.Add(new ImapAccountConfig { Host = $"mail.{domain}", Port = 143, UseSsl = false });
                    break;
            }

            // 设置邮箱和密码
            foreach (var config in configs)
            {
                config.Email = "";
                config.Password = "";
            }

            return configs;
        }

        private void Log(string message)
        {
            var logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs", "mailconverter.log");
            File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [AutoDiscover] {message}\n");
        }
    }
}
