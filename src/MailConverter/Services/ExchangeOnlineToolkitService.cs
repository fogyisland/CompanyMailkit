using System;
using Microsoft.Exchange.WebServices.Data;
using Serilog;

namespace MailConverter
{
    /// <summary>
    /// Exchange Online 管理工具箱
    /// </summary>
    public class ExchangeOnlineToolkitService
    {
        private ExchangeService _service;
        private string _email;
        private string _password;
        private string _domain;

        public bool Connect(string email, string password, string domain = null)
        {
            try
            {
                _email = email;
                _password = password;
                _domain = domain;

                Log.Information("连接 Exchange Online: {Email}", email);

                _service = new ExchangeService(ExchangeVersion.Exchange2016);
                _service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");

                if (!string.IsNullOrEmpty(domain))
                {
                    _service.Credentials = new System.Net.NetworkCredential(email, password, domain);
                }
                else
                {
                    _service.Credentials = new System.Net.NetworkCredential(email, password);
                }

                _service.AutodiscoverUrl(email, RedirectionUrlValidationCallback);

                // 验证连接 - 只测试连接
                var inbox = Folder.Bind(_service, WellKnownFolderName.Inbox);

                Log.Information("Exchange Online 连接成功");
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Exchange Online 连接失败: {Msg}", ex.Message);
                return false;
            }
        }

        private bool RedirectionUrlValidationCallback(string url)
        {
            return url.ToLower().StartsWith("https://");
        }

        /// <summary>
        /// 发送测试邮件
        /// </summary>
        public bool SendTestEmail(string to, string subject, string body)
        {
            try
            {
                var email = new EmailMessage(_service)
                {
                    Subject = subject,
                    Body = new MessageBody(BodyType.Text, body)
                };
                email.ToRecipients.Add(to);
                email.Send();
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "发送测试邮件失败");
                return false;
            }
        }
    }

    public class MailboxStats
    {
        public string Email { get; set; }
        public int TotalItems { get; set; }
        public string Size { get; set; }
    }

    public class SearchResult
    {
        public string Subject { get; set; }
        public DateTime Date { get; set; }
        public string Sender { get; set; }
    }

    public class FolderInfo
    {
        public string Name { get; set; }
        public string Id { get; set; }
        public int TotalItems { get; set; }
    }

    public class MailboxQuota
    {
        public string Email { get; set; }
        public int TotalItems { get; set; }
        public long ProhibitedSendReceiveQuota { get; set; }
        public long ProhibitedSendQuota { get; set; }
        public int TotalItemCount { get; set; }
    }
}
