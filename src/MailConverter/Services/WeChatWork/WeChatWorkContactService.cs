using System;
using System.Net.Http;
using System.Text.Json;
using Serilog;

namespace MailConverter.Services.WeChatWork
{
    public class WeChatWorkContactService
    {
        public const string DefaultApiBase = "https://qyapi.weixin.qq.com/cgi-bin";

        private readonly HttpClient _http;
        private readonly string _apiBase;

        public WeChatWorkContactService(string apiBase = null, HttpClient httpClient = null)
        {
            _apiBase = string.IsNullOrWhiteSpace(apiBase) ? DefaultApiBase : apiBase.TrimEnd('/');
            _http = httpClient ?? new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
        }

        /// <summary>
        /// 调用 /cgi-bin/gettoken 获取 access_token
        /// 成功返回 token, 失败抛 WeChatWorkApiException
        /// </summary>
        public string GetAccessToken(string corpId, string corpSecret)
        {
            if (string.IsNullOrWhiteSpace(corpId))
                throw new ArgumentException("CorpID 不能为空", nameof(corpId));
            if (string.IsNullOrWhiteSpace(corpSecret))
                throw new ArgumentException("CorpSecret 不能为空", nameof(corpSecret));

            var url = $"{_apiBase}/gettoken?corpid={Uri.EscapeDataString(corpId)}&corpsecret={Uri.EscapeDataString(corpSecret)}";
            try
            {
                var json = _http.GetStringAsync(url).GetAwaiter().GetResult();
                var resp = JsonSerializer.Deserialize<GetTokenResponse>(json);
                if (resp == null || resp.ErrCode != 0 || string.IsNullOrEmpty(resp.AccessToken))
                {
                    var msg = resp?.ErrMsg ?? "unknown";
                    Log.Warning("企业微信 gettoken 失败: errcode={Errcode}, errmsg={Errmsg}", resp?.ErrCode, msg);
                    throw new WeChatWorkApiException(resp?.ErrCode ?? -1, $"gettoken 失败: {msg}");
                }
                Log.Information("企业微信 gettoken 成功, expires_in={ExpiresIn}s", resp.ExpiresIn);
                return resp.AccessToken;
            }
            catch (WeChatWorkApiException) { throw; }
            catch (Exception ex)
            {
                Log.Error(ex, "企业微信 gettoken 网络异常");
                throw new WeChatWorkApiException(-1, $"gettoken 网络异常: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 测试连接: 调用 gettoken 验证凭据
        /// </summary>
        public bool TestConnection(string corpId, string corpSecret, out string error)
        {
            try
            {
                GetAccessToken(corpId, corpSecret);
                error = "";
                return true;
            }
            catch (WeChatWorkApiException ex)
            {
                error = ex.Message;
                return false;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }
    }

    public class WeChatWorkApiException : Exception
    {
        public int ErrCode { get; }
        public WeChatWorkApiException(int errCode, string message) : base(message) { ErrCode = errCode; }
        public WeChatWorkApiException(int errCode, string message, Exception inner) : base(message, inner) { ErrCode = errCode; }
    }
}
