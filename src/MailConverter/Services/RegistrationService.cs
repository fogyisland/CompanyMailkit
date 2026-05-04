using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.NetworkInformation;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace MailConverter
{
    /// <summary>
    /// 注册结果信息
    /// </summary>
    public class RegistrationResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public int? RemainingDays { get; set; }
        public int? TotalDays { get; set; }
        public string ExpireDate { get; set; }
        public string InstallDate { get; set; }
    }

    public class RegistrationService
    {
        private const string ApiInstallUrl = "https://www.booming.one/api/install";
        private const string ApiActivateUrl = "https://www.booming.one/api/activate-by-code";
        private const string ApiCheckUrl = "https://www.booming.one/api/install/check";

        /// <summary>
        /// 获取物理网卡的MAC地址（排除虚拟网卡）
        /// </summary>
        public string GetPhysicalMacAddress()
        {
            try
            {
                var interfaces = NetworkInterface.GetAllNetworkInterfaces()
                    .Where(ni => ni.OperationalStatus == OperationalStatus.Up)
                    .Where(ni => ni.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                    .Where(ni => ni.NetworkInterfaceType != NetworkInterfaceType.Tunnel)
                    .ToList();

                // 优先选择以太网卡和Wi-Fi
                var physical = interfaces
                    .Where(ni => ni.NetworkInterfaceType == NetworkInterfaceType.Ethernet ||
                                 ni.NetworkInterfaceType == NetworkInterfaceType.Wireless80211)
                    .FirstOrDefault();

                if (physical != null)
                {
                    return FormatMacAddress(physical.GetPhysicalAddress());
                }

                // 如果没有找到以太/Wi-Fi，返回第一个有效的物理网卡
                var first = interfaces.FirstOrDefault();
                if (first != null)
                {
                    return FormatMacAddress(first.GetPhysicalAddress());
                }

                return "";
            }
            catch (Exception ex)
            {
                Serilog.Log.Error(ex, "获取MAC地址失败");
                return "";
            }
        }

        /// <summary>
        /// 格式化MAC地址为标准格式 (XX-XX-XX-XX-XX-XX)
        /// </summary>
        private string FormatMacAddress(PhysicalAddress address)
        {
            var bytes = address.GetAddressBytes();
            return string.Join("-", bytes.Select(b => b.ToString("X2")));
        }

        /// <summary>
        /// 检查是否有物理网卡
        /// </summary>
        public bool HasPhysicalNetworkAdapter()
        {
            return NetworkInterface.GetAllNetworkInterfaces()
                .Any(ni => ni.OperationalStatus == OperationalStatus.Up &&
                           ni.NetworkInterfaceType != NetworkInterfaceType.Loopback &&
                           ni.NetworkInterfaceType != NetworkInterfaceType.Tunnel);
        }

        /// <summary>
        /// 检查注册状态（每日调用）
        /// </summary>
        public async Task<RegistrationResult> CheckRegistrationStatusAsync(string softwareName, string userEmail)
        {
            try
            {
                var jsonData = new Dictionary<string, object>
                {
                    { "softwareName", softwareName },
                    { "userEmail", userEmail }
                };

                var jsonContent = new StringContent(
                    JsonSerializer.Serialize(jsonData),
                    Encoding.UTF8,
                    "application/json");

                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(30);
                    var response = await client.PostAsync(ApiCheckUrl, jsonContent);
                    var result = await response.Content.ReadAsStringAsync();

                    Serilog.Log.Information("注册状态检查结果: {Result}, StatusCode: {StatusCode}", result, response.StatusCode);

                    return ParseCheckResult(result);
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Error(ex, "注册状态检查失败");
                return new RegistrationResult
                {
                    Success = false,
                    Message = ex.Message
                };
            }
        }

        private RegistrationResult ParseCheckResult(string result)
        {
            try
            {
                using (var doc = JsonDocument.Parse(result))
                {
                    var root = doc.RootElement;

                    // 检查返回的registered字段
                    bool registered = false;
                    if (root.TryGetProperty("registered", out var regProp))
                        registered = regProp.GetBoolean();

                    bool expired = false;
                    if (root.TryGetProperty("expired", out var expProp))
                        expired = expProp.GetBoolean();

                    if (registered && !expired && root.TryGetProperty("installation", out var installation))
                    {
                        int remainingDays = GetIntProperty(installation, "remainingDays", "remaining_days");
                        string installDate = GetStringProperty(installation, "installDate", "install_date");
                        string expireDate = GetStringProperty(installation, "expireDate", "expire_date");

                        // 如果有激活信息，优先使用激活的过期日期
                        if (root.TryGetProperty("activation", out var activation) && activation.ValueKind != JsonValueKind.Null)
                        {
                            string activationExpireDate = GetStringProperty(activation, "expireDate", "expire_date");
                            if (!string.IsNullOrEmpty(activationExpireDate))
                            {
                                expireDate = activationExpireDate;
                            }
                        }

                        // 如果剩余天数为0或未提供，尝试从过期日期计算
                        if (remainingDays <= 0 && !string.IsNullOrEmpty(expireDate))
                        {
                            if (DateTime.TryParse(expireDate, out DateTime expireDt))
                            {
                                remainingDays = (int)Math.Max(0, (expireDt - DateTime.Now).TotalDays);
                            }
                        }

                        return new RegistrationResult
                        {
                            Success = true,
                            Message = "有效",
                            RemainingDays = remainingDays,
                            ExpireDate = expireDate,
                            InstallDate = installDate
                        };
                    }
                    else if (!registered || expired)
                    {
                        return new RegistrationResult
                        {
                            Success = false,
                            Message = expired ? "已过期" : "未注册"
                        };
                    }

                    return new RegistrationResult
                    {
                        Success = false,
                        Message = "未知状态"
                    };
                }
            }
            catch
            {
                return new RegistrationResult
                {
                    Success = false,
                    Message = "解析响应失败"
                };
            }
        }

        /// <summary>
        /// 提交注册信息到服务器（试用版注册）
        /// </summary>
        public async Task<RegistrationResult> RegisterAsync(string softwareName, string softwareVersion, string userName, string userEmail, string organization, string macAddress)
        {
            try
            {
                var jsonData = new Dictionary<string, object>
                {
                    { "softwareName", softwareName },
                    { "softwareVersion", softwareVersion },
                    { "userName", userName },
                    { "userEmail", userEmail },
                    { "organization", organization ?? "" },
                    { "macAddress", macAddress }
                };

                var jsonContent = new StringContent(
                    JsonSerializer.Serialize(jsonData),
                    Encoding.UTF8,
                    "application/json");

                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(30);
                    var response = await client.PostAsync(ApiInstallUrl, jsonContent);
                    var result = await response.Content.ReadAsStringAsync();

                    Serilog.Log.Information("注册提交结果: {Result}, StatusCode: {StatusCode}", result, response.StatusCode);

                    return ParseRegistrationResult(result);
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Error(ex, "注册提交失败");
                return new RegistrationResult
                {
                    Success = false,
                    Message = ex.Message
                };
            }
        }

        /// <summary>
        /// 正式版激活（通过授权码）
        /// </summary>
        public async Task<RegistrationResult> ActivateByCodeAsync(string activationCode, string macAddress, string userName, string userEmail, string installDate)
        {
            try
            {
                var jsonData = new Dictionary<string, object>
                {
                    { "activationCode", activationCode },
                    { "macAddress", macAddress },
                    { "userName", userName },
                    { "userEmail", userEmail },
                    { "installDate", installDate }
                };

                var jsonContent = new StringContent(
                    JsonSerializer.Serialize(jsonData),
                    Encoding.UTF8,
                    "application/json");

                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(30);
                    var response = await client.PostAsync(ApiActivateUrl, jsonContent);
                    var result = await response.Content.ReadAsStringAsync();

                    Serilog.Log.Information("激活提交结果: {Result}, StatusCode: {StatusCode}", result, response.StatusCode);

                    return ParseActivationResult(result);
                }
            }
            catch (Exception ex)
            {
                Serilog.Log.Error(ex, "激活提交失败");
                return new RegistrationResult
                {
                    Success = false,
                    Message = ex.Message
                };
            }
        }

        /// <summary>
        /// 解析激活API响应（新的激活响应格式）
        /// </summary>
        private RegistrationResult ParseActivationResult(string result)
        {
            try
            {
                using (var doc = JsonDocument.Parse(result))
                {
                    var root = doc.RootElement;

                    if (root.TryGetProperty("success", out var successProp) && successProp.GetBoolean())
                    {
                        // 成功响应
                        string message = root.TryGetProperty("message", out var msgProp) ? msgProp.GetString() : "激活成功";
                        int totalDays = GetIntProperty(root, "totalDays", "total_days");
                        string registrationDate = GetStringProperty(root, "registrationDate", "registration_date");
                        string activateDate = GetStringProperty(root, "activateDate", "activate_date");
                        string expireDate = GetStringProperty(root, "expireDate", "expire_date");
                        string duration = GetStringProperty(root, "duration");
                        bool isRenewal = root.TryGetProperty("isRenewal", out var renewalProp) && renewalProp.GetBoolean();

                        // 计算剩余天数（从激活日期+总天数到今天）
                        int remainingDays = 0;
                        DateTime expireDt;
                        if (DateTime.TryParse(expireDate, out expireDt))
                        {
                            remainingDays = (int)Math.Max(0, (expireDt - DateTime.Now).TotalDays);
                        }
                        else if (totalDays > 0)
                        {
                            var activateDt = DateTime.TryParse(activateDate, out var ad) ? ad : DateTime.Now;
                            expireDt = activateDt.AddDays(totalDays);
                            remainingDays = (int)Math.Max(0, (expireDt - DateTime.Now).TotalDays);
                        }

                        return new RegistrationResult
                        {
                            Success = true,
                            Message = message,
                            RemainingDays = remainingDays,
                            TotalDays = totalDays,
                            ExpireDate = expireDate,
                            InstallDate = activateDate
                        };
                    }
                    else
                    {
                        // 失败响应
                        string errorMsg = root.TryGetProperty("error", out var errProp) ? errProp.GetString() : "激活失败";
                        return new RegistrationResult
                        {
                            Success = false,
                            Message = errorMsg
                        };
                    }
                }
            }
            catch
            {
                return new RegistrationResult
                {
                    Success = false,
                    Message = "解析响应失败"
                };
            }
        }

        private RegistrationResult ParseRegistrationResult(string result)
        {
            try
            {
                using (var doc = JsonDocument.Parse(result))
                {
                    var root = doc.RootElement;

                    if (root.TryGetProperty("success", out var successProp) && successProp.GetBoolean())
                    {
                        var installation = root.GetProperty("installation");

                        // 尝试获取各种可能的字段名（驼峰和下划线格式）
                        int remainingDays = GetIntProperty(installation, "remainingDays", "remaining_days", "remainingDays");
                        int totalDays = GetIntProperty(installation, "totalDays", "total_days", "TotalDays");
                        string installDate = GetStringProperty(installation, "installDate", "install_date", "installDate");
                        string expireDate = GetStringProperty(installation, "expireDate", "expire_date", "expireDate");

                        if (string.IsNullOrEmpty(installDate))
                            installDate = DateTime.Now.ToString("yyyy-MM-dd");

                        // 如果没有过期日期，根据 totalDays 计算
                        if (string.IsNullOrEmpty(expireDate) && DateTime.TryParse(installDate, out var parsedDate))
                        {
                            expireDate = parsedDate.AddDays(totalDays > 0 ? totalDays : 30).ToString("yyyy-MM-dd");
                        }

                        // 如果剩余天数为0或未提供，尝试从过期日期计算
                        if (remainingDays <= 0 && !string.IsNullOrEmpty(expireDate))
                        {
                            if (DateTime.TryParse(expireDate, out DateTime expireDt))
                            {
                                remainingDays = (int)Math.Max(0, (expireDt - DateTime.Now).TotalDays);
                            }
                        }

                        return new RegistrationResult
                        {
                            Success = true,
                            Message = "成功",
                            RemainingDays = remainingDays,
                            TotalDays = totalDays,
                            ExpireDate = expireDate,
                            InstallDate = installDate
                        };
                    }
                    else
                    {
                        string errorMsg = root.TryGetProperty("error", out var errProp) ? errProp.GetString() : "操作失败";

                        // 如果是"已注册"错误，从 installation 中提取信息
                        if (errorMsg.Contains("已注册") || errorMsg.Contains("already registered"))
                        {
                            if (root.TryGetProperty("installation", out var installation))
                            {
                                int remainingDays = GetIntProperty(installation, "remainingDays", "remaining_days");
                                string installDate = GetStringProperty(installation, "installDate", "install_date");
                                string expireDate = GetStringProperty(installation, "expireDate", "expire_date");

                                // 如果有激活信息，优先使用激活的过期日期
                                if (root.TryGetProperty("activation", out var activation) && activation.ValueKind != JsonValueKind.Null)
                                {
                                    string activationExpireDate = GetStringProperty(activation, "expireDate", "expire_date");
                                    if (!string.IsNullOrEmpty(activationExpireDate))
                                    {
                                        expireDate = activationExpireDate;
                                    }
                                }

                                // 如果剩余天数为0或未提供，尝试从过期日期计算
                                if (remainingDays <= 0 && !string.IsNullOrEmpty(expireDate))
                                {
                                    if (DateTime.TryParse(expireDate, out DateTime expireDt))
                                    {
                                        remainingDays = (int)Math.Max(0, (expireDt - DateTime.Now).TotalDays);
                                    }
                                }

                                return new RegistrationResult
                                {
                                    Success = true,
                                    Message = "已注册",
                                    RemainingDays = remainingDays,
                                    ExpireDate = expireDate,
                                    InstallDate = installDate
                                };
                            }
                        }

                        return new RegistrationResult
                        {
                            Success = false,
                            Message = errorMsg
                        };
                    }
                }
            }
            catch
            {
                return new RegistrationResult
                {
                    Success = false,
                    Message = "解析响应失败"
                };
            }
        }

        private int GetIntProperty(JsonElement element, params string[] propertyNames)
        {
            foreach (var name in propertyNames)
            {
                if (element.TryGetProperty(name, out var prop))
                {
                    return prop.GetInt32();
                }
            }
            return 0;
        }

        private string GetStringProperty(JsonElement element, params string[] propertyNames)
        {
            foreach (var name in propertyNames)
            {
                if (element.TryGetProperty(name, out var prop))
                {
                    var str = prop.GetString();
                    if (!string.IsNullOrEmpty(str))
                        return str;
                }
            }
            return "";
        }
    }
}
