using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace MailConverter
{
    /// <summary>
    /// 激活服务 - 基于邮件、姓名、安装日期生成激活码
    /// </summary>
    public static class ActivationService
    {
        /// <summary>
        /// 生成激活码
        /// </summary>
        /// <param name="email">邮箱地址</param>
        /// <param name="name">姓名</param>
        /// <param name="installDate">安装日期</param>
        /// <returns>激活码</returns>
        public static string GenerateActivationCode(string email, string name, DateTime installDate)
        {
            // 清理输入
            email = (email ?? "").Trim().ToLower();
            name = (name ?? "").Trim();

            // 组合原始字符串
            string rawString = $"{email}|{name}|{installDate:yyyy-MM-dd}|MailConverter_v1";

            // 使用SHA256生成哈希
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(rawString));
                StringBuilder sb = new StringBuilder();

                // 将哈希转换为可读格式 (XXXX-XXXX-XXXX-XXXX)
                for (int i = 0; i < bytes.Length; i++)
                {
                    sb.Append(bytes[i].ToString("X2"));
                    if (i > 0 && (i + 1) % 4 == 0 && i < bytes.Length - 1)
                        sb.Append("-");
                }

                // 取前19个字符 + 4位校验 (格式: XXXX-XXXX-XXXX-XXXX)
                string code = sb.ToString().Substring(0, 19).ToUpper();

                // 添加校验位
                code = AddCheckDigit(code);

                return code;
            }
        }

        /// <summary>
        /// 添加校验位
        /// </summary>
        private static string AddCheckDigit(string code)
        {
            int sum = 0;
            foreach (char c in code.Replace("-", ""))
            {
                sum += (c >= 'A' && c <= 'Z') ? c - 'A' + 10 : c - '0';
            }
            int checkDigit = (100 - (sum % 100)) % 100;
            return $"{code}-{checkDigit:D2}";
        }

        /// <summary>
        /// 验证激活码
        /// </summary>
        public static bool VerifyActivationCode(string email, string name, DateTime installDate, string activationCode)
        {
            if (string.IsNullOrWhiteSpace(activationCode))
                return false;

            string expectedCode = GenerateActivationCode(email, name, installDate);
            return string.Equals(expectedCode, activationCode, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 从激活码中提取信息（用于验证）
        /// </summary>
        public static bool ValidateActivationCode(string activationCode, string storedEmail, string storedName, DateTime installDate)
        {
            if (string.IsNullOrWhiteSpace(activationCode))
                return false;

            // 验证格式
            if (!IsValidFormat(activationCode))
                return false;

            // 验证校验位
            if (!VerifyCheckDigit(activationCode))
                return false;

            // 验证与存储的信息是否匹配
            string expectedCode = GenerateActivationCode(storedEmail, storedName, installDate);
            return string.Equals(expectedCode, activationCode, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 验证激活码格式
        /// </summary>
        public static bool IsValidFormat(string code)
        {
            if (string.IsNullOrWhiteSpace(code))
                return false;

            // 格式: XXXX-XXXX-XXXX-XXXX-XX
            string[] parts = code.Split('-');
            if (parts.Length != 5)
                return false;

            if (parts[0].Length != 4 || parts[1].Length != 4 ||
                parts[2].Length != 4 || parts[3].Length != 4 || parts[4].Length != 2)
                return false;

            return true;
        }

        /// <summary>
        /// 验证校验位
        /// </summary>
        private static bool VerifyCheckDigit(string code)
        {
            string[] parts = code.Split('-');
            if (parts.Length != 5)
                return false;

            string codeWithoutCheck = $"{parts[0]}-{parts[1]}-{parts[2]}-{parts[3]}";
            int sum = 0;
            foreach (char c in codeWithoutCheck.Replace("-", ""))
            {
                sum += (c >= 'A' && c <= 'Z') ? c - 'A' + 10 : c - '0';
            }

            int expectedCheck = (100 - (sum % 100)) % 100;
            int actualCheck = int.Parse(parts[4]);

            return expectedCheck == actualCheck;
        }

        /// <summary>
        /// 保存激活信息
        /// </summary>
        public static void SaveActivation(string email, string name, string activationCode)
        {
            var configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config");
            if (!Directory.Exists(configPath))
                Directory.CreateDirectory(configPath);

            var data = $"{email}|{name}|{activationCode}";
            File.WriteAllText(Path.Combine(configPath, "activation.dat"), data);
        }

        /// <summary>
        /// 加载激活信息
        /// </summary>
        public static bool LoadActivation(out string email, out string name, out string activationCode)
        {
            email = "";
            name = "";
            activationCode = "";

            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config", "activation.dat");
            if (!File.Exists(filePath))
                return false;

            try
            {
                string data = File.ReadAllText(filePath);
                data = data.Trim();
                string[] parts = data.Split('|');
                if (parts.Length >= 3)
                {
                    email = parts[0];
                    name = parts[1];
                    activationCode = parts[2];
                    return true;
                }
            }
            catch { }

            return false;
        }

    }
}
