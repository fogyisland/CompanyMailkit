using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;

namespace MailConverter
{
    /// <summary>
    /// Excel/CSV导入服务
    /// </summary>
    public class ExcelCsvImportService
    {
        private Action<string> _logCallback;
        private Action<int, int> _progressCallback;

        public void SetCallbacks(Action<string> logCallback, Action<int, int> progressCallback)
        {
            _logCallback = logCallback;
            _progressCallback = progressCallback;
        }

        /// <summary>
        /// 导入Excel/CSV文件并转换为邮件
        /// </summary>
        public async Task<bool> ImportToPstAsync(
            string[] filePaths,
            string outputPstPath,
            string subjectPrefix = "")
        {
            try
            {
                Log($"开始导入 {filePaths.Length} 个文件");

                // 收集所有邮件数据
                var allMails = new List<Dictionary<string, string>>();

                foreach (var filePath in filePaths)
                {
                    Log($"正在读取: {Path.GetFileName(filePath)}");

                    string ext = Path.GetExtension(filePath).ToLower();
                    List<Dictionary<string, string>> records;

                    if (ext == ".xlsx" || ext == ".xls")
                    {
                        records = ReadExcel(filePath);
                    }
                    else if (ext == ".csv")
                    {
                        records = ReadCsv(filePath);
                    }
                    else
                    {
                        Log($"不支持的文件格式: {ext}");
                        continue;
                    }

                    allMails.AddRange(records);
                    Log($"  读取到 {records.Count} 条记录");
                }

                if (allMails.Count == 0)
                {
                    throw new Exception("未读取到任何数据");
                }

                Log($"共 {allMails.Count} 条记录，准备生成PST");

                // 创建临时目录保存EML
                var tempDir = Path.Combine(Path.GetTempPath(), "excel_import_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(tempDir);

                try
                {
                    // 转换为EML
                    int processed = 0;
                    foreach (var record in allMails)
                    {
                        await Task.Run(() => ConvertToEml(record, tempDir, subjectPrefix, processed + 1));
                        processed++;
                        _progressCallback?.Invoke(processed, allMails.Count);

                        if (processed % 50 == 0)
                        {
                            Log($"已处理: {processed}/{allMails.Count}");
                        }
                    }

                    Log($"EML生成完成，开始创建PST");

                    // 调用Python脚本创建PST
                    await CreatePstFromEml(tempDir, outputPstPath, "Excel导入");
                }
                finally
                {
                    // 清理临时目录
                    try { Directory.Delete(tempDir, true); } catch { }
                }

                Log($"导入完成: {outputPstPath}");
                return true;
            }
            catch (Exception ex)
            {
                Log($"导入失败: {ex.Message}");
                throw;
            }
        }

        private List<Dictionary<string, string>> ReadExcel(string filePath)
        {
            var records = new List<Dictionary<string, string>>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheets.First();

                // 获取表头
                var headerRow = worksheet.Row(1);
                var headers = new List<string>();
                foreach (var cell in headerRow.CellsUsed())
                {
                    headers.Add(cell.GetValue<string>().Trim());
                }

                // 读取数据行
                var rows = worksheet.RowsUsed().Skip(1);
                foreach (var row in rows)
                {
                    var record = new Dictionary<string, string>();
                    for (int i = 0; i < headers.Count; i++)
                    {
                        var cell = row.Cell(i + 1);
                        string value = cell.IsEmpty() ? "" : cell.GetValue<string>();
                        record[headers[i]] = value;
                    }

                    // 跳过空行
                    if (record.Values.Any(v => !string.IsNullOrWhiteSpace(v)))
                    {
                        records.Add(record);
                    }
                }
            }

            return records;
        }

        private List<Dictionary<string, string>> ReadCsv(string filePath)
        {
            var records = new List<Dictionary<string, string>>();

            using (var reader = new StreamReader(filePath, System.Text.Encoding.UTF8))
            {
                var config = new CsvConfiguration(System.Globalization.CultureInfo.InvariantCulture)
                {
                    HasHeaderRecord = true,
                    MissingFieldFound = null
                };

                using (var csv = new CsvReader(reader, config))
                {
                    csv.Read();
                    csv.ReadHeader();
                    var headers = csv.HeaderRecord;

                    while (csv.Read())
                    {
                        var record = new Dictionary<string, string>();
                        foreach (var header in headers)
                        {
                            record[header] = csv.GetField(header) ?? "";
                        }
                        records.Add(record);
                    }
                }
            }

            return records;
        }

        private void ConvertToEml(Dictionary<string, string> data, string outputDir, string subjectPrefix, int index)
        {
            // 确定字段映射
            string subject = GetFieldValue(data, new[] { "Subject", "主题", "标题", "subject", "Title" });
            string from = GetFieldValue(data, new[] { "From", "发件人", "发送者", "from", "Sender" });
            string to = GetFieldValue(data, new[] { "To", "收件人", "接收者", "to", "Recipient" });
            string cc = GetFieldValue(data, new[] { "CC", "抄送", "cc", "Cc" });
            string body = GetFieldValue(data, new[] { "Body", "内容", "正文", "body", "Content", "Description", "描述" });
            string date = GetFieldValue(data, new[] { "Date", "日期", "时间", "date", "Created", "创建时间" });
            string priority = GetFieldValue(data, new[] { "Priority", "优先级", "重要程度", "priority" });

            // 生成主题
            if (string.IsNullOrWhiteSpace(subject))
            {
                subject = $"{subjectPrefix} - 记录 {index}";
            }
            else
            {
                subject = $"{subjectPrefix}{subject}";
            }

            // 清理文件名
            string safeSubject = SanitizeFileName(subject);

            string emlPath = Path.Combine(outputDir, $"{safeSubject}_{index}.eml");

            // 生成EML内容
            var lines = new List<string>();
            lines.Add($"Subject: {subject}");
            if (!string.IsNullOrWhiteSpace(from))
                lines.Add($"From: {from}");
            if (!string.IsNullOrWhiteSpace(to))
                lines.Add($"To: {to}");
            if (!string.IsNullOrWhiteSpace(cc))
                lines.Add($"Cc: {cc}");
            if (!string.IsNullOrWhiteSpace(date))
                lines.Add($"Date: {date}");

            lines.Add("MIME-Version: 1.0");
            lines.Add("Content-Type: text/plain; charset=utf-8");
            lines.Add("Content-Transfer-Encoding: 7bit");
            lines.Add("");

            // 添加正文和所有字段
            var fullBody = new List<string>();
            if (!string.IsNullOrWhiteSpace(body))
            {
                fullBody.Add(body);
            }

            // 添加其他字段
            foreach (var kvp in data)
            {
                string key = kvp.Key;
                string value = kvp.Value;

                // 跳过已使用的字段
                if (new[] { "Subject", "主题", "标题", "subject", "Title", "Body", "内容", "正文", "body", "Content" }.Contains(key))
                    continue;

                if (!string.IsNullOrWhiteSpace(value))
                {
                    fullBody.Add($"{key}: {value}");
                }
            }

            lines.Add(string.Join("\r\n", fullBody));

            File.WriteAllText(emlPath, string.Join("\r\n", lines), System.Text.Encoding.UTF8);
        }

        private string GetFieldValue(Dictionary<string, string> data, string[] possibleKeys)
        {
            foreach (var key in possibleKeys)
            {
                if (data.TryGetValue(key, out string value) && !string.IsNullOrWhiteSpace(value))
                {
                    return value.Trim();
                }
            }
            return "";
        }

        private string SanitizeFileName(string fileName)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(c, '_');
            }
            if (fileName.Length > 100)
                fileName = fileName.Substring(0, 100);
            return fileName;
        }

        private async Task CreatePstFromEml(string emlDir, string pstPath, string pstName)
        {
            // 删除旧PST
            if (File.Exists(pstPath))
            {
                File.Delete(pstPath);
            }

            var exeDir = AppDomain.CurrentDomain.BaseDirectory;
            var scriptPath = Path.Combine(exeDir, "script", "create_pst.py");

            if (!File.Exists(scriptPath))
            {
                throw new Exception("Python脚本不存在: " + scriptPath);
            }

            var pythonExe = Program.GetPythonExecutable();
            if (string.IsNullOrEmpty(pythonExe))
            {
                throw new Exception("Python环境未找到，请重新安装程序");
            }

            var startInfo = Program.CreatePythonStartInfo(pythonExe, $"\"{scriptPath}\" \"{pstPath}\" \"{pstName}\"");

            using (var process = System.Diagnostics.Process.Start(startInfo))
            {
                if (process != null)
                {
                    process.WaitForExit(300000);

                    var output = process.StandardOutput.ReadToEnd();
                    var error = process.StandardError.ReadToEnd();

                    Log("PST创建输出: " + output);
                    if (!string.IsNullOrEmpty(error))
                        Log("PST创建错误: " + error);
                }
            }

            if (!File.Exists(pstPath))
            {
                throw new Exception("PST文件创建失败");
            }
        }

        private void Log(string message)
        {
            _logCallback?.Invoke(message);

            var logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs", "mailconverter.log");
            File.AppendAllText(logPath, $"[ExcelCsvImport] {message}\n");
        }
    }
}
