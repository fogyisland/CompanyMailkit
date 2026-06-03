using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Serilog;
using MailConverter.Services.Calendars;
using MailConverter.Services.Contacts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailConverter
{
    public class PstExtractService
    {
        private const int olMailItem = 43;

        public bool ExtractToEml(string pstPath, string outputDir, IProgress<int> progress = null)
        {
            Log.Information("开始提取 PST: {PstPath}", pstPath);

            if (!File.Exists(pstPath))
            {
                Log.Error("PST 文件不存在: {Path}", pstPath);
                return false;
            }

            Directory.CreateDirectory(outputDir);
            Outlook.Application outlookApp = null;
            Outlook.NameSpace ns = null;

            try
            {
                try
                {
                    outlookApp = (Outlook.Application)Marshal.GetActiveObject("Outlook.Application");
                }
                catch
                {
                    outlookApp = new Outlook.Application();
                    System.Threading.Thread.Sleep(5000);
                }

                ns = outlookApp.GetNamespace("MAPI");
                Outlook.Folder targetFolder = null;

                // 尝试加载 PST
                try
                {
                    ns.AddStoreEx(pstPath, Outlook.OlStoreType.olStoreUnicode);
                }
                catch (Exception ex) { Log.Warning("添加 Store 失败，可能已挂载: {Msg}", ex.Message); }

                // 寻找对应的 Folder
                foreach (Outlook.Folder folder in ns.Folders)
                {
                    try
                    {
                        if (folder.Store != null && !string.IsNullOrEmpty(folder.Store.FilePath))
                        {
                            if (Path.GetFullPath(folder.Store.FilePath).Equals(Path.GetFullPath(pstPath), StringComparison.OrdinalIgnoreCase))
                            {
                                targetFolder = folder;
                                break;
                            }
                        }
                    }
                    catch { }
                }

                if (targetFolder == null) return false;

                int totalExtracted = 0;
                // 标记 PST 根目录：根目录下的邮件直接放入 basePath，根目录本身不创建子目录
                ExtractFolder(targetFolder, outputDir, ref totalExtracted, progress, isPstRoot: true);

                Log.Information("提取完成，共计: {Count} 封邮件", totalExtracted);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "提取失败");
                return false;
            }
            finally
            {
                if (ns != null) Marshal.ReleaseComObject(ns);
            }
        }

        private void ExtractFolder(Outlook.Folder folder, string basePath, ref int totalCount, IProgress<int> progress, bool isPstRoot = false)
        {
            string folderName = folder.Name;
            // PST 根目录不创建子目录（处理中文/英文/日文等不同语言的 "Outlook Data File"）
            if (isPstRoot || folderName == "Outlook Data File")
            {
                folderName = "";
            }

            string folderPath = string.IsNullOrEmpty(folderName) ? basePath : Path.Combine(basePath, CleanFileName(folderName));

            if (!string.IsNullOrEmpty(folderName))
            {
                Directory.CreateDirectory(folderPath);
            }

            Outlook.Items items = folder.Items;
            for (int i = 1; i <= items.Count; i++)
            {
                object item = null;
                try
                {
                    item = items[i];
                    if (item is Outlook.MailItem mailItem && (int)mailItem.Class == olMailItem)
                    {
                        string subject = string.IsNullOrEmpty(mailItem.Subject) ? "无主题" : mailItem.Subject;
                        string fileName = $"{CleanFileName(subject)}_{i}.eml";
                        string emlPath = Path.Combine(folderPath ?? basePath, fileName);

                        // 处理重名文件
                        int counter = 1;
                        while (File.Exists(emlPath))
                        {
                            fileName = $"{CleanFileName(subject)}_{i}_{counter}.eml";
                            emlPath = Path.Combine(folderPath ?? basePath, fileName);
                            counter++;
                        }

                        SaveMailAsEml(mailItem, emlPath);
                        totalCount++;
                        if (progress != null && totalCount % 50 == 0)
                            progress.Report(totalCount);
                    }
                }
                catch (Exception ex) { Log.Warning("处理邮件失败: {Msg}", ex.Message); }
                finally { if (item != null) Marshal.ReleaseComObject(item); }
            }
            Marshal.ReleaseComObject(items);

            // 处理子文件夹
            try
            {
                foreach (Outlook.Folder sub in folder.Folders)
                {
                    ExtractFolder(sub, basePath, ref totalCount, progress);
                    Marshal.ReleaseComObject(sub);
                }
            }
            catch { }
        }

        private void SaveMailAsEml(Outlook.MailItem mailItem, string emlPath)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                var attachments = mailItem.Attachments;
                bool hasAttachments = attachments.Count > 0;
                string boundary = "----=_Part_" + Guid.NewGuid().ToString("N");

                // --- HEADER 区 ---
                // 1. Subject (RFC 2047 编码处理中文)
                sb.AppendLine($"Subject: {EncodeHeaderValue(mailItem.Subject ?? "无主题")}");

                // 2. From - 简单直接获取
                string fromLine = GetFromLine(mailItem);
                sb.AppendLine($"From: {fromLine}");

                // 3. To / CC / BCC
                string to = "";
                string cc = "";
                string bcc = "";
                try { to = mailItem.To ?? ""; } catch { }
                try { cc = mailItem.CC ?? ""; } catch { }
                try { bcc = mailItem.BCC ?? ""; } catch { }

                if (!string.IsNullOrEmpty(to)) sb.AppendLine($"To: {to}");
                if (!string.IsNullOrEmpty(cc)) sb.AppendLine($"Cc: {cc}");
                if (!string.IsNullOrEmpty(bcc)) sb.AppendLine($"Bcc: {bcc}");

                // 4. Date (符合 RFC 2822 标准)
                sb.AppendLine($"Date: {GetRfc2822Date(mailItem)}");

                // 4.1 X-Original-Received-Time: 保留 PST 中的原始接收时间
                // 导入到 O365 时 Graph API 的 ReceivedDateTime 才能正确反映邮件的实际接收时间
                DateTime? extractedReceivedTime = null;
                string receivedTimeSource = "";
                try
                {
                    DateTime receivedTime = mailItem.ReceivedTime;
                    if (receivedTime.Year >= 1990 && receivedTime.Year <= 2100)
                    {
                        extractedReceivedTime = receivedTime;
                        receivedTimeSource = "ReceivedTime";
                    }
                }
                catch (Exception ex)
                {
                    Log.Warning("mailItem.ReceivedTime 读取失败: {Msg}", ex.Message);
                }

                // 兜底: 使用 MAPI 属性 PR_MESSAGE_DELIVERY_TIME (0x0E0F0003)
                if (!extractedReceivedTime.HasValue)
                {
                    try
                    {
                        var propAccessor = mailItem.PropertyAccessor;
                        var deliveryTime = propAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E0F0003");
                        if (deliveryTime is DateTime dt && dt.Year >= 1990 && dt.Year <= 2100)
                        {
                            extractedReceivedTime = dt;
                            receivedTimeSource = "PR_MESSAGE_DELIVERY_TIME";
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Warning("PR_MESSAGE_DELIVERY_TIME 读取失败: {Msg}", ex.Message);
                    }
                }

                if (extractedReceivedTime.HasValue)
                {
                    // 用 UTC 格式 (Z) 写入，避免后续解析时区歧义
                    DateTime utcReceived = extractedReceivedTime.Value.Kind == DateTimeKind.Utc
                        ? extractedReceivedTime.Value
                        : extractedReceivedTime.Value.ToUniversalTime();
                    sb.AppendLine($"X-Original-Received-Time: {utcReceived.ToString("yyyy-MM-ddTHH:mm:ssZ", System.Globalization.CultureInfo.InvariantCulture)}");
                    Log.Debug("写入 X-Original-Received-Time: {Time} (来源: {Source})", utcReceived, receivedTimeSource);
                }
                else
                {
                    Log.Warning("未能获取邮件接收时间: {Subject}", mailItem.Subject ?? "(无主题)");
                }

                // 5. Message-ID
                string entryId = GetProperty(mailItem, "EntryID");
                if (!string.IsNullOrEmpty(entryId))
                    sb.AppendLine($"Message-ID: <{entryId.GetHashCode()}_{DateTime.Now.Ticks}@pst-converter.local>");
                else
                    sb.AppendLine($"Message-ID: <{DateTime.Now.Ticks}@pst-converter.local>");

                sb.AppendLine("MIME-Version: 1.0");

                // --- BODY 区 ---
                string htmlBody = null;
                string body = null;
                try { htmlBody = mailItem.HTMLBody; } catch { }
                try { body = mailItem.Body; } catch { }

                if (hasAttachments)
                {
                    // 有附件，使用 multipart/mixed
                    sb.AppendLine($"Content-Type: multipart/mixed; boundary=\"{boundary}\"");
                    sb.AppendLine();
                    sb.AppendLine($"--{boundary}");
                }

                if (!string.IsNullOrEmpty(htmlBody))
                {
                    sb.AppendLine("Content-Type: text/html; charset=utf-8");
                    sb.AppendLine("Content-Transfer-Encoding: base64");
                    sb.AppendLine();

                    string base64Body = Convert.ToBase64String(Encoding.UTF8.GetBytes(htmlBody));
                    for (int i = 0; i < base64Body.Length; i += 76)
                        sb.AppendLine(base64Body.Substring(i, Math.Min(76, base64Body.Length - i)));
                }
                else if (!string.IsNullOrEmpty(body))
                {
                    sb.AppendLine("Content-Type: text/plain; charset=utf-8");
                    sb.AppendLine("Content-Transfer-Encoding: 8bit");
                    sb.AppendLine();
                    sb.Append(body);
                }

                // --- 附件区 ---
                if (hasAttachments)
                {
                    for (int i = 1; i <= attachments.Count; i++)
                    {
                        var att = attachments[i];
                        try
                        {
                            string fileName = att.FileName;
                            if (string.IsNullOrEmpty(fileName)) continue;

                            sb.AppendLine();
                            sb.AppendLine($"--{boundary}");
                            sb.AppendLine($"Content-Type: application/octet-stream; name=\"{EncodeHeaderValue(fileName)}\"");
                            sb.AppendLine("Content-Transfer-Encoding: base64");
                            sb.AppendLine($"Content-Disposition: attachment; filename=\"{EncodeHeaderValue(fileName)}\"");
                            sb.AppendLine();

                            // 获取附件内容
                            string attContent = "";
                            if (att.Type == Outlook.OlAttachmentType.olEmbeddeditem)
                            {
                                // 嵌入邮件
                                var embeddedItem = att.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37010102");
                                // 处理嵌入邮件内容
                            }
                            else
                            {
                                // 普通附件 - 使用临时文件读取
                                string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + "_" + fileName);
                                att.SaveAsFile(tempPath);
                                byte[] fileBytes = File.ReadAllBytes(tempPath);
                                File.Delete(tempPath);

                                string base64Att = Convert.ToBase64String(fileBytes);
                                for (int j = 0; j < base64Att.Length; j += 76)
                                    sb.AppendLine(base64Att.Substring(j, Math.Min(76, base64Att.Length - j)));
                            }

                            Log.Information("添加附件: {Name}", fileName);
                        }
                        catch (Exception ex)
                        {
                            Log.Warning("处理附件失败: {Msg}", ex.Message);
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(att);
                        }
                    }

                    sb.AppendLine();
                    sb.AppendLine($"--{boundary}--");
                }

                // 写入文件 (EML 建议不带 UTF-8 BOM)
                File.WriteAllText(emlPath, sb.ToString(), new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                Log.Error("保存邮件失败: {Error}", ex.Message);
            }
        }

        // RFC 2047 编码
        private string EncodeHeaderValue(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";

            // 如果全是 ASCII，直接返回
            bool isAscii = true;
            foreach (char c in value)
            {
                if (c > 127) { isAscii = false; break; }
            }
            if (isAscii) return value;

            // 否则进行 Base64 编码 (RFC 2047)
            return $"=?utf-8?B?{Convert.ToBase64String(Encoding.UTF8.GetBytes(value))}?=";
        }

        // RFC 2822 日期格式
        private string GetRfc2822Date(Outlook.MailItem mailItem)
        {
            DateTime dt = DateTime.Now;
            try { dt = mailItem.SentOn; } catch { }
            if (dt.Year < 1990 || dt.Year > 2100)
            {
                try { dt = mailItem.CreationTime; } catch { }
            }
            if (dt.Year < 1990 || dt.Year > 2100)
            {
                dt = DateTime.Now;
            }

            if (dt.Year < 1900) dt = DateTime.Now;

            // 使用固定时区 +0800
            return dt.ToString("ddd, dd MMM yyyy HH:mm:ss +0800", System.Globalization.CultureInfo.InvariantCulture);
        }

        // 安全获取属性
        private string GetProperty(Outlook.MailItem mailItem, string propName)
        {
            try
            {
                return mailItem.GetType().GetProperty(propName)?.GetValue(mailItem, null) as string ?? "";
            }
            catch
            {
                return "";
            }
        }

        // 简单获取发件人，和Python方案一样
        private string GetFromLine(Outlook.MailItem mailItem)
        {
            try
            {
                string addr = "";
                string name = "";

                try { addr = mailItem.SenderEmailAddress; } catch { }
                try { name = mailItem.SenderName; } catch { }

                if (string.IsNullOrEmpty(addr))
                    addr = "unknown@unknown.com";
                if (string.IsNullOrEmpty(name))
                    name = addr;

                return $"{name} <{addr}>";
            }
            catch
            {
                return "Unknown <unknown@unknown.com>";
            }
        }

        private string CleanFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "Unnamed";
            foreach (char c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
            return name.Length > 100 ? name.Substring(0, 100) : name;
        }

        // 联系人 facade
        public bool ExtractContactsToVcf(string pstPath, string outputDir, IProgress<int> progress = null)
        {
            return ContactSyncService.ExtractContactsToVcf(pstPath, outputDir, progress);
        }

        public List<ContactData> ExtractContactsFromPst(string pstPath, IProgress<int> progress = null)
        {
            return ContactSyncService.ExtractContactsFromPst(pstPath, progress);
        }

        // 日历 facade
        public bool ExtractCalendarToIcs(string pstPath, string outputDir, IProgress<int> progress = null)
        {
            return CalendarSyncService.ExtractCalendarToIcs(pstPath, outputDir, progress);
        }
    }
}
