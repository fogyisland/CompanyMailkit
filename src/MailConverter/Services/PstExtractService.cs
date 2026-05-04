using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Serilog;
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
                ExtractFolder(targetFolder, outputDir, ref totalExtracted, progress);

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

        private void ExtractFolder(Outlook.Folder folder, string basePath, ref int totalCount, IProgress<int> progress)
        {
            string folderName = folder.Name;
            if (folderName == "Outlook Data File")
                folderName = "";

            string folderPath = string.IsNullOrEmpty(folderName) ? basePath : Path.Combine(basePath, CleanFileName(folderName));

            if (!string.IsNullOrEmpty(folderName) && folderName != "Outlook Data File")
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

        /// <summary>
        /// 从PST文件中提取联系人到VCF格式
        /// </summary>
        public bool ExtractContactsToVcf(string pstPath, string outputDir, IProgress<int> progress = null)
        {
            Log.Information("开始提取 PST 联系人: {PstPath}", pstPath);

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

                // 打开PST文件
                var pstFolder = ns.Folders.Add(pstPath, Type.Missing) as Outlook.Folder;
                if (pstFolder == null)
                {
                    Log.Error("无法打开 PST 文件: {Path}", pstPath);
                    return false;
                }

                // 遍历文件夹查找联系人
                int contactCount = 0;
                ExtractContactsRecursive(pstFolder, outputDir, ref contactCount, progress);

                Log.Information("联系人提取完成，共 {Count} 个", contactCount);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "提取 PST 联系人失败: {Path}", pstPath);
                return false;
            }
            finally
            {
                try { if (ns != null) Marshal.ReleaseComObject(ns); } catch { }
                try { if (outlookApp != null) Marshal.ReleaseComObject(outlookApp); } catch { }
            }
        }

        /// <summary>
        /// PST联系人数据类 (用于直接同步到Graph)
        /// </summary>
        public class PstContactData
        {
            public string DisplayName { get; set; }
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public string MiddleName { get; set; }
            public string Title { get; set; }
            public string Suffix { get; set; }
            public string Email { get; set; }
            public string Email2 { get; set; }
            public string Email3 { get; set; }
            public string Phone { get; set; }
            public string Phone2 { get; set; }
            public string MobilePhone { get; set; }
            public string CompanyName { get; set; }
            public string Department { get; set; }
            public string JobTitle { get; set; }
            public string BusinessAddress { get; set; }
            public string HomeAddress { get; set; }
            public string PersonalNotes { get; set; }
            public DateTime? Birthday { get; set; }
        }

        /// <summary>
        /// PST日历数据类 (用于直接同步到Graph)
        /// </summary>
        public class PstCalendarData
        {
            public string Subject { get; set; }
            public string Body { get; set; }
            public DateTime? StartTime { get; set; }
            public DateTime? EndTime { get; set; }
            public string Location { get; set; }
            public bool IsAllDayEvent { get; set; }
            public string ReminderMinutesBeforeStart { get; set; }
            public bool ReminderSet { get; set; }
            public string Categories { get; set; }
            public string RequiredAttendees { get; set; }
            public string OptionalAttendees { get; set; }
            public string ResourceAttendees { get; set; }
            public bool IsRecurring { get; set; }
            public string RecurrencePattern { get; set; }
        }

        /// <summary>
        /// 从PST文件中提取联系人到内存 (跳过VCF文件，直接用于Graph同步)
        /// </summary>
        public List<PstContactData> ExtractContactsFromPst(string pstPath, IProgress<int> progress = null)
        {
            Program.BatchToO365Logger.Information("开始提取 PST 联系人(直接模式): {PstPath}", pstPath);

            var contacts = new List<PstContactData>();

            if (!File.Exists(pstPath))
            {
                Program.BatchToO365Logger.Error("PST 文件不存在: {Path}", pstPath);
                return contacts;
            }

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
                Outlook.Folder pstFolder = null;

                // 使用 AddStoreEx 挂载 PST 文件
                try
                {
                    ns.AddStoreEx(pstPath, Outlook.OlStoreType.olStoreUnicode);
                }
                catch (Exception ex)
                {
                    Program.BatchToO365Logger.Warning("添加 Store 失败，可能已挂载: {Msg}", ex.Message);
                }

                // 寻找对应的 Folder
                foreach (Outlook.Folder folder in ns.Folders)
                {
                    try
                    {
                        if (folder.Store != null && !string.IsNullOrEmpty(folder.Store.FilePath))
                        {
                            if (Path.GetFullPath(folder.Store.FilePath).Equals(Path.GetFullPath(pstPath), StringComparison.OrdinalIgnoreCase))
                            {
                                pstFolder = folder;
                                break;
                            }
                        }
                    }
                    catch { }
                }

                if (pstFolder == null)
                {
                    Program.BatchToO365Logger.Error("无法找到 PST 文件对应的 Folder: {Path}", pstPath);
                    return contacts;
                }

                // 遍历文件夹提取联系人
                ExtractContactsRecursiveToMemory(pstFolder, contacts, progress);

                Program.BatchToO365Logger.Information("联系人提取完成，共 {Count} 个", contacts.Count);
                return contacts;
            }
            catch (Exception ex)
            {
                Program.BatchToO365Logger.Error(ex, "提取 PST 联系人失败: {Path}", pstPath);
                return contacts;
            }
            finally
            {
                try { if (ns != null) Marshal.ReleaseComObject(ns); } catch { }
                try { if (outlookApp != null) Marshal.ReleaseComObject(outlookApp); } catch { }
            }
        }

        private void ExtractContactsRecursiveToMemory(Outlook.Folder folder, List<PstContactData> contacts, IProgress<int> progress)
        {
            try
            {
                Outlook.Items items = folder.Items;
                for (int i = 1; i <= items.Count; i++)
                {
                    object item = null;
                    try
                    {
                        item = items[i];
                        if (item is Outlook.ContactItem contact)
                        {
                            try
                            {
                                var contactData = new PstContactData
                                {
                                    DisplayName = contact.FullName ?? "",
                                    FirstName = contact.FirstName ?? "",
                                    LastName = contact.LastName ?? "",
                                    MiddleName = contact.MiddleName ?? "",
                                    Title = contact.Title ?? "",
                                    Suffix = contact.Suffix ?? "",
                                    Email = contact.Email1Address ?? "",
                                    Email2 = contact.Email2Address ?? "",
                                    Email3 = contact.Email3Address ?? "",
                                    Phone = contact.PrimaryTelephoneNumber ?? "",
                                    Phone2 = contact.BusinessTelephoneNumber ?? "",
                                    MobilePhone = contact.MobileTelephoneNumber ?? "",
                                    CompanyName = contact.CompanyName ?? "",
                                    Department = contact.Department ?? "",
                                    JobTitle = contact.JobTitle ?? "",
                                    Birthday = contact.Birthday,
                                    PersonalNotes = contact.Body ?? ""
                                };

                                // 处理地址
                                if (contact.BusinessAddress != null)
                                    contactData.BusinessAddress = contact.BusinessAddress;

                                contacts.Add(contactData);
                                progress?.Report(contacts.Count);
                            }
                            catch (Exception ex)
                            {
                                Program.BatchToO365Logger.Warning(ex, "提取联系人失败: {Name}", contact.FullName);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Program.BatchToO365Logger.Warning(ex, "处理联系人失败");
                    }
                    finally { if (item != null) Marshal.ReleaseComObject(item); }
                }
                Marshal.ReleaseComObject(items);
            }
            catch (Exception ex)
            {
                Program.BatchToO365Logger.Warning(ex, "遍历文件夹失败: {Name}", folder.Name);
            }

            // 递归处理子文件夹
            try
            {
                foreach (Outlook.Folder subFolder in folder.Folders)
                {
                    ExtractContactsRecursiveToMemory(subFolder, contacts, progress);
                    Marshal.ReleaseComObject(subFolder);
                }
            }
            catch { }
        }

        private void ExtractContactsRecursive(Outlook.Folder folder, string outputDir, ref int count, IProgress<int> progress)
        {
            try
            {
                Outlook.Items items = folder.Items;
                for (int i = 1; i <= items.Count; i++)
                {
                    object item = null;
                    try
                    {
                        item = items[i];
                        if (item is Outlook.ContactItem contact)
                        {
                            try
                            {
                                var vcfPath = Path.Combine(outputDir, CleanFileName(contact.FullName) + "_" + count + ".vcf");
                                contact.SaveAs(vcfPath, Outlook.OlSaveAsType.olVCard);
                                count++;
                                progress?.Report(count);
                            }
                            catch (Exception ex)
                            {
                                Log.Warning(ex, "导出联系人失败: {Name}", contact.FullName);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Warning(ex, "处理联系人失败");
                    }
                    finally { if (item != null) Marshal.ReleaseComObject(item); }
                }
                Marshal.ReleaseComObject(items);
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "遍历文件夹失败: {Name}", folder.Name);
            }

            // 递归处理子文件夹
            try
            {
                foreach (Outlook.Folder subFolder in folder.Folders)
                {
                    ExtractContactsRecursive(subFolder, outputDir, ref count, progress);
                    Marshal.ReleaseComObject(subFolder);
                }
            }
            catch { }
        }

        /// <summary>
        /// 从PST文件中提取日历到ICS格式
        /// </summary>
        public bool ExtractCalendarToIcs(string pstPath, string outputDir, IProgress<int> progress = null)
        {
            Log.Information("开始提取 PST 日历: {PstPath}", pstPath);

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

                // 打开PST文件
                var pstFolder = ns.Folders.Add(pstPath, Type.Missing) as Outlook.Folder;
                if (pstFolder == null)
                {
                    Log.Error("无法打开 PST 文件: {Path}", pstPath);
                    return false;
                }

                // 遍历文件夹查找日历项
                int calendarCount = 0;
                ExtractCalendarRecursive(pstFolder, outputDir, ref calendarCount, progress);

                Log.Information("日历提取完成，共 {Count} 个事件", calendarCount);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "提取 PST 日历失败: {Path}", pstPath);
                return false;
            }
            finally
            {
                try { if (ns != null) Marshal.ReleaseComObject(ns); } catch { }
                try { if (outlookApp != null) Marshal.ReleaseComObject(outlookApp); } catch { }
            }
        }

        private void ExtractCalendarRecursive(Outlook.Folder folder, string outputDir, ref int count, IProgress<int> progress)
        {
            try
            {
                Outlook.Items items = folder.Items;
                for (int i = 1; i <= items.Count; i++)
                {
                    object item = null;
                    try
                    {
                        item = items[i];
                        if (item is Outlook.AppointmentItem appointment)
                        {
                            try
                            {
                                var icsPath = Path.Combine(outputDir, CleanFileName(appointment.Subject ?? "日历") + "_" + count + ".ics");
                                SaveAppointmentAsIcs(appointment, icsPath);
                                count++;
                                progress?.Report(count);
                            }
                            catch (Exception ex)
                            {
                                Log.Warning(ex, "导出日历失败: {Subject}", appointment.Subject);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Warning(ex, "处理日历项失败");
                    }
                    finally { if (item != null) Marshal.ReleaseComObject(item); }
                }
                Marshal.ReleaseComObject(items);
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "遍历文件夹失败: {Name}", folder.Name);
            }

            // 递归处理子文件夹
            try
            {
                foreach (Outlook.Folder subFolder in folder.Folders)
                {
                    ExtractCalendarRecursive(subFolder, outputDir, ref count, progress);
                    Marshal.ReleaseComObject(subFolder);
                }
            }
            catch { }
        }

        private void SaveAppointmentAsIcs(Outlook.AppointmentItem appointment, string icsPath)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("BEGIN:VCALENDAR");
                sb.AppendLine("VERSION:2.0");
                sb.AppendLine("PRODID:-//MailConverter//PST Extract//EN");
                sb.AppendLine("BEGIN:VEVENT");

                // UID
                string entryId = "";
                try { entryId = appointment.EntryID; } catch { }
                sb.AppendLine($"UID:{entryId ?? Guid.NewGuid().ToString()}@pst-converter.local");

                // Subject
                string subject = "";
                try { subject = appointment.Subject ?? ""; } catch { }
                sb.AppendLine($"SUMMARY:{EncodeIcsText(subject)}");

                // Description
                string description = "";
                try { description = appointment.Body ?? ""; } catch { }
                sb.AppendLine($"DESCRIPTION:{EncodeIcsText(description)}");

                // Start time
                DateTime start = DateTime.Now;
                try { start = appointment.Start; } catch { }
                sb.AppendLine($"DTSTART:{start.ToString("yyyyMMddTHHmmss")}");

                // End time
                DateTime end = DateTime.Now.AddHours(1);
                try { end = appointment.End; } catch { }
                sb.AppendLine($"DTEND:{end.ToString("yyyyMMddTHHmmss")}");

                // Location
                string location = "";
                try { location = appointment.Location ?? ""; } catch { }
                if (!string.IsNullOrEmpty(location))
                    sb.AppendLine($"LOCATION:{EncodeIcsText(location)}");

                // Organizer
                string organizer = "";
                try { organizer = appointment.Organizer ?? ""; } catch { }
                if (!string.IsNullOrEmpty(organizer))
                    sb.AppendLine($"ORGANIZER;CN={EncodeIcsText(organizer)}:mailto:{organizer}");

                // Creation time
                DateTime created = DateTime.Now;
                try { created = appointment.CreationTime; } catch { }
                sb.AppendLine($"CREATED:{created.ToString("yyyyMMddTHHmmss")}");

                // Last modified
                DateTime modified = DateTime.Now;
                try { modified = appointment.LastModificationTime; } catch { }
                sb.AppendLine($"LAST-MODIFIED:{modified.ToString("yyyyMMddTHHmmss")}");

                sb.AppendLine("END:VEVENT");
                sb.AppendLine("END:VCALENDAR");

                File.WriteAllText(icsPath, sb.ToString(), new UTF8Encoding(false));
                Log.Information("导出日历: {Subject}", subject);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "生成 ICS 文件失败: {Path}", icsPath);
                throw;
            }
        }

        /// <summary>
        /// 从PST文件中提取日历到内存 (跳过ICS文件，直接用于Graph同步)
        /// </summary>
        public List<PstCalendarData> ExtractCalendarFromPst(string pstPath, IProgress<int> progress = null)
        {
            Program.BatchToO365Logger.Information("开始提取 PST 日历(直接模式): {PstPath}", pstPath);

            var calendars = new List<PstCalendarData>();

            if (!File.Exists(pstPath))
            {
                Program.BatchToO365Logger.Error("PST 文件不存在: {Path}", pstPath);
                return calendars;
            }

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
                Outlook.Folder pstFolder = null;

                // 使用 AddStoreEx 挂载 PST 文件
                try
                {
                    ns.AddStoreEx(pstPath, Outlook.OlStoreType.olStoreUnicode);
                }
                catch (Exception ex)
                {
                    Program.BatchToO365Logger.Warning("添加 Store 失败，可能已挂载: {Msg}", ex.Message);
                }

                // 寻找对应的 Folder
                foreach (Outlook.Folder folder in ns.Folders)
                {
                    try
                    {
                        if (folder.Store != null && !string.IsNullOrEmpty(folder.Store.FilePath))
                        {
                            if (Path.GetFullPath(folder.Store.FilePath).Equals(Path.GetFullPath(pstPath), StringComparison.OrdinalIgnoreCase))
                            {
                                pstFolder = folder;
                                break;
                            }
                        }
                    }
                    catch { }
                }

                if (pstFolder == null)
                {
                    Program.BatchToO365Logger.Error("无法找到 PST 文件对应的 Folder: {Path}", pstPath);
                    return calendars;
                }

                // 遍历文件夹提取日历
                ExtractCalendarRecursiveToMemory(pstFolder, calendars, progress);

                Program.BatchToO365Logger.Information("日历提取完成，共 {Count} 个事件", calendars.Count);
                return calendars;
            }
            catch (Exception ex)
            {
                Program.BatchToO365Logger.Error(ex, "提取 PST 日历失败: {Path}", pstPath);
                return calendars;
            }
            finally
            {
                try { if (ns != null) Marshal.ReleaseComObject(ns); } catch { }
                try { if (outlookApp != null) Marshal.ReleaseComObject(outlookApp); } catch { }
            }
        }

        private void ExtractCalendarRecursiveToMemory(Outlook.Folder folder, List<PstCalendarData> calendars, IProgress<int> progress)
        {
            try
            {
                Outlook.Items items = folder.Items;
                for (int i = 1; i <= items.Count; i++)
                {
                    object item = null;
                    try
                    {
                        item = items[i];
                        if (item is Outlook.AppointmentItem appointment)
                        {
                            try
                            {
                                var calData = new PstCalendarData
                                {
                                    Subject = appointment.Subject ?? "",
                                    Body = appointment.Body ?? "",
                                    StartTime = appointment.Start,
                                    EndTime = appointment.End,
                                    Location = appointment.Location ?? "",
                                    IsAllDayEvent = appointment.AllDayEvent,
                                    ReminderSet = appointment.ReminderSet,
                                    Categories = appointment.Categories ?? ""
                                };

                                // 处理与会者信息
                                try
                                {
                                    calData.RequiredAttendees = appointment.RequiredAttendees ?? "";
                                    calData.OptionalAttendees = appointment.OptionalAttendees ?? "";
                                    calData.ResourceAttendees = appointment.Resources ?? "";
                                }
                                catch { }

                                // 检查是否是重复日程
                                try
                                {
                                    calData.IsRecurring = appointment.IsRecurring;
                                }
                                catch { }

                                calendars.Add(calData);
                                progress?.Report(calendars.Count);
                            }
                            catch (Exception ex)
                            {
                                Program.BatchToO365Logger.Warning(ex, "提取日历事件失败: {Subject}", appointment.Subject);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Program.BatchToO365Logger.Warning(ex, "处理日历项失败");
                    }
                    finally { if (item != null) Marshal.ReleaseComObject(item); }
                }
                Marshal.ReleaseComObject(items);
            }
            catch (Exception ex)
            {
                Program.BatchToO365Logger.Warning(ex, "遍历文件夹失败: {Name}", folder.Name);
            }

            // 递归处理子文件夹
            try
            {
                foreach (Outlook.Folder subFolder in folder.Folders)
                {
                    ExtractCalendarRecursiveToMemory(subFolder, calendars, progress);
                    Marshal.ReleaseComObject(subFolder);
                }
            }
            catch { }
        }

        private string EncodeIcsText(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            // Escape special characters for ICS format
            return text.Replace("\\", "\\\\")
                       .Replace(",", "\\,")
                       .Replace(";", "\\;")
                       .Replace("\n", "\\n")
                       .Replace("\r", "");
        }
    }
}
