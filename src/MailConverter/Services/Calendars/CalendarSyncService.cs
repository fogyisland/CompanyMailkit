using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Graph;
using Outlook = Microsoft.Office.Interop.Outlook;
using Serilog;

namespace MailConverter.Services.Calendars
{
    /// <summary>
    /// 日历同步服务: 整合 PST 日历提取 (ICS / 内存) + Microsoft Graph 批量导入
    /// </summary>
    public static class CalendarSyncService
    {
        /// <summary>
        /// 从PST文件中提取日历到ICS格式
        /// </summary>
        public static bool ExtractCalendarToIcs(string pstPath, string outputDir, IProgress<int> progress = null)
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

        private static void ExtractCalendarRecursive(Outlook.Folder folder, string outputDir, ref int count, IProgress<int> progress)
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

        private static void SaveAppointmentAsIcs(Outlook.AppointmentItem appointment, string icsPath)
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

        private static string EncodeIcsText(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            // Escape special characters for ICS format
            return text.Replace("\\", "\\\\")
                       .Replace(",", "\\,")
                       .Replace(";", "\\;")
                       .Replace("\n", "\\n")
                       .Replace("\r", "");
        }

        private static string CleanFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "Unnamed";
            foreach (char c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
            return name.Length > 100 ? name.Substring(0, 100) : name;
        }

        /// <summary>
        /// 批量导入PST日历数据到目标邮箱 (直接模式，跳过ICS文件)
        /// 使用 Client Secret + Graph Batch API，并发控制 + 指数退避重试
        /// </summary>
        /// <param name="graphClient">Graph 客户端</param>
        /// <param name="targetEmail">目标用户邮箱</param>
        /// <param name="calendars">PST日历数据列表</param>
        /// <param name="progressCallback">进度回调 (当前索引, 总数, 状态消息)</param>
        /// <param name="maxDegreeOfParallelism">最大并发数，默认10</param>
        /// <param name="timeZone">时区，默认 "China Standard Time"</param>
        /// <returns>成功导入数量</returns>
        public static async Task<int> ImportCalendarBatchDirectAsync(
            GraphServiceClient graphClient,
            string targetEmail,
            IEnumerable<CalendarData> calendars,
            Action<int, int, string> progressCallback = null,
            int maxDegreeOfParallelism = 10,
            string timeZone = "China Standard Time")
        {
            if (graphClient == null)
            {
                Program.BatchToO365Logger.Error("Graph客户端未初始化，请先调用 ConnectWithClientSecret");
                return 0;
            }

            if (string.IsNullOrWhiteSpace(targetEmail))
            {
                Program.BatchToO365Logger.Error("目标邮箱不能为空");
                return 0;
            }

            // 过滤出有主题和时间的日历事件
            var calendarList = calendars
                .Where(c => !string.IsNullOrWhiteSpace(c.Subject) && c.StartTime.HasValue)
                .ToList();

            int totalCount = calendarList.Count;
            int skippedCount = calendars.Count() - totalCount;

            Program.BatchToO365Logger.Information("筛选后有效日历: {Valid}/{Total}, 跳过无效: {Skipped}",
                totalCount, calendars.Count(), skippedCount);

            if (totalCount == 0)
            {
                Program.BatchToO365Logger.Warning("没有有效的日历事件需要导入");
                return 0;
            }

            int successCount = 0;
            int currentIndex = 0;

            Program.BatchToO365Logger.Information("开始直接批量导入日历到 {Email}, 共 {Count} 个, 并发数: {Parallelism}",
                targetEmail, totalCount, maxDegreeOfParallelism);

            var semaphore = new SemaphoreSlim(maxDegreeOfParallelism);
            var lockObj = new object();

            var tasks = calendarList.Select(async calendarData =>
            {
                await semaphore.WaitAsync();
                try
                {
                    int index;
                    lock (lockObj)
                    {
                        index = currentIndex++;
                    }

                    bool success = await ImportSingleCalendarDirectWithRetryAsync(graphClient, targetEmail, calendarData, timeZone);
                    if (success)
                    {
                        lock (lockObj) { successCount++; }
                    }

                    progressCallback?.Invoke(index + 1, totalCount, $"已导入 {index + 1}/{totalCount}");
                    return success;
                }
                finally
                {
                    semaphore.Release();
                }
            });

            await Task.WhenAll(tasks);

            Program.BatchToO365Logger.Information("直接批量导入日历完成: 成功 {Success}/{Total}", successCount, totalCount);
            return successCount;
        }

        private static async Task<bool> ImportSingleCalendarDirectWithRetryAsync(GraphServiceClient graphClient, string targetEmail, CalendarData calendarData, string timeZone, int maxRetries = 3)
        {
            for (int retry = 0; retry <= maxRetries; retry++)
            {
                try
                {
                    var evt = ConvertToGraphEvent(calendarData, timeZone);
                    await graphClient.Users[targetEmail].Calendar.Events.PostAsync(evt);

                    Program.BatchToO365Logger.Information("日历创建成功: {Subject}", evt.Subject ?? "Unknown");
                    return true;
                }
                catch (Exception ex)
                {
                    bool isThrottled = IsThrottledException(ex);

                    if (isThrottled && retry < maxRetries)
                    {
                        int delaySeconds = (int)Math.Pow(2, retry);
                        Program.BatchToO365Logger.Warning("限流 (429)，{Delay} 秒后重试 (第 {Retry}/{Max} 次): {Subject}",
                            delaySeconds, retry + 1, maxRetries, calendarData.Subject);
                        await Task.Delay(delaySeconds * 1000);
                    }
                    else
                    {
                        Program.BatchToO365Logger.Error(ex, "导入日历失败 (重试 {Retry}/{Max}): {Subject}",
                            retry, maxRetries, calendarData.Subject);
                        return false;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// 将 UTC 偏移格式 (如 "UTC+8") 转换为 IANA 时区名称 (如 "China Standard Time")
        /// Graph API DateTimeTimeZone.TimeZone 需要使用 IANA 时区名称
        /// </summary>
        private static string ConvertToIanaTimeZone(string utcOffset)
        {
            if (string.IsNullOrWhiteSpace(utcOffset)) return "UTC";

            // UTC 偏移格式映射到 IANA 时区
            var offsetToTimeZone = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "UTC-12", "Dateline Standard Time" },
                { "UTC-11", "UTC-11" },
                { "UTC-10", "Hawaiian Standard Time" },
                { "UTC-9", "Alaskan Standard Time" },
                { "UTC-8", "Pacific Standard Time" },
                { "UTC-7", "Mountain Standard Time" },
                { "UTC-6", "Central Standard Time" },
                { "UTC-5", "Eastern Standard Time" },
                { "UTC-4", "Atlantic Standard Time" },
                { "UTC-3", "SA Eastern Standard Time" },
                { "UTC-2", "Mid-Atlantic Standard Time" },
                { "UTC-1", "Azores Standard Time" },
                { "UTC", "UTC" },
                { "UTC+1", "Romance Standard Time" },
                { "UTC+2", "Egypt Standard Time" },
                { "UTC+3", "Russian Standard Time" },
                { "UTC+4", "Arabian Standard Time" },
                { "UTC+5", "West Asia Standard Time" },
                { "UTC+5:30", "India Standard Time" },
                { "UTC+6", "Central Asia Standard Time" },
                { "UTC+6:30", "Myanmar Standard Time" },
                { "UTC+7", "SE Asia Standard Time" },
                { "UTC+8", "China Standard Time" },  // 中国标准时间 (北京)
                { "UTC+9", "Tokyo Standard Time" },
                { "UTC+9:30", "AUS Central Standard Time" },
                { "UTC+10", "AUS Eastern Standard Time" },
                { "UTC+11", "Central Pacific Standard Time" },
                { "UTC+12", "New Zealand Standard Time" }
            };

            // 尝试直接匹配
            if (offsetToTimeZone.TryGetValue(utcOffset, out string ianaZone))
            {
                Program.BatchToO365Logger.Debug("时区转换: {Input} -> {Output}", utcOffset, ianaZone);
                return ianaZone;
            }

            // 如果匹配不到，返回 UTC
            Program.BatchToO365Logger.Warning("未知的时区格式: {Input}, 使用 UTC", utcOffset);
            return "UTC";
        }

        /// <summary>
        /// 将PST日历数据转换为Graph Event模型
        /// </summary>
        public static Microsoft.Graph.Models.Event ConvertToGraphEvent(CalendarData data, string timeZone = "China Standard Time")
        {
            var evt = new Microsoft.Graph.Models.Event();

            evt.Subject = data.Subject ?? "";
            evt.Body = new Microsoft.Graph.Models.ItemBody
            {
                ContentType = Microsoft.Graph.Models.BodyType.Text,
                Content = data.Body ?? ""
            };
            evt.Location = new Microsoft.Graph.Models.Location
            {
                DisplayName = data.Location ?? ""
            };

            // 将 UTC 偏移格式转换为 IANA 时区名称 (Graph API 需要)
            string ianaTimeZone = ConvertToIanaTimeZone(timeZone);

            if (data.StartTime.HasValue)
            {
                var startTime = data.StartTime.Value;
                evt.Start = new Microsoft.Graph.Models.DateTimeTimeZone
                {
                    DateTime = startTime.ToString("yyyy-MM-ddTHH:mm:ss"),
                    TimeZone = ianaTimeZone
                };
            }

            if (data.EndTime.HasValue)
            {
                var endTime = data.EndTime.Value;
                evt.End = new Microsoft.Graph.Models.DateTimeTimeZone
                {
                    DateTime = endTime.ToString("yyyy-MM-ddTHH:mm:ss"),
                    TimeZone = ianaTimeZone
                };
            }

            evt.IsAllDay = data.IsAllDayEvent;

            // 防止向与会者发送通知邮件 (静默迁移)
            evt.ResponseRequested = false;

            if (data.ReminderSet && !string.IsNullOrWhiteSpace(data.ReminderMinutesBeforeStart))
            {
                evt.ReminderMinutesBeforeStart = int.TryParse(data.ReminderMinutesBeforeStart, out int minutes) ? minutes : 30;
                evt.IsReminderOn = true;
            }

            // 处理与会者
            if (!string.IsNullOrWhiteSpace(data.RequiredAttendees))
            {
                evt.Attendees = new List<Microsoft.Graph.Models.Attendee>();
                var requiredList = data.RequiredAttendees.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var attendee in requiredList)
                {
                    var email = attendee.Trim();
                    if (!string.IsNullOrWhiteSpace(email) && email.Contains("@"))
                    {
                        evt.Attendees.Add(new Microsoft.Graph.Models.Attendee
                        {
                            EmailAddress = new Microsoft.Graph.Models.EmailAddress { Address = email },
                            Type = Microsoft.Graph.Models.AttendeeType.Required
                        });
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(data.OptionalAttendees))
            {
                if (evt.Attendees == null) evt.Attendees = new List<Microsoft.Graph.Models.Attendee>();
                var optionalList = data.OptionalAttendees.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var attendee in optionalList)
                {
                    var email = attendee.Trim();
                    if (!string.IsNullOrWhiteSpace(email) && email.Contains("@"))
                    {
                        evt.Attendees.Add(new Microsoft.Graph.Models.Attendee
                        {
                            EmailAddress = new Microsoft.Graph.Models.EmailAddress { Address = email },
                            Type = Microsoft.Graph.Models.AttendeeType.Optional
                        });
                    }
                }
            }

            return evt;
        }

        /// <summary>
        /// 检查异常是否为限流 (429)
        /// </summary>
        private static bool IsThrottledException(Exception ex)
        {
            // 检查异常消息中是否包含429
            var exMsg = ex.Message.ToLower();

            if (exMsg.Contains("429") ||
                exMsg.Contains("throttled") ||
                exMsg.Contains("too many requests") ||
                exMsg.Contains("service unavailable"))
            {
                return true;
            }

            // 递归检查内部异常
            if (ex.InnerException != null)
            {
                return IsThrottledException(ex.InnerException);
            }

            return false;
        }

        /// <summary>
        /// 从PST文件中提取日历到内存 (跳过ICS文件，直接用于Graph同步)
        /// </summary>
        public static List<CalendarData> ExtractCalendarFromPst(string pstPath, IProgress<int> progress = null)
        {
            Program.BatchToO365Logger.Information("开始提取 PST 日历(直接模式): {PstPath}", pstPath);

            var calendars = new List<CalendarData>();

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

                Program.BatchToO365Logger.Information("日历提取完成，共 {Count} 个", calendars.Count);
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

        private static void ExtractCalendarRecursiveToMemory(Outlook.Folder folder, List<CalendarData> calendars, IProgress<int> progress)
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
                                var calData = new CalendarData
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

                                try
                                {
                                    calData.RequiredAttendees = appointment.RequiredAttendees ?? "";
                                    calData.OptionalAttendees = appointment.OptionalAttendees ?? "";
                                    calData.ResourceAttendees = appointment.Resources ?? "";
                                }
                                catch { }

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
    }
}
