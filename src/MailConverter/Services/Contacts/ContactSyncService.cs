using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Graph;
using Outlook = Microsoft.Office.Interop.Outlook;
using Serilog;

namespace MailConverter.Services.Contacts
{
    /// <summary>
    /// 联系人同步服务: 整合 PST 联系人提取 (VCF / 内存) + Microsoft Graph 批量导入
    /// </summary>
    public static class ContactSyncService
    {
        /// <summary>
        /// 从PST文件中提取联系人到VCF格式
        /// </summary>
        public static bool ExtractContactsToVcf(string pstPath, string outputDir, IProgress<int> progress = null)
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
        /// 从PST文件中提取联系人到内存 (跳过VCF文件，直接用于Graph同步)
        /// </summary>
        public static List<ContactData> ExtractContactsFromPst(string pstPath, IProgress<int> progress = null)
        {
            Program.BatchToO365Logger.Information("开始提取 PST 联系人(直接模式): {PstPath}", pstPath);

            var contacts = new List<ContactData>();

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

        private static void ExtractContactsRecursiveToMemory(Outlook.Folder folder, List<ContactData> contacts, IProgress<int> progress)
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
                                var contactData = new ContactData
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

        private static void ExtractContactsRecursive(Outlook.Folder folder, string outputDir, ref int count, IProgress<int> progress)
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

        private static string CleanFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "Unnamed";
            foreach (char c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
            return name.Length > 100 ? name.Substring(0, 100) : name;
        }

        /// <summary>
        /// 批量导入PST联系人数据到目标邮箱 (直接模式，跳过VCF文件)
        /// 使用 Client Secret + Graph Batch API，并发控制 + 指数退避重试
        /// </summary>
        /// <param name="graphClient">Graph 客户端</param>
        /// <param name="targetEmail">目标用户邮箱</param>
        /// <param name="contacts">PST联系人数据列表</param>
        /// <param name="progressCallback">进度回调 (当前索引, 总数, 状态消息)</param>
        /// <param name="maxDegreeOfParallelism">最大并发数，默认10</param>
        /// <returns>成功导入数量</returns>
        public static async Task<int> ImportContactsBatchDirectAsync(
            GraphServiceClient graphClient,
            string targetEmail,
            IEnumerable<ContactData> contacts,
            Action<int, int, string> progressCallback = null,
            int maxDegreeOfParallelism = 10)
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

            // 过滤出有邮箱的联系人
            var contactList = contacts
                .Where(c => !string.IsNullOrWhiteSpace(c.Email))
                .ToList();

            int totalCount = contactList.Count;
            int successCount = 0;
            int skippedCount = contacts.Count() - totalCount;

            Program.BatchToO365Logger.Information("筛选后有邮箱联系人: {Valid}/{Total}, 跳过无邮箱: {Skipped}",
                totalCount, contacts.Count(), skippedCount);

            if (totalCount == 0)
            {
                Program.BatchToO365Logger.Warning("没有有邮箱的联系人需要导入");
                return 0;
            }

            int currentIndex = 0;

            Program.BatchToO365Logger.Information("开始直接批量导入联系人到 {Email}, 共 {Count} 个, 并发数: {Parallelism}",
                targetEmail, totalCount, maxDegreeOfParallelism);

            // 使用 SemaphoreSlim 控制并发数
            var semaphore = new SemaphoreSlim(maxDegreeOfParallelism);
            var lockObj = new object();

            // 并行处理所有联系人
            var tasks = contactList.Select(async contactData =>
            {
                await semaphore.WaitAsync();
                try
                {
                    int index;
                    lock (lockObj)
                    {
                        index = currentIndex++;
                    }

                    // 带重试的导入
                    bool success = await ImportSingleContactDirectWithRetryAsync(graphClient, targetEmail, contactData);
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

            Program.BatchToO365Logger.Information("直接批量导入完成: 成功 {Success}/{Total}", successCount, totalCount);
            return successCount;
        }

        /// <summary>
        /// 带重试的单联系人直接导入 (指数退避)
        /// </summary>
        private static async Task<bool> ImportSingleContactDirectWithRetryAsync(GraphServiceClient graphClient, string targetEmail, ContactData contactData, int maxRetries = 3)
        {
            for (int retry = 0; retry <= maxRetries; retry++)
            {
                try
                {
                    // 将PST联系人数据转换为Graph Contact
                    var contact = ConvertToGraphContact(contactData);

                    // 调用Graph API创建联系人
                    await graphClient.Users[targetEmail].Contacts.PostAsync(contact);

                    Program.BatchToO365Logger.Information("联系人创建成功: {Name}", contact.DisplayName ?? "Unknown");
                    return true;
                }
                catch (Exception ex)
                {
                    bool isThrottled = IsThrottledException(ex);

                    if (isThrottled && retry < maxRetries)
                    {
                        int delaySeconds = (int)Math.Pow(2, retry);
                        Program.BatchToO365Logger.Warning("限流 (429)，{Delay} 秒后重试 (第 {Retry}/{Max} 次): {Name}",
                            delaySeconds, retry + 1, maxRetries, contactData.DisplayName);
                        await Task.Delay(delaySeconds * 1000);
                    }
                    else
                    {
                        Program.BatchToO365Logger.Error(ex, "导入联系人失败 (重试 {Retry}/{Max}): {Name}",
                            retry, maxRetries, contactData.DisplayName);
                        return false;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// 将PST联系人数据转换为Graph Contact模型
        /// </summary>
        public static Microsoft.Graph.Models.Contact ConvertToGraphContact(ContactData data)
        {
            var contact = new Microsoft.Graph.Models.Contact();

            // 姓名
            if (!string.IsNullOrWhiteSpace(data.DisplayName))
                contact.DisplayName = data.DisplayName;
            else if (!string.IsNullOrWhiteSpace(data.Email))
                contact.DisplayName = data.Email;
            else if (!string.IsNullOrWhiteSpace(data.CompanyName))
                contact.DisplayName = data.CompanyName;
            else
                contact.DisplayName = "Unknown Contact";

            contact.GivenName = data.FirstName ?? "";
            contact.Surname = data.LastName ?? "";
            contact.MiddleName = data.MiddleName ?? "";
            contact.Title = data.Title ?? "";

            // 公司信息
            contact.CompanyName = data.CompanyName ?? "";
            contact.Department = data.Department ?? "";
            contact.JobTitle = data.JobTitle ?? "";

            // 备注 - 包含邮箱和电话信息
            var notes = new System.Text.StringBuilder();
            if (!string.IsNullOrWhiteSpace(data.PersonalNotes))
                notes.AppendLine(data.PersonalNotes);
            if (!string.IsNullOrWhiteSpace(data.Email))
                notes.AppendLine($"Email: {data.Email}");
            if (!string.IsNullOrWhiteSpace(data.Phone))
                notes.AppendLine($"Phone: {data.Phone}");
            if (!string.IsNullOrWhiteSpace(data.MobilePhone))
                notes.AppendLine($"Mobile: {data.MobilePhone}");
            if (notes.Length > 0)
                contact.PersonalNotes = notes.ToString().TrimEnd();

            // 生日
            if (data.Birthday.HasValue)
                contact.Birthday = data.Birthday.Value;

            return contact;
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
    }
}
