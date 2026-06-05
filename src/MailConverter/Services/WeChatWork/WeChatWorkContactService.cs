using System;
using System.Collections.Generic;
using System.Linq;
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

        /// <summary>
        /// 拉取所有部门下的活跃成员 (status==1)
        /// 流程: /department/list 拿根部门 → 递归拿子部门 →
        ///       每个部门 /user/simplelist 拿 userid 列表 →
        ///       /user/get 拿每个 user 完整信息
        /// progress 回调: 前半为部门扫描, 后半为用户详情拉取
        /// </summary>
        public List<UserDetailResponse> GetAllMembers(string accessToken, Action<int, int> progress = null)
        {
            if (string.IsNullOrWhiteSpace(accessToken))
                throw new ArgumentException("access_token 不能为空", nameof(accessToken));

            // 1. 拉根部门列表
            var departments = FetchAllDepartments(accessToken);
            if (departments.Count == 0)
            {
                Log.Warning("企业微信未返回任何部门");
                return new List<UserDetailResponse>();
            }
            Log.Information("企业微信部门总数: {Count}", departments.Count);

            // 2. 收集每个部门的 userid (去重)
            var userIdSet = new HashSet<string>();
            var totalDepts = departments.Count;
            for (int i = 0; i < totalDepts; i++)
            {
                var dept = departments[i];
                FetchUserIdsInDepartment(accessToken, dept.Id, userIdSet);
                progress?.Invoke(i + 1, totalDepts * 2);
            }

            Log.Information("企业微信去重后用户数: {Count}", userIdSet.Count);

            // 3. 逐个 user 拉详情
            var result = new List<UserDetailResponse>();
            var allUserIds = userIdSet.ToList();
            var totalUsers = allUserIds.Count;
            for (int i = 0; i < totalUsers; i++)
            {
                var uid = allUserIds[i];
                try
                {
                    var detail = FetchUserDetail(accessToken, uid);
                    if (detail != null && detail.Status == 1)  // 仅激活用户
                    {
                        result.Add(detail);
                    }
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "拉取用户详情失败: userid={Uid}", uid);
                }
                // 每 10 个回调一次进度 (避免 UI 风暴)
                if ((i + 1) % 10 == 0 || i + 1 == totalUsers)
                {
                    progress?.Invoke(totalDepts + i + 1, totalDepts + totalUsers);
                }
            }

            Log.Information("企业微信拉取完成: 共 {Total} 个用户 (已激活)", result.Count);
            return result;
        }

        /// <summary>递归拿所有部门 (含子部门), BFS 防止深度递归栈溢出</summary>
        private List<DepartmentInfo> FetchAllDepartments(string accessToken)
        {
            var all = new List<DepartmentInfo>();
            var queue = new Queue<int>();
            queue.Enqueue(1);  // 根部门 ID = 1

            while (queue.Count > 0)
            {
                var deptId = queue.Dequeue();
                var url = $"{_apiBase}/department/list?access_token={accessToken}&id={deptId}";
                try
                {
                    var json = _http.GetStringAsync(url).GetAwaiter().GetResult();
                    var resp = JsonSerializer.Deserialize<DepartmentListResponse>(json);
                    if (resp == null || resp.ErrCode != 0)
                    {
                        Log.Warning("拉取部门失败: deptId={DeptId}, errcode={Errcode}, errmsg={Errmsg}",
                            deptId, resp?.ErrCode, resp?.ErrMsg);
                        continue;
                    }
                    foreach (var d in resp.Department)
                    {
                        if (!all.Any(x => x.Id == d.Id))
                        {
                            all.Add(d);
                            // 子部门继续 BFS (id != 自身时)
                            if (d.Id != deptId) queue.Enqueue(d.Id);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "拉取部门网络异常: deptId={DeptId}", deptId);
                }
            }
            return all;
        }

        /// <summary>拉取某部门下所有 userid, 含子部门, 累加到 userIdSet</summary>
        private void FetchUserIdsInDepartment(string accessToken, int departmentId, HashSet<string> userIdSet)
        {
            var url = $"{_apiBase}/user/simplelist?access_token={accessToken}&department_id={departmentId}&fetch_child=1";
            try
            {
                var json = _http.GetStringAsync(url).GetAwaiter().GetResult();
                var resp = JsonSerializer.Deserialize<UserSimpleListResponse>(json);
                if (resp == null || resp.ErrCode != 0)
                {
                    Log.Warning("拉取部门用户失败: deptId={DeptId}, errcode={Errcode}, errmsg={Errmsg}",
                        departmentId, resp?.ErrCode, resp?.ErrMsg);
                    return;
                }
                foreach (var u in resp.UserList)
                {
                    if (!string.IsNullOrEmpty(u.UserId))
                        userIdSet.Add(u.UserId);
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "拉取部门用户网络异常: deptId={DeptId}", departmentId);
            }
        }

        /// <summary>拉取单个用户的完整信息</summary>
        private UserDetailResponse FetchUserDetail(string accessToken, string userId)
        {
            var url = $"{_apiBase}/user/get?access_token={accessToken}&userid={Uri.EscapeDataString(userId)}";
            var json = _http.GetStringAsync(url).GetAwaiter().GetResult();
            var resp = JsonSerializer.Deserialize<UserDetailResponse>(json);
            if (resp == null || resp.ErrCode != 0)
            {
                Log.Warning("拉取用户详情失败: userid={Uid}, errcode={Errcode}, errmsg={Errmsg}",
                    userId, resp?.ErrCode, resp?.ErrMsg);
                return null;
            }
            return resp;
        }
    }

    public class WeChatWorkApiException : Exception
    {
        public int ErrCode { get; }
        public WeChatWorkApiException(int errCode, string message) : base(message) { ErrCode = errCode; }
        public WeChatWorkApiException(int errCode, string message, Exception inner) : base(message, inner) { ErrCode = errCode; }
    }
}
