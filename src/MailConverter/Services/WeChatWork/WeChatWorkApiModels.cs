using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace MailConverter.Services.WeChatWork
{
    /// <summary>企业微信 /cgi-bin/gettoken 响应</summary>
    public class GetTokenResponse
    {
        [JsonPropertyName("errcode")]
        public int ErrCode { get; set; }

        [JsonPropertyName("errmsg")]
        public string ErrMsg { get; set; } = "";

        [JsonPropertyName("access_token")]
        public string AccessToken { get; set; } = "";

        [JsonPropertyName("expires_in")]
        public int ExpiresIn { get; set; }
    }

    /// <summary>/cgi-bin/department/list 响应</summary>
    public class DepartmentListResponse
    {
        [JsonPropertyName("errcode")]
        public int ErrCode { get; set; }

        [JsonPropertyName("errmsg")]
        public string ErrMsg { get; set; } = "";

        [JsonPropertyName("department")]
        public List<DepartmentInfo> Department { get; set; } = new List<DepartmentInfo>();
    }

    public class DepartmentInfo
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; } = "";

        [JsonPropertyName("parentid")]
        public int ParentId { get; set; }

        [JsonPropertyName("order")]
        public int Order { get; set; }
    }

    /// <summary>/cgi-bin/user/simplelist 响应</summary>
    public class UserSimpleListResponse
    {
        [JsonPropertyName("errcode")]
        public int ErrCode { get; set; }

        [JsonPropertyName("errmsg")]
        public string ErrMsg { get; set; } = "";

        [JsonPropertyName("userlist")]
        public List<UserSimpleInfo> UserList { get; set; } = new List<UserSimpleInfo>();
    }

    public class UserSimpleInfo
    {
        [JsonPropertyName("userid")]
        public string UserId { get; set; } = "";

        [JsonPropertyName("name")]
        public string Name { get; set; } = "";

        [JsonPropertyName("department")]
        public List<int> Department { get; set; } = new List<int>();

        [JsonPropertyName("open_userid")]
        public string OpenUserId { get; set; } = "";
    }

    /// <summary>/cgi-bin/user/get 响应 (用于拉取单个用户的完整信息, 含 email/mobile/position)</summary>
    public class UserDetailResponse
    {
        [JsonPropertyName("errcode")]
        public int ErrCode { get; set; }

        [JsonPropertyName("errmsg")]
        public string ErrMsg { get; set; } = "";

        [JsonPropertyName("userid")]
        public string UserId { get; set; } = "";

        [JsonPropertyName("name")]
        public string Name { get; set; } = "";

        [JsonPropertyName("mobile")]
        public string Mobile { get; set; } = "";

        [JsonPropertyName("email")]
        public string Email { get; set; } = "";

        [JsonPropertyName("position")]
        public string Position { get; set; } = "";

        [JsonPropertyName("department")]
        public List<int> Department { get; set; } = new List<int>();

        [JsonPropertyName("status")]
        public int Status { get; set; }

        [JsonPropertyName("avatar")]
        public string Avatar { get; set; } = "";

        [JsonPropertyName("extattr")]
        public ExtAttrWrapper ExtAttr { get; set; }
    }

    public class ExtAttrWrapper
    {
        [JsonPropertyName("attrs")]
        public List<ExtAttrItem> Attrs { get; set; } = new List<ExtAttrItem>();
    }

    public class ExtAttrItem
    {
        [JsonPropertyName("name")]
        public string Name { get; set; } = "";

        [JsonPropertyName("value")]
        public string Value { get; set; } = "";
    }

    /// <summary>
    /// GetAllMembers 的返回: 成员列表 + 完整部门树 (id→DepartmentInfo)
    /// 调用方通过部门树的 parentid 反查可拼出完整路径
    /// </summary>
    public class WeChatWorkSyncResult
    {
        public List<UserDetailResponse> Members { get; set; } = new List<UserDetailResponse>();
        public List<DepartmentInfo> Departments { get; set; } = new List<DepartmentInfo>();
    }

    /// <summary>
    /// /cgi-bin/externalcontact/list 响应 (某员工名下的所有客户 external_userid)
    /// </summary>
    public class ExternalContactListResponse
    {
        [JsonPropertyName("errcode")]
        public int ErrCode { get; set; }

        [JsonPropertyName("errmsg")]
        public string ErrMsg { get; set; } = "";

        [JsonPropertyName("external_userid")]
        public List<string> ExternalUserId { get; set; } = new List<string>();
    }

    /// <summary>
    /// /cgi-bin/externalcontact/get 响应 (单个客户详情)
    /// </summary>
    public class ExternalContactDetailResponse
    {
        [JsonPropertyName("errcode")]
        public int ErrCode { get; set; }

        [JsonPropertyName("errmsg")]
        public string ErrMsg { get; set; } = "";

        [JsonPropertyName("external_contact")]
        public ExternalContactInfo ExternalContact { get; set; }

        [JsonPropertyName("follow_user")]
        public List<FollowUserInfo> FollowUser { get; set; } = new List<FollowUserInfo>();
    }

    public class ExternalContactInfo
    {
        [JsonPropertyName("external_userid")]
        public string ExternalUserId { get; set; } = "";

        [JsonPropertyName("name")]
        public string Name { get; set; } = "";

        /// <summary>1=微信用户, 2=企业微信用户</summary>
        [JsonPropertyName("type")]
        public int Type { get; set; }

        [JsonPropertyName("avatar")]
        public string Avatar { get; set; } = "";

        /// <summary>0=未知, 1=男, 2=女</summary>
        [JsonPropertyName("gender")]
        public int Gender { get; set; }

        [JsonPropertyName("unionid")]
        public string UnionId { get; set; } = "";

        /// <summary>客户所属员工的 userid (本字段由服务在调用时附加, 非 API 返回)</summary>
        [JsonIgnore]
        public string OwnerUserId { get; set; } = "";
    }

    public class FollowUserInfo
    {
        [JsonPropertyName("userid")]
        public string UserId { get; set; } = "";

        [JsonPropertyName("remark")]
        public string Remark { get; set; } = "";
    }
}
