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
}
