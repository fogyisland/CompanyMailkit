using System;

namespace MailConverter.Services.Contacts
{
    /// <summary>
    /// PST 联系人数据模型 (从 Outlook ContactItem 抽取后用于 Graph 同步)
    /// </summary>
    public class ContactData
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
}
