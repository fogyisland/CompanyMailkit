using System;

namespace MailConverter.Services.Calendars
{
    /// <summary>
    /// PST 日历数据模型 (从 Outlook AppointmentItem 抽取后用于 Graph 同步)
    /// </summary>
    public class CalendarData
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
}
