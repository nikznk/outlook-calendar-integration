using System;

namespace Outlook_Calendar.Models
{
    public class CalendarEvent
    {
        public CalendarEvent()
        {
            this.Body = new Body()
            {
                ContentType = "text"
            };
            this.Start = new EventDateTime()
            {
                TimeZone = "Asia/Kolkata"
            };
            this.End = new EventDateTime()
            {
                TimeZone = "Asia/Kolkata"
            };
            this.IsOnlineMeeting = true;
            this.OnlineMeetingProvider = "TeamsForBusiness";
        }

        public string Id { get; set; }
        public string Subject { get; set; }
        public Body Body { get; set; }
        public EventDateTime Start { get; set; }
        public EventDateTime End { get; set; }
        public bool IsOnlineMeeting { get; set; }
        public string OnlineMeetingProvider { get; set; }
        public OnlineMeeting OnlineMeeting { get; set; }
    }

    public class Body
    {
        public string ContentType { get; set; }
        public string Content { get; set; }
    }

    public class EventDateTime
    {
        public DateTime DateTime { get; set; }
        public string TimeZone { get; set; }
    }

    public class OnlineMeeting
    {
        public string JoinUrl { get; set; }
    }
}