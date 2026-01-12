using System;

namespace OutlookToClaudeApp.Models
{
    public class CalendarEvent
    {
        public string Subject { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public string Location { get; set; }
        public string Body { get; set; }
        public bool IsAllDayEvent { get; set; }
        public string Organizer { get; set; }
        public string Categories { get; set; }
        public bool IsSelected { get; set; } = false;

        public string DisplayTime => IsAllDayEvent
            ? "All Day"
            : $"{Start:h:mm tt} - {End:h:mm tt}";

        public string DisplayDate => Start.ToString("ddd, MMM dd");

        public string DisplayTitle => string.IsNullOrEmpty(Subject)
            ? "(No Subject)"
            : Subject;
    }
}
