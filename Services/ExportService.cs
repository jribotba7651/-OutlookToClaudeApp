using System;
using System.Collections.Generic;
using System.Text;
using OutlookToClaudeApp.Models;

namespace OutlookToClaudeApp.Services
{
    public class ExportService
    {
        public string GenerateMarkdown(List<CalendarEvent> events)
        {
            var sb = new StringBuilder();
            sb.AppendLine("# Calendar Events");
            sb.AppendLine();
            sb.AppendLine($"**Export Date:** {DateTime.Now:yyyy-MM-dd HH:mm}");
            sb.AppendLine($"**Total Events:** {events.Count}");
            sb.AppendLine();
            sb.AppendLine("---");

            foreach (var evt in events)
            {
                sb.AppendLine($"## {evt.DisplayTitle}");
                sb.AppendLine($"**Date:** {evt.DisplayDate}");
                sb.AppendLine($"**Time:** {evt.DisplayTime}");
                if (!string.IsNullOrEmpty(evt.Location)) sb.AppendLine($"**Location:** {evt.Location}");
                if (!string.IsNullOrEmpty(evt.Organizer)) sb.AppendLine($"**Organizer:** {evt.Organizer}");
                if (!string.IsNullOrEmpty(evt.Body))
                {
                    sb.AppendLine();
                    sb.AppendLine(evt.Body.Trim());
                }
                sb.AppendLine();
                sb.AppendLine("---");
            }
            return sb.ToString();
        }

        public string GenerateCsv(List<CalendarEvent> events)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Subject,Start Date,Start Time,End Date,End Time,Location,Organizer,Body");

            foreach (var evt in events)
            {
                var row = new List<string>
                {
                    EscapeCsv(evt.Subject),
                    evt.Start.ToShortDateString(),
                    evt.Start.ToShortTimeString(),
                    evt.End.ToShortDateString(),
                    evt.End.ToShortTimeString(),
                    EscapeCsv(evt.Location),
                    EscapeCsv(evt.Organizer),
                    EscapeCsv(evt.Body)
                };
                sb.AppendLine(string.Join(",", row));
            }
            return sb.ToString();
        }

        public string GenerateTxt(List<CalendarEvent> events)
        {
            var sb = new StringBuilder();
            sb.AppendLine("CALENDAR EVENTS EXPORT");
            sb.AppendLine($"Generated: {DateTime.Now}");
            sb.AppendLine("============================================");
            sb.AppendLine();

            foreach (var evt in events)
            {
                sb.AppendLine($"EVENT: {evt.DisplayTitle}");
                sb.AppendLine($"WHEN:  {evt.DisplayDate} | {evt.DisplayTime}");
                if (!string.IsNullOrEmpty(evt.Location)) sb.AppendLine($"WHERE: {evt.Location}");
                if (!string.IsNullOrEmpty(evt.Organizer)) sb.AppendLine($"WHO:   {evt.Organizer}");
                sb.AppendLine("--------------------------------------------");
                if (!string.IsNullOrEmpty(evt.Body))
                {
                    sb.AppendLine(evt.Body.Trim());
                }
                sb.AppendLine("============================================");
                sb.AppendLine();
            }
            return sb.ToString();
        }

        private string EscapeCsv(string field)
        {
            if (string.IsNullOrEmpty(field)) return "";
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
            {
                return $"\"{field.Replace("\"", "\"\"")}\"";
            }
            return field;
        }
    }
}
