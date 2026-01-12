using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using OutlookToClaudeApp.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookToClaudeApp.Services
{
    public class OutlookService : IDisposable
    {
        private Outlook.Application _outlookApp;
        private Outlook.NameSpace _nameSpace;

        public OutlookService()
        {
            try
            {
                // Create new Outlook Application instance
                _outlookApp = new Outlook.Application();
                _nameSpace = _outlookApp.GetNamespace("MAPI");
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Failed to connect to Outlook: {ex.Message}. Make sure Outlook is installed and configured.", ex);
            }
        }

        public List<CalendarEvent> GetEvents(DateTime startDate, DateTime endDate)
        {
            var events = new List<CalendarEvent>();

            try
            {
                var calendarFolder = _nameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                var items = calendarFolder.Items;
                items.Sort("[Start]", false);
                items.IncludeRecurrences = true;

                var filter = $"[Start] >= '{startDate:M/d/yyyy}' AND [End] <= '{endDate:M/d/yyyy 11:59:59 PM}'";
                var restrictedItems = items.Restrict(filter);

                foreach (object item in restrictedItems)
                {
                    if (item is Outlook.AppointmentItem appointment)
                    {
                        try
                        {
                            events.Add(new CalendarEvent
                            {
                                Subject = appointment.Subject ?? string.Empty,
                                Start = appointment.Start,
                                End = appointment.End,
                                Location = appointment.Location ?? string.Empty,
                                Body = appointment.Body ?? string.Empty,
                                IsAllDayEvent = appointment.AllDayEvent,
                                Organizer = appointment.Organizer ?? string.Empty,
                                Categories = appointment.Categories ?? string.Empty
                            });
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(appointment);
                        }
                    }
                }

                Marshal.ReleaseComObject(restrictedItems);
                Marshal.ReleaseComObject(items);
                Marshal.ReleaseComObject(calendarFolder);
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Failed to retrieve calendar events: {ex.Message}", ex);
            }

            return events.OrderBy(e => e.Start).ToList();
        }

        public bool IsOutlookRunning()
        {
            try
            {
                return _outlookApp != null && _nameSpace != null;
            }
            catch
            {
                return false;
            }
        }

        public void Dispose()
        {
            if (_nameSpace != null)
            {
                Marshal.ReleaseComObject(_nameSpace);
                _nameSpace = null;
            }

            if (_outlookApp != null)
            {
                Marshal.ReleaseComObject(_outlookApp);
                _outlookApp = null;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
