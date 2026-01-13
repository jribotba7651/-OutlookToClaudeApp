using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using OutlookToClaudeApp.Models;

namespace OutlookToClaudeApp.Services
{
    public class OutlookServiceV3 : IDisposable
    {
        private object _outlookApp;
        private object _nameSpace;

        public OutlookServiceV3()
        {
            try
            {
                // Try to get the COM type for Outlook
                Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null)
                    throw new Exception("Outlook is not installed on this system.");

                // Attempt to get an active instance or create a new one
                try
                {
                    _outlookApp = Marshal.GetActiveObject("Outlook.Application");
                }
                catch
                {
                    _outlookApp = Activator.CreateInstance(outlookType);
                }

                if (_outlookApp == null)
                    throw new Exception("Could not start Outlook. Application object is null.");

                // Use dynamic to access MAPI namespace
                dynamic app = _outlookApp;
                _nameSpace = app.GetNamespace("MAPI");
                
                if (_nameSpace == null)
                    throw new Exception("Could not access Outlook MAPI namespace.");

                _nameSpace.Logon(Type.Missing, Type.Missing, false, false);
            }
            catch (Exception ex)
            {
                string detail = ex.InnerException != null ? $"\nInner: {ex.InnerException.Message}" : "";
                throw new Exception($"Outlook Connection Error: {ex.Message}{detail}\n\nTIP: Check if Outlook is open and NOT using 'New Outlook' switch.", ex);
            }
        }

        public List<CalendarEvent> GetEvents(DateTime startDate, DateTime endDate)
        {
            var events = new List<CalendarEvent>();
            try
            {
                dynamic ns = _nameSpace;
                // 9 = olFolderCalendar
                dynamic calendarFolder = ns.GetDefaultFolder(9);
                dynamic items = calendarFolder.Items;

                items.IncludeRecurrences = true;
                items.Sort("[Start]", false);

                var endDateInclusive = endDate.AddDays(1).AddSeconds(-1);

                foreach (var item in items)
                {
                    try
                    {
                        // Use dynamic properties to avoid type checking issues
                        dynamic appointment = item;
                        
                        DateTime start = appointment.Start;
                        DateTime end = appointment.End;

                        if (start <= endDateInclusive && end >= startDate)
                        {
                            var calEvent = new CalendarEvent
                            {
                                Subject = appointment.Subject ?? string.Empty,
                                Start = start,
                                End = end,
                                Location = appointment.Location ?? string.Empty,
                                Body = CleanBody(appointment.Body),
                                IsAllDayEvent = appointment.AllDayEvent,
                                Organizer = GetOrganizerName(appointment),
                                Categories = appointment.Categories ?? string.Empty
                            };

                            events.Add(calEvent);
                        }
                    }
                    catch
                    {
                        // Skip items that aren't appointments or have access errors
                    }
                    finally
                    {
                        if (item != null && Marshal.IsComObject(item))
                            Marshal.ReleaseComObject(item);
                    }
                }

                if (items != null) Marshal.ReleaseComObject(items);
                if (calendarFolder != null) Marshal.ReleaseComObject(calendarFolder);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error retrieving events: {ex.Message}", ex);
            }

            return events.OrderBy(e => e.Start).ToList();
        }

        private string CleanBody(string body)
        {
            if (string.IsNullOrWhiteSpace(body)) return string.Empty;
            var cleaned = body.Trim();
            if (cleaned.Length > 1000) cleaned = cleaned.Substring(0, 1000) + "...";
            return cleaned;
        }

        private string GetOrganizerName(dynamic appointment)
        {
            try
            {
                return appointment.Organizer ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public void Dispose()
        {
            try
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
            catch { }
        }
    }
}
