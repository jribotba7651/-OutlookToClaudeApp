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

        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        public OutlookServiceV3()
        {
            try
            {
                // Get the CLSID for Outlook
                Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null)
                    throw new Exception("Outlook is not installed.");

                // Alternative to Marshal.GetActiveObject for .NET 8
                try
                {
                    Guid clsid = outlookType.GUID;
                    GetActiveObject(ref clsid, IntPtr.Zero, out var obj);
                    _outlookApp = obj;
                }
                catch
                {
                    _outlookApp = Activator.CreateInstance(outlookType);
                }

                if (_outlookApp == null)
                    throw new Exception("Could not start Outlook.");

                dynamic app = _outlookApp;
                _nameSpace = app.GetNamespace("MAPI");
                
                if (_nameSpace == null)
                    throw new Exception("Could not access MAPI.");

                // Use dynamic to call Logon to avoid object type errors
                dynamic ns = _nameSpace;
                ns.Logon(Type.Missing, Type.Missing, false, false);
            }
            catch (Exception ex)
            {
                throw new Exception($"Outlook Connection Error: {ex.Message}", ex);
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

                // Simple restriction string to reduce data handled
                string startFilter = startDate.ToString("g");
                string endFilter = endDate.AddDays(1).ToString("g");
                
                items.IncludeRecurrences = true;
                items.Sort("[Start]", false);

                // Use a basic loop with a counter to prevent infinite hangs
                int count = items.Count;
                int processed = 0;
                
                foreach (var item in items)
                {
                    if (processed > 500) break; // Safety limit
                    processed++;

                    try
                    {
                        dynamic appointment = item;
                        DateTime start = appointment.Start;
                        DateTime end = appointment.End;

                        if (start <= endDate.AddDays(1) && end >= startDate)
                        {
                            events.Add(new CalendarEvent
                            {
                                Subject = appointment.Subject ?? "No Subject",
                                Start = start,
                                End = end,
                                Location = appointment.Location ?? "",
                                Body = CleanBody(appointment.Body),
                                IsAllDayEvent = appointment.AllDayEvent,
                                Organizer = GetOrganizerName(appointment)
                            });
                        }
                    }
                    catch { /* Skip and continue */ }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error reading items: {ex.Message}");
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
