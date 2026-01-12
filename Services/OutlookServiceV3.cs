using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using OutlookToClaudeApp.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookToClaudeApp.Services
{
    public class OutlookServiceV3 : IDisposable
    {
        private Outlook.Application _outlookApp;
        private Outlook.NameSpace _nameSpace;

        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        private static object GetActiveObject(string progId)
        {
            try
            {
                var type = Type.GetTypeFromProgID(progId);
                if (type == null) return null;

                var clsid = type.GUID;
                GetActiveObject(ref clsid, IntPtr.Zero, out var obj);
                return obj;
            }
            catch
            {
                return null;
            }
        }

        public OutlookServiceV3()
        {
            try
            {
                // Try to get existing Outlook instance first, then create new if not found
                try
                {
                    _outlookApp = (Outlook.Application)GetActiveObject("Outlook.Application");
                }
                catch
                {
                    // Ignore and create new
                }

                if (_outlookApp == null)
                {
                    _outlookApp = new Outlook.Application();
                }

                _nameSpace = _outlookApp.GetNamespace("MAPI");
                _nameSpace.Logon(Type.Missing, Type.Missing, false, false);
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Failed to connect to Outlook: {ex.Message}. Make sure Outlook is installed and configured.", ex);
            }
        }

        public List<CalendarEvent> GetEvents(DateTime startDate, DateTime endDate)
        {
            var events = new List<CalendarEvent>();
            Outlook.MAPIFolder calendarFolder = null;
            Outlook.Items items = null;

            try
            {
                calendarFolder = _nameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                items = calendarFolder.Items;

                // IMPORTANT: Include recurring appointments BEFORE sorting
                items.IncludeRecurrences = true;

                // Sort by start date
                items.Sort("[Start]", false);

                // Instead of using Restrict() which is problematic, iterate all items
                // and filter manually - this is more reliable
                var endDateInclusive = endDate.AddDays(1).AddSeconds(-1);

                foreach (object item in items)
                {
                    if (item is Outlook.AppointmentItem appointment)
                    {
                        try
                        {
                            // Filter: event overlaps if it starts before range end AND ends after range start
                            if (appointment.Start <= endDateInclusive && appointment.End >= startDate)
                            {
                                var calEvent = new CalendarEvent
                                {
                                    Subject = appointment.Subject ?? string.Empty,
                                    Start = appointment.Start,
                                    End = appointment.End,
                                    Location = appointment.Location ?? string.Empty,
                                    Body = CleanBody(appointment.Body),
                                    IsAllDayEvent = appointment.AllDayEvent,
                                    Organizer = GetOrganizerName(appointment),
                                    Categories = appointment.Categories ?? string.Empty
                                };

                                events.Add(calEvent);
                            }
                        }
                        catch (System.Exception ex)
                        {
                            // Log but continue processing other events
                            System.Diagnostics.Debug.WriteLine($"Error processing event: {ex.Message}");
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(appointment);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Failed to retrieve calendar events: {ex.Message}", ex);
            }
            finally
            {
                // Always cleanup COM objects
                if (items != null)
                    Marshal.ReleaseComObject(items);

                if (calendarFolder != null)
                    Marshal.ReleaseComObject(calendarFolder);
            }

            return events.OrderBy(e => e.Start).ToList();
        }

        private string CleanBody(string body)
        {
            if (string.IsNullOrWhiteSpace(body))
                return string.Empty;

            var cleaned = body.Trim();

            // Limit length to avoid huge text blocks
            if (cleaned.Length > 1000)
                cleaned = cleaned.Substring(0, 1000) + "...";

            return cleaned;
        }

        private string GetOrganizerName(Outlook.AppointmentItem appointment)
        {
            try
            {
                if (!string.IsNullOrEmpty(appointment.Organizer))
                    return appointment.Organizer;

                // Try to get organizer from GetOrganizer method
                var organizer = appointment.GetOrganizer();
                if (organizer != null)
                {
                    var name = organizer.Name;
                    Marshal.ReleaseComObject(organizer);
                    return name ?? string.Empty;
                }
            }
            catch
            {
                // If we can't get organizer, just return empty
            }

            return string.Empty;
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
            try
            {
                if (_nameSpace != null)
                {
                    _nameSpace.Logoff();
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
            catch
            {
                // Suppress errors during disposal
            }
        }
    }
}
