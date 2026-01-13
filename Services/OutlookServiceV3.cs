using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
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
                Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null) throw new Exception("Outlook no instalado.");

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

                if (_outlookApp == null) throw new Exception("No se pudo iniciar Outlook.");

                _nameSpace = _outlookApp.GetType().InvokeMember("GetNamespace", BindingFlags.InvokeMethod, null, _outlookApp, new object[] { "MAPI" });
                _nameSpace.GetType().InvokeMember("Logon", BindingFlags.InvokeMethod, null, _nameSpace, new object[] { Missing.Value, Missing.Value, false, false });
            }
            catch (Exception ex)
            {
                throw new Exception($"Error de Conexión: {ex.Message}\n\nTIP: Verifica que Outlook esté abierto y NO sea la versión 'New Outlook'.", ex);
            }
        }

        public List<CalendarEvent> GetEvents(DateTime startDate, DateTime endDate)
        {
            var events = new List<CalendarEvent>();
            object items = null;
            object calendarFolder = null;

            try
            {
                // 9 = olFolderCalendar
                calendarFolder = _nameSpace.GetType().InvokeMember("GetDefaultFolder", BindingFlags.InvokeMethod, null, _nameSpace, new object[] { 9 });
                items = calendarFolder.GetType().InvokeMember("Items", BindingFlags.GetProperty, null, calendarFolder, null);

                // Configurar Recurrencias
                items.GetType().InvokeMember("IncludeRecurrences", BindingFlags.SetProperty, null, items, new object[] { true });
                items.GetType().InvokeMember("Sort", BindingFlags.InvokeMethod, null, items, new object[] { "[Start]", false });

                // Filtro para que sea RÁPIDO y no se congele
                string filter = $"[Start] >= '{startDate:g}' AND [End] <= '{endDate.AddDays(1):g}'";
                object restrictedItems = items.GetType().InvokeMember("Restrict", BindingFlags.InvokeMethod, null, items, new object[] { filter });

                int count = (int)restrictedItems.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, restrictedItems, null);
                
                // Limitar a los primeros 200 por seguridad
                int toProcess = Math.Min(count, 200);

                for (int i = 1; i <= toProcess; i++)
                {
                    object appointment = null;
                    try
                    {
                        appointment = restrictedItems.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, restrictedItems, new object[] { i });
                        
                        events.Add(new CalendarEvent
                        {
                            Subject = GetProp(appointment, "Subject")?.ToString() ?? "Sin Asunto",
                            Start = (DateTime)GetProp(appointment, "Start"),
                            End = (DateTime)GetProp(appointment, "End"),
                            Location = GetProp(appointment, "Location")?.ToString() ?? "",
                            Body = CleanBody(GetProp(appointment, "Body")?.ToString()),
                            IsAllDayEvent = (bool)GetProp(appointment, "AllDayEvent"),
                            Organizer = GetProp(appointment, "Organizer")?.ToString() ?? ""
                        });
                    }
                    catch { }
                    finally { if (appointment != null) Marshal.ReleaseComObject(appointment); }
                }

                if (restrictedItems != null) Marshal.ReleaseComObject(restrictedItems);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al leer eventos: {ex.Message}");
            }
            finally
            {
                if (items != null) Marshal.ReleaseComObject(items);
                if (calendarFolder != null) Marshal.ReleaseComObject(calendarFolder);
            }

            return events.OrderBy(e => e.Start).ToList();
        }

        private object GetProp(object obj, string name)
        {
            try { return obj.GetType().InvokeMember(name, BindingFlags.GetProperty, null, obj, null); }
            catch { return null; }
        }

        private string CleanBody(string body)
        {
            if (string.IsNullOrWhiteSpace(body)) return "";
            var cleaned = body.Trim();
            if (cleaned.Length > 800) cleaned = cleaned.Substring(0, 800) + "...";
            return cleaned;
        }

        public void Dispose()
        {
            try
            {
                if (_nameSpace != null) Marshal.ReleaseComObject(_nameSpace);
                if (_outlookApp != null) Marshal.ReleaseComObject(_outlookApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch { }
        }
    }
}
