using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using OutlookToClaudeApp.Models;

namespace OutlookToClaudeApp.Services
{
    public class OutlookServiceV2 : IDisposable
    {
        public List<CalendarEvent> GetEvents(DateTime startDate, DateTime endDate)
        {
            var events = new List<CalendarEvent>();

            try
            {
                // Create PowerShell script to read Outlook calendar
                var script = CreatePowerShellScript(startDate, endDate);

                // Execute PowerShell and parse results
                var output = ExecutePowerShell(script);
                events = ParseEventsFromJson(output);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to retrieve calendar events: {ex.Message}. Make sure Outlook is installed and running.", ex);
            }

            return events.OrderBy(e => e.Start).ToList();
        }

        private string CreatePowerShellScript(DateTime start, DateTime end)
        {
            var startStr = start.ToString("M/d/yyyy");
            var endStr = end.ToString("M/d/yyyy");

            var script = @"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace('MAPI')
    $folder = $namespace.GetDefaultFolder(9)
    $items = $folder.Items
    $items.Sort('[Start]')
    $items.IncludeRecurrences = $false

    $startDate = [DateTime]::Parse('" + startStr + @"')
    $endDate = [DateTime]::Parse('" + endStr + @"').AddDays(1).AddSeconds(-1)

    $events = @()
    foreach ($item in $items) {
        if ($item.Class -eq 26) {
            $itemStart = [DateTime]$item.Start
            $itemEnd = [DateTime]$item.End

            if ($itemStart -ge $startDate -and $itemEnd -le $endDate) {
                $bodyText = if ($item.Body) { $item.Body.Trim() -replace '[\r\n]+', ' ' } else { '' }
                if ($bodyText.Length -gt 500) { $bodyText = $bodyText.Substring(0, 500) }

                $eventObj = @{
                    Subject = if ($item.Subject) { $item.Subject.Trim() } else { '' }
                    Start = $item.Start.ToString('yyyy-MM-ddTHH:mm:ss')
                    End = $item.End.ToString('yyyy-MM-ddTHH:mm:ss')
                    Location = if ($item.Location) { $item.Location.Trim() } else { '' }
                    Body = $bodyText
                    IsAllDayEvent = $item.AllDayEvent
                    Organizer = if ($item.Organizer) { $item.Organizer.Trim() } else { '' }
                    Categories = if ($item.Categories) { $item.Categories.Trim() } else { '' }
                }
                $events += $eventObj
            }
        }
    }

    if ($events.Count -eq 0) {
        Write-Output '[]'
    } else {
        $events | ConvertTo-Json -Compress -Depth 3
    }
} catch {
    Write-Error $_.Exception.Message
    exit 1
}
";
            return script;
        }

        private string ExecutePowerShell(string script)
        {
            var processInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{script.Replace("\"", "`\"")}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (var process = Process.Start(processInfo))
            {
                var output = process.StandardOutput.ReadToEnd();
                var error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (process.ExitCode != 0 || !string.IsNullOrEmpty(error))
                {
                    throw new Exception($"PowerShell execution failed: {error}");
                }

                return output;
            }
        }

        private List<CalendarEvent> ParseEventsFromJson(string json)
        {
            var events = new List<CalendarEvent>();

            if (string.IsNullOrWhiteSpace(json))
                return events;

            try
            {
                dynamic data = Newtonsoft.Json.JsonConvert.DeserializeObject(json);

                // Handle both single object and array
                var items = data is Newtonsoft.Json.Linq.JArray ? data : new[] { data };

                foreach (var item in items)
                {
                    events.Add(new CalendarEvent
                    {
                        Subject = item.Subject?.ToString() ?? string.Empty,
                        Start = DateTime.Parse(item.Start.ToString()),
                        End = DateTime.Parse(item.End.ToString()),
                        Location = item.Location?.ToString() ?? string.Empty,
                        Body = item.Body?.ToString() ?? string.Empty,
                        IsAllDayEvent = bool.Parse(item.IsAllDayEvent.ToString()),
                        Organizer = item.Organizer?.ToString() ?? string.Empty,
                        Categories = item.Categories?.ToString() ?? string.Empty
                    });
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to parse calendar data: {ex.Message}", ex);
            }

            return events;
        }

        public bool IsOutlookRunning()
        {
            var processes = Process.GetProcessesByName("OUTLOOK");
            return processes.Length > 0;
        }

        public void Dispose()
        {
            // Nothing to dispose
        }
    }
}
