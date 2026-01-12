using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using OutlookToClaudeApp.Models;

namespace OutlookToClaudeApp.Services
{
    public class ClaudeApiService
    {
        private readonly string _apiKey;
        private readonly HttpClient _httpClient;
        private const string BaseUrl = "https://api.anthropic.com/v1";
        private const string AnthropicVersion = "2023-06-01";
        private const string BetaHeader = "files-api-2025-04-14";

        public ClaudeApiService(string apiKey)
        {
            _apiKey = apiKey ?? throw new ArgumentNullException(nameof(apiKey));
            _httpClient = new HttpClient();
        }

        public async Task<ExportResult> UploadCalendarAsync(List<CalendarEvent> events)
        {
            try
            {
                // 1. Generate Markdown content
                var markdown = GenerateMarkdown(events);

                // 2. Save to temp file
                var tempFile = Path.Combine(Path.GetTempPath(), $"calendar-{DateTime.Now:yyyyMMdd-HHmmss}.md");
                await File.WriteAllTextAsync(tempFile, markdown);

                // 3. Upload to Claude API
                var fileId = await UploadFileAsync(tempFile);

                // 4. Clean up temp file
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }

                var result = ExportResult.SuccessResult(
                    fileId,
                    $"Successfully uploaded {events.Count} events to Claude API",
                    ServiceType.Claude
                );
                result.MarkdownContent = markdown;
                return result;
            }
            catch (Exception ex)
            {
                return ExportResult.ErrorResult(
                    $"Failed to upload to Claude API: {ex.Message}",
                    ServiceType.Claude
                );
            }
        }

        private async Task<string> UploadFileAsync(string filePath)
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, $"{BaseUrl}/files");

            request.Headers.Add("x-api-key", _apiKey);
            request.Headers.Add("anthropic-version", AnthropicVersion);
            request.Headers.Add("anthropic-beta", BetaHeader);

            var content = new MultipartFormDataContent();
            var fileContent = new ByteArrayContent(await File.ReadAllBytesAsync(filePath));
            fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("text/markdown");
            content.Add(fileContent, "file", Path.GetFileName(filePath));

            request.Content = content;

            var response = await _httpClient.SendAsync(request);
            var responseContent = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                throw new Exception($"API returned {response.StatusCode}: {responseContent}");
            }

            var result = JsonConvert.DeserializeObject<dynamic>(responseContent);
            return result.id.ToString();
        }

        public string GenerateMarkdown(List<CalendarEvent> events)
        {
            var sb = new StringBuilder();

            sb.AppendLine("# Calendar Events");
            sb.AppendLine();
            sb.AppendLine($"**Export Date:** {DateTime.Now:yyyy-MM-dd HH:mm}");
            sb.AppendLine($"**Total Events:** {events.Count}");
            sb.AppendLine();
            sb.AppendLine("---");
            sb.AppendLine();

            var eventsByDate = events.GroupBy(e => e.Start.Date).OrderBy(g => g.Key);

            foreach (var dateGroup in eventsByDate)
            {
                sb.AppendLine($"## {dateGroup.Key:dddd, MMMM dd, yyyy}");
                sb.AppendLine();

                foreach (var evt in dateGroup.OrderBy(e => e.Start))
                {
                    sb.AppendLine($"### {evt.DisplayTitle}");
                    sb.AppendLine();

                    if (evt.IsAllDayEvent)
                    {
                        sb.AppendLine("**Time:** All Day Event");
                    }
                    else
                    {
                        sb.AppendLine($"**Time:** {evt.Start:h:mm tt} - {evt.End:h:mm tt}");
                    }

                    if (!string.IsNullOrWhiteSpace(evt.Location))
                    {
                        sb.AppendLine($"**Location:** {evt.Location}");
                    }

                    if (!string.IsNullOrWhiteSpace(evt.Organizer))
                    {
                        sb.AppendLine($"**Organizer:** {evt.Organizer}");
                    }

                    if (!string.IsNullOrWhiteSpace(evt.Categories))
                    {
                        sb.AppendLine($"**Categories:** {evt.Categories}");
                    }

                    if (!string.IsNullOrWhiteSpace(evt.Body))
                    {
                        sb.AppendLine();
                        sb.AppendLine("**Details:**");
                        sb.AppendLine();
                        // Clean up body text (remove excessive whitespace)
                        var cleanBody = evt.Body.Trim();
                        if (cleanBody.Length > 500)
                        {
                            cleanBody = cleanBody.Substring(0, 500) + "...";
                        }
                        sb.AppendLine(cleanBody);
                    }

                    sb.AppendLine();
                    sb.AppendLine("---");
                    sb.AppendLine();
                }
            }

            sb.AppendLine();
            sb.AppendLine("*Generated by Outlook to Claude Calendar App*");

            return sb.ToString();
        }

        public async Task<bool> TestApiKeyAsync()
        {
            try
            {
                // Test by attempting to list files
                using var request = new HttpRequestMessage(HttpMethod.Get, $"{BaseUrl}/files");

                request.Headers.Add("x-api-key", _apiKey);
                request.Headers.Add("anthropic-version", AnthropicVersion);
                request.Headers.Add("anthropic-beta", BetaHeader);

                var response = await _httpClient.SendAsync(request);
                return response.IsSuccessStatusCode;
            }
            catch
            {
                return false;
            }
        }
    }
}
