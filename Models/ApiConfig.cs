using System;

namespace OutlookToClaudeApp.Models
{
    public class ApiConfig
    {
        public string ClaudeApiKey { get; set; }
        public string ChatGPTApiKey { get; set; }
        public string GeminiApiKey { get; set; }
        public string PerplexityApiKey { get; set; }

        public ExportMode DefaultExportMode { get; set; } = ExportMode.ApiOnly;
    }

    public enum ExportMode
    {
        ApiOnly,
        CopyToClipboard
    }

    public enum ServiceType
    {
        Claude,
        ChatGPT,
        Gemini,
        Perplexity
    }
}
