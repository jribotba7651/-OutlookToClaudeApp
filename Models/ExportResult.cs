using System;

namespace OutlookToClaudeApp.Models
{
    public class ExportResult
    {
        public bool Success { get; set; }
        public string FileId { get; set; }
        public string Message { get; set; }
        public string MarkdownContent { get; set; }
        public ServiceType ServiceType { get; set; }
        public DateTime ExportedAt { get; set; } = DateTime.Now;

        public static ExportResult SuccessResult(string fileId, string message, ServiceType service)
        {
            return new ExportResult
            {
                Success = true,
                FileId = fileId,
                Message = message,
                ServiceType = service
            };
        }

        public static ExportResult ErrorResult(string message, ServiceType service)
        {
            return new ExportResult
            {
                Success = false,
                Message = message,
                ServiceType = service
            };
        }
    }
}
