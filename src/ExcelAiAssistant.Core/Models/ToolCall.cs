using System.Collections.Generic;

namespace ExcelAiAssistant.Core.Models
{
    public sealed class ToolCallRequest
    {
        public string Id { get; set; }

        public string Name { get; set; }

        public Dictionary<string, object> Arguments { get; set; }
    }

    public sealed class ToolEnvelope
    {
        public string Tool { get; set; }

        public Dictionary<string, object> Args { get; set; } =
            new Dictionary<string, object>();

        public string Reason { get; set; }

        public string CallId { get; set; }
    }

    public sealed class ToolResult
    {
        public bool Ok { get; set; }

        public object Data { get; set; }

        public string Warning { get; set; }

        public List<string> AffectedRanges { get; set; } = new List<string>();

        public bool Truncated { get; set; }

        public string ErrorCode { get; set; }
    }

    public enum RiskLevel
    {
        ReadOnly = 0,
        NormalWrite = 1,
        StructuralChange = 2,
        Dangerous = 3
    }
}
