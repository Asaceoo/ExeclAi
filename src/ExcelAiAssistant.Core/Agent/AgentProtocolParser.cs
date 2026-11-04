using System;
using ExcelAiAssistant.Core.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelAiAssistant.Core.Agent
{
    public sealed class ParsedAgentAction
    {
        public string Action { get; set; }

        public Models.ToolEnvelope ToolCall { get; set; }

        public string Content { get; set; }
    }

    public static class AgentProtocolParser
    {
        public static bool TryParse(string modelOutput, out ParsedAgentAction action)
        {
            action = null;
            if (string.IsNullOrWhiteSpace(modelOutput))
            {
                return false;
            }

            if (TryParseObject(modelOutput, out var parsed) && IsProtocol(parsed))
            {
                action = FromJObject(parsed);
                return true;
            }

            foreach (var candidate in ExtractBalancedObjects(modelOutput))
            {
                if (TryParseObject(candidate, out var nested) && IsProtocol(nested))
                {
                    action = FromJObject(nested);
                    return true;
                }
            }

            action = new ParsedAgentAction
            {
                Action = "final",
                Content = modelOutput
            };
            return true;
        }

        private static bool TryParseObject(string value, out JObject result)
        {
            try
            {
                result = JObject.Parse(value);
                return true;
            }
            catch (JsonException)
            {
                result = null;
                return false;
        }
        }

        private static bool IsProtocol(JObject value)
        {
            var action = value.Value<string>("action");
            return string.Equals(action, "tool", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(action, "final", StringComparison.OrdinalIgnoreCase);
        }

        private static ParsedAgentAction FromJObject(JObject value)
        {
            var actionName = value.Value<string>("action");
            if (!IsProtocol(value))
            {
                throw new FormatException("不支持的 Agent action");
            }

            if (string.Equals(actionName, "final", StringComparison.OrdinalIgnoreCase))
            {
                return new ParsedAgentAction
                {
                    Action = "final",
                    Content = value.Value<string>("content") ?? string.Empty
                };
            }

            var args = value["args"] as JObject ?? new JObject();
            return new ParsedAgentAction
            {
                Action = "tool",
                ToolCall = new Models.ToolEnvelope
                {
                    Tool = value.Value<string>("tool"),
                    Args = args.ToObject<Dictionary<string, object>>() ??
                           new Dictionary<string, object>(),
                    Reason = value.Value<string>("reason"),
                    CallId = value.Value<string>("call_id")
                }
            };
        }

        private static IEnumerable<string> ExtractBalancedObjects(string text)
        {
            for (var index = 0; index < text.Length; index++)
            {
                if (text[index] != '{')
                {
                    continue;
                }

                var depth = 0;
                var inString = false;
                var escaped = false;

                for (var end = index; end < text.Length; end++)
                {
                    var current = text[end];
                    if (inString)
                    {
                        if (escaped)
                        {
                            escaped = false;
                        }
                        else if (current == '\\')
                        {
                            escaped = true;
                        }
                        else if (current == '"')
                        {
                            inString = false;
                        }

                        continue;
                    }

                    if (current == '"')
                    {
                        inString = true;
                    }
                    else if (current == '{')
                    {
                        depth++;
                    }
                    else if (current == '}')
                    {
                        depth--;
                        if (depth == 0)
                        {
                            yield return text.Substring(index, end - index + 1);
                            break;
                        }
                    }
                }
            }
        }
    }
}
