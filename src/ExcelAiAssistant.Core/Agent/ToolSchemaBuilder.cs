using System.Collections.Generic;
using ExcelAiAssistant.Core.Diagnostics;
using System.Linq;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelAiAssistant.Core.Agent
{
    public static class ToolSchemaBuilder
    {
        public static JArray Build(IReadOnlyList<ToolDescriptor> tools)
        {
            var array = new JArray();
            foreach (var tool in tools.OrderBy(item => item.Name))
            {
                array.Add(new JObject
                {
                    ["type"] = "function",
                    ["function"] = new JObject
                    {
                        ["name"] = tool.Name,
                        ["description"] = tool.Description,
                        ["parameters"] = tool.Parameters == null
                            ? new JObject { ["type"] = "object", ["properties"] = new JObject() }
                            : JObject.FromObject(tool.Parameters)
                    }
                });
            }

            return array;
        }
    }

    public static class ToolCallFactory
    {
        public static ToolCallRequest FromNative(JToken token)
        {
            var functionId = token.Value<string>("id");
            var function = token["function"];
            var name = function?.Value<string>("name") ?? string.Empty;
            var argumentsText = function?.Value<string>("arguments") ?? "{}";
            Dictionary<string, object> arguments;
            try
            {
                arguments = JObject.Parse(argumentsText)
                    .ToObject<Dictionary<string, object>>() ?? new Dictionary<string, object>();
            }
            catch
            {
                AppLog.Warn("ToolSchemaBuilder", "[LogCover] ToolSchemaBuilder.cs:49  silent-catch covered (no exception variable)");
                arguments = new Dictionary<string, object>();
            }

            return new ToolCallRequest
            {
                Id = functionId,
                Name = name,
                Arguments = arguments
            };
        }
    }

    public static class ChatMessageSerializer
    {
        public static JArray Serialize(IReadOnlyList<ChatMessage> messages)
        {
            var array = new JArray();
            foreach (var message in messages)
            {
                var obj = new JObject { ["role"] = message.Role };
                if (message.ToolCalls != null && message.ToolCalls.Count > 0)
                {
                    obj["tool_calls"] = new JArray(message.ToolCalls.Select(call => new JObject
                    {
                        ["id"] = call.Id ?? call.Name,
                        ["type"] = "function",
                        ["function"] = new JObject
                        {
                            ["name"] = call.Name,
                            ["arguments"] = call.Arguments == null
                                ? "{}"
                                : JObject.FromObject(call.Arguments).ToString(Formatting.None)
                        }
                    }));
                }

                if (!string.IsNullOrEmpty(message.Content))
                {
                    obj["content"] = message.Content;
                }

                if (message.Role == "tool" && !string.IsNullOrEmpty(message.ToolCallId))
                {
                    obj["tool_call_id"] = message.ToolCallId;
                }

                if (message.Role == "function" && !string.IsNullOrEmpty(message.Name))
                {
                    obj["name"] = message.Name;
                }

                array.Add(obj);
            }

            return array;
        }
    }
}