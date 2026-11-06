using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelAiAssistant.Core.Diagnostics;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Search;
using ExcelAiAssistant.Core.Tools;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelAiAssistant.Core.Agent
{
    public sealed class ChatMessage
    {
        public string Role { get; set; }

        public string Content { get; set; }

        public IReadOnlyList<ToolCallRequest> ToolCalls { get; set; }

        public string ToolCallId { get; set; }

        public string Name { get; set; }

        /// <summary>
        /// 当轮工具调用的界面展示记录，仅用于会话持久化与前端回看，
        /// 不会通过 ChatMessageSerializer 发送给模型。
        /// </summary>
        public IReadOnlyList<ToolEventView> ToolEvents { get; set; }

        /// <summary>
        /// 本轮模型的思维链推理文字（reasoning_content / reasoning），
        /// 仅用于界面展示与持久化，不发送给模型。
        /// </summary>
        public string Reasoning { get; set; }

        /// <summary>
        /// 按真实顺序组织的步骤序列（思考片段 / 工具调用及其结果），
        /// 仅用于界面展示与持久化，不发送给模型。
        /// </summary>
        public IReadOnlyList<AgentStepView> Steps { get; set; }

        /// <summary>本轮次标签（如「第 2 轮」），用于把回滚按钮对回该轮产生的操作。</summary>
        public string TurnLabel { get; set; }

        /// <summary>
        /// 本轮成功执行 run_vba_snippet 后由系统自动沉淀为可复用工具（true 时前端不再显示
        /// 「沉淀为工具」兜底按钮，改为「已自动沉淀」徽标）。仅用于界面展示与持久化，不发送给模型。
        /// </summary>
        public bool AutoSinkedVba { get; set; }

        public static ChatMessage System(string content)
        {
            return new ChatMessage { Role = "system", Content = content };
        }

        public static ChatMessage User(string content)
        {
            return new ChatMessage { Role = "user", Content = content };
        }

        public static ChatMessage Assistant(string content)
        {
            return new ChatMessage { Role = "assistant", Content = content };
        }

        public static ChatMessage AssistantWithToolCalls(IReadOnlyList<ToolCallRequest> toolCalls)
        {
            return new ChatMessage
            {
                Role = "assistant",
                Content = string.Empty,
                ToolCalls = toolCalls
            };
        }

        public static ChatMessage ToolResult(string toolCallId, string content)
        {
            return new ChatMessage
            {
                Role = "tool",
                Content = content,
                ToolCallId = toolCallId
            };
        }

        public static ChatMessage FunctionResult(string name, string content)
        {
            return new ChatMessage
            {
                Role = "function",
                Content = content,
                Name = name
            };
        }
    }

    public sealed class ChatCompletionResult
    {
        public string Content { get; set; }

        /// <summary>模型的思维链推理文字（OpenAI 系为 reasoning_content，部分厂商为 reasoning）。</summary>
        public string Reasoning { get; set; }

        public IReadOnlyList<ToolCallRequest> ToolCalls { get; set; } =
            Array.Empty<ToolCallRequest>();

        public int PromptTokens { get; set; }

        public int CompletionTokens { get; set; }

        public int TotalTokens { get; set; }

        /// <summary>是否因原生工具调用失败而自动降级为 Fallback(JSON) 协议。</summary>
        public bool FallbackUsed { get; set; }
    }

    public sealed class OpenAiCompatibleClient : IChatCompletionClient
    {
        private readonly HttpClient httpClient;
        private readonly ModelProfile profile;

        /// <summary>
        /// 联网搜索探测结果（v3.8.1.83 批次，主 A 兜 B）：由宿主桥接按设置开关与 baseUrl 探测后设置；
        /// null（缺省）= 不注入任何搜索参数，行为与历史版本完全一致。ProviderNative 时按提供方
        /// 注入服务端搜索参数（智谱 web_search / Kimi builtin_function / 通义 enable_search /
        /// OpenRouter :online / 千帆 web_search），协议降级重试时自动跳过注入（防注入参数引发 400 循环）。
        /// </summary>
        public SearchProviderDetector.Detection SearchDetection { get; set; }

        public OpenAiCompatibleClient(ModelProfile profile)
            : this(profile, new HttpClient())
        {
        }

        public OpenAiCompatibleClient(ModelProfile profile, HttpClient httpClient)
        {
            this.profile = profile ?? throw new ArgumentNullException(nameof(profile));
            this.httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            this.profile.Validate();
            this.httpClient.Timeout = Timeout.InfiniteTimeSpan;
            if (profile.AuthScheme == AuthScheme.ApiKey)
            {
                this.httpClient.DefaultRequestHeaders.Add("api-key", profile.ApiKey);
            }
            else
            {
                this.httpClient.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", profile.ApiKey);
            }
        }

        public async Task<ChatCompletionResult> CompleteAsync(
            IReadOnlyList<ChatMessage> messages,
            CancellationToken cancellationToken)
        {
            return await CompleteAsync(messages, null, cancellationToken, null).ConfigureAwait(false);
        }

        public async Task<ChatCompletionResult> CompleteAsync(
            IReadOnlyList<ChatMessage> messages,
            IReadOnlyList<ToolDescriptor> tools,
            CancellationToken cancellationToken)
        {
            return await CompleteAsync(messages, tools, cancellationToken, null).ConfigureAwait(false);
        }

        public async Task<ChatCompletionResult> CompleteAsync(
            IReadOnlyList<ChatMessage> messages,
            IReadOnlyList<ToolDescriptor> tools,
            CancellationToken cancellationToken,
            Action<AgentProgress> onProgress)
        {
            if (messages == null || messages.Count == 0)
            {
                throw new ArgumentException("消息不能为空。", nameof(messages));
            }

            var useNative = tools != null && tools.Count > 0 &&
                            profile.ProtocolMode != AgentProtocolMode.Fallback;

            try
            {
                return await CompleteAttemptAsync(
                    messages,
                    useNative ? tools : null,
                    fallbackInstruction: false,
                    onProgress,
                    cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex) when (
                Retriable(ex) &&
                useNative &&
                profile.ProtocolMode == AgentProtocolMode.Auto)
            {
                // 原生工具调用失败时自动降级为 Fallback(JSON) 协议重试一次。
                var fallbackMessages = WithFallbackInstruction(messages);
                var fallback = await CompleteAttemptAsync(
                    fallbackMessages,
                    null,
                    fallbackInstruction: true,
                    onProgress,
                    cancellationToken).ConfigureAwait(false);
                fallback.FallbackUsed = true;
                return fallback;
            }
        }

        private async Task<ChatCompletionResult> CompleteAttemptAsync(
            IReadOnlyList<ChatMessage> messages,
            IReadOnlyList<ToolDescriptor> tools,
            bool fallbackInstruction,
            Action<AgentProgress> onProgress,
            CancellationToken cancellationToken)
        {
            EnsureTls12();

            var request = new JObject
            {
                ["model"] = profile.ModelName,
                ["temperature"] = profile.Temperature,
                ["stream"] = true,
                ["messages"] = ChatMessageSerializer.Serialize(messages)
            };

            if (profile.MaxTokens.HasValue)
            {
                request["max_tokens"] = profile.MaxTokens.Value;
            }

            if (!fallbackInstruction && tools != null && tools.Count > 0)
            {
                request["tools"] = ToolSchemaBuilder.Build(tools);
            }

            InjectProviderSearch(request, fallbackInstruction);

            using var http = new HttpRequestMessage(HttpMethod.Post, BuildChatCompletionsUrl());
            http.Content = new StringContent(
                request.ToString(Formatting.None),
                Encoding.UTF8,
                "application/json");

            return await SendStreamingAsync(http, onProgress, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// 以 SSE 增量读流方式请求 chat/completions：逐行读取，每收到一段 reasoning_content
        /// 就立刻回调 onProgress(ReasoningDelta)（实时流式思考）。若响应体不含 SSE data: 行
        /// （如个别兼容服务器直接回普通 JSON），则回落为非流式解析。
        /// </summary>
        private async Task<ChatCompletionResult> SendStreamingAsync(
            HttpRequestMessage http,
            Action<AgentProgress> onProgress,
            CancellationToken cancellationToken)
        {
            using var timeoutSource = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeoutSource.CancelAfter(TimeSpan.FromSeconds(180));
            using var response = await httpClient.SendAsync(
                http,
                HttpCompletionOption.ResponseHeadersRead,
                timeoutSource.Token);
            if (!response.IsSuccessStatusCode)
            {
                var errBody = await response.Content.ReadAsStringAsync();
                throw new HttpRequestException(
                    $"模型请求失败：{(int)response.StatusCode} {response.StatusCode}\n{Truncate(errBody)}");
            }

            var content = new StringBuilder();
            var reasoning = new StringBuilder();
            var sawDataLine = false;
            var rawBody = new StringBuilder();
            var toolCalls = new JArray();
            var toolCallIndex = new Dictionary<int, JObject>();
            var totalTokens = 0;

            using var stream = await response.Content.ReadAsStreamAsync();
            using var reader = new StreamReader(stream, Encoding.UTF8);
            string line;
            while ((line = await reader.ReadLineAsync()) != null)
            {
                rawBody.Append(line).Append('\n');
                if (!line.StartsWith("data:", StringComparison.Ordinal))
                {
                    continue;
                }

                var payload = line.Substring(5).Trim();
                if (payload == "[DONE]")
                {
                    break;
                }

                sawDataLine = true;
                JObject chunk;
                try
                {
                    chunk = JObject.Parse(payload);
                }
                catch (JsonException)
                {
                    AppLog.Warn("OpenAiCompatibleClient", "[LogCover] OpenAiCompatibleClient.cs:308  silent-catch covered (no exception variable)");
                    // 单块损坏可跳过，不影响已累积内容。
                    continue;
                }

                var choices = chunk["choices"] as JArray;
                var choice = (choices != null && choices.Count > 0) ? choices[0] as JObject : null;
                var delta = choice?["delta"];
                if (delta != null)
                {
                    var rc = delta.Value<string>("reasoning_content") ?? delta.Value<string>("reasoning");
                    if (!string.IsNullOrWhiteSpace(rc))
                    {
                        reasoning.Append(rc);
                        if (onProgress != null)
                        {
                            onProgress(new AgentProgress
                            {
                                Kind = AgentProgressKind.ReasoningDelta,
                                Reasoning = rc
                            });
                        }
                    }

                    content.Append(delta.Value<string>("content"));

                    var rawCalls = delta["tool_calls"] as JArray;
                    if (rawCalls != null)
                    {
                        foreach (var token in rawCalls.OfType<JObject>())
                        {
                            var idx = token.Value<int?>("index") ?? 0;
                            if (!toolCallIndex.TryGetValue(idx, out var slot))
                            {
                                slot = new JObject
                                {
                                    ["id"] = token.Value<string>("id") ?? string.Empty,
                                    ["type"] = token.Value<string>("type") ?? "function",
                                    ["function"] = new JObject { ["name"] = "", ["arguments"] = "" }
                                };
                                toolCallIndex[idx] = slot;
                                toolCalls.Add(slot);
                            }

                            var fn = token["function"] as JObject;
                            if (fn != null)
                            {
                                var name = fn.Value<string>("name");
                                if (!string.IsNullOrEmpty(name))
                                {
                                    slot["function"]["name"] = name;
                                }

                                var args = fn.Value<string>("arguments");
                                if (args != null)
                                {
                                    slot["function"]["arguments"] =
                                        (slot["function"].Value<string>("arguments") ?? string.Empty) + args;
                                }
                            }
                        }
                    }
                }

                if (chunk["usage"] is JObject usage)
                {
                    totalTokens = usage.Value<int?>("total_tokens") ?? totalTokens;
                }
            }

            var toolsForResult = new List<ToolCallRequest>();
            foreach (var token in toolCalls.OfType<JObject>())
            {
                var call = ToolCallFactory.FromNative(token);
                if (!string.IsNullOrWhiteSpace(call.Name))
                {
                    toolsForResult.Add(call);
                }
            }

            // 若服务器根本没给 SSE data: 行（普通 JSON 响应），直接解析已累积的原文兜底。
            if (!sawDataLine)
            {
                return ParseCompletionBody(rawBody.ToString());
            }

            return new ChatCompletionResult
            {
                Content = content.ToString(),
                Reasoning = reasoning.ToString(),
                ToolCalls = toolsForResult,
                TotalTokens = totalTokens
            };
        }

        private static ChatCompletionResult ParseCompletionBody(string body)
        {
            var payload = TryParse(body);
            var choices = payload?["choices"] as JArray;
            var message = choices?[0]?["message"];
            var usage = payload["usage"];

            var toolCalls = new List<ToolCallRequest>();
            var rawToolCalls = message?["tool_calls"] as JArray;
            if (rawToolCalls != null)
            {
                foreach (var token in rawToolCalls.OfType<JObject>())
                {
                    var call = ToolCallFactory.FromNative(token);
                    if (!string.IsNullOrWhiteSpace(call.Name))
                    {
                        toolCalls.Add(call);
                    }
                }
            }

            return new ChatCompletionResult
            {
                Content = message?.Value<string>("content") ?? string.Empty,
                Reasoning = message?.Value<string>("reasoning_content")
                            ?? message?.Value<string>("reasoning")
                            ?? string.Empty,
                ToolCalls = toolCalls,
                PromptTokens = usage?.Value<int?>("prompt_tokens") ?? 0,
                CompletionTokens = usage?.Value<int?>("completion_tokens") ?? 0,
                TotalTokens = usage?.Value<int?>("total_tokens") ?? 0
            };
        }

        private async Task<ChatCompletionResult> SendAndReadAsync(
            HttpRequestMessage http,
            CancellationToken cancellationToken)
        {
            using var timeoutSource = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeoutSource.CancelAfter(TimeSpan.FromSeconds(180));
            using var response = await httpClient.SendAsync(
                http,
                HttpCompletionOption.ResponseHeadersRead,
                timeoutSource.Token);
            var body = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                throw new HttpRequestException($"模型请求失败：{(int)response.StatusCode} {response.StatusCode}\n{Truncate(body)}");
            }

            var payload = TryParse(body);
            var choices = payload?["choices"] as JArray;
            var message = choices?[0]?["message"];
            var usage = payload["usage"];

            var toolCalls = new List<ToolCallRequest>();
            var rawToolCalls = message?["tool_calls"] as JArray;
            if (rawToolCalls != null)
            {
                foreach (var token in rawToolCalls.OfType<JObject>())
                {
                    var call = ToolCallFactory.FromNative(token);
                    if (!string.IsNullOrWhiteSpace(call.Name))
                    {
                        toolCalls.Add(call);
                    }
                }
            }

            return new ChatCompletionResult
            {
                Content = message?.Value<string>("content") ?? string.Empty,
                Reasoning = message?.Value<string>("reasoning_content")
                            ?? message?.Value<string>("reasoning")
                            ?? string.Empty,
                ToolCalls = toolCalls,
                PromptTokens = usage?.Value<int?>("prompt_tokens") ?? 0,
                CompletionTokens = usage?.Value<int?>("completion_tokens") ?? 0,
                TotalTokens = usage?.Value<int?>("total_tokens") ?? 0
            };
        }

        private static bool Retriable(Exception ex)
        {
            return ex is HttpRequestException ||
                   ex is FormatException ||
                   ex is Newtonsoft.Json.JsonException;
        }

        /// <summary>
        /// .NET Framework 的 HttpWebRequest/HttpClient 在宿主进程（如 Excel）中可能被
        /// 其他组件把 ServicePointManager.SecurityProtocol 降级到 TLS 1.0/1.1，导致
        /// 「未能创建 SSL/TLS 安全通道」；请求前幂等地确保 TLS 1.2 可用。
        /// </summary>
        private static void EnsureTls12()
        {
            var current = ServicePointManager.SecurityProtocol;
            if ((current & SecurityProtocolType.Tls12) == 0)
            {
                ServicePointManager.SecurityProtocol = current | SecurityProtocolType.Tls12;
            }
        }

        private static IReadOnlyList<ChatMessage> WithFallbackInstruction(IReadOnlyList<ChatMessage> messages)
        {
            var list = messages.ToList();
            list.Add(ChatMessage.System(
                "当前模型的工具调用接口不可用，已自动降级。现在只输出一个 JSON 对象，不要输出 Markdown。" +
                "回答用户时输出：{\"action\":\"final\",\"content\":\"...\"}。" +
                "需要操作时输出：{\"action\":\"tool\",\"tool\":\"名称\",\"args\":{},\"reason\":\"...\",\"call_id\":\"唯一ID\"}。" +
                "每次只调用一个工具。"));
            return list;
        }

        public async Task<bool> TestConnectionAsync(CancellationToken cancellationToken)
        {
            try
            {
                var result = await CompleteAsync(
                    new List<ChatMessage> { ChatMessage.User("ping") },
                    cancellationToken);
                return result != null;
            }
            catch (Exception ex) when (
                ex is HttpRequestException ||
                ex is OperationCanceledException ||
                ex is JsonException)
            {
                return false;
            }
        }

        /// <summary>
        /// 按探测结果注入提供方服务端搜索参数（仅 ProviderNative 且非 Fallback 降级重试时注入；
        /// 降级重试跳过注入——若注入参数曾引发 400，重试时不再携带以免死循环）。
        /// </summary>
        private void InjectProviderSearch(JObject request, bool fallbackInstruction)
        {
            var detection = SearchDetection;
            if (detection == null ||
                detection.Capability != SearchProviderDetector.SearchCapability.ProviderNative ||
                fallbackInstruction)
            {
                return;
            }

            SearchRequestInjector.Apply(request, detection.Provider);
            AppLog.Info("search", "已注入服务端搜索参数 provider=" + detection.Provider);
        }

        private string BuildChatCompletionsUrl()
        {
            var baseUri = new Uri(profile.BaseUrl);
            var path = baseUri.AbsolutePath.Trim('/');
            var segments = string.IsNullOrEmpty(path)
                ? new List<string>()
                : new List<string>(path.Split('/'));
            if (segments.Any(s => string.Equals(s, "chat.completions", StringComparison.OrdinalIgnoreCase)))
            {
                // 已含 chat/completions 段（可能自定义完整路径），原样保留。
            }
            else
            {
                if (segments.Count == 0 || !segments.Contains("v1", StringComparer.OrdinalIgnoreCase))
                {
                    // 确保 /v1 版本段存在，兼容「域名+ /v1/chat/completions 或裸域名」。
                    segments.Insert(0, "v1");
                }

                segments.Add("chat");
                segments.Add("completions");
            }

            var builder = new UriBuilder(baseUri)
            {
                Path = "/" + string.Join("/", segments),
                Query = string.Empty
            };
            return builder.Uri.ToString();
        }

        private static JObject TryParse(string value)
        {
            try
            {
                return JObject.Parse(value);
            }
            catch (JsonException)
            {
                throw new FormatException($"模型返回不是 JSON：{Truncate(value)}");
            }
        }

        private static string Truncate(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            value = value.Trim();
            return value.Length <= 500 ? value : value.Substring(0, 500) + "...";
        }

        public void Dispose()
        {
            httpClient?.Dispose();
        }
    }
}
