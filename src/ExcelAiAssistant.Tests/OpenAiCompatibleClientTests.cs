using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Models;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>HttpListener 类测试串行执行：避免跨类并行时的端口竞争抖动。</summary>
    [CollectionDefinition("HttpListenerSerial", DisableParallelization = true)]
    public sealed class HttpListenerSerialCollection
    {
    }

    [Collection("HttpListenerSerial")]
    public sealed class OpenAiCompatibleClientTests : IDisposable
    {
        private readonly HttpListener listener;
        private readonly string baseUrl;

        public OpenAiCompatibleClientTests()
        {
            // GetFreePort 探测与 HttpListener 绑定之间存在竞窗（并行测试偶发抢占端口），
            // 绑定失败换端口重试，最多 8 次——消除门禁中的偶发红测。
            for (var attempt = 0; ; attempt++)
            {
                var port = GetFreePort();
                baseUrl = $"http://127.0.0.1:{port}/v1/";
                listener = new HttpListener();
                listener.Prefixes.Add(baseUrl);
                try
                {
                    listener.Start();
                    break;
                }
                catch (HttpListenerException) when (attempt < 7)
                {
                    // 端口被并行测试抢占：换端口重试。
                }
            }
        }

        [Fact]
        public async Task Parses_chat_completion_response_and_appends_chat_completions_path()
        {
            var responseTask = RespondOnce(
                "{\"choices\":[{\"message\":{\"content\":\"{\\\"action\\\":\\\"final\\\",\\\"content\\\":\\\"pong\\\"}\"}}]," +
                "\"usage\":{\"prompt_tokens\":1,\"completion_tokens\":2,\"total_tokens\":3}}");
            using var client = CreateClient($"{baseUrl}");
            var result = await client.CompleteAsync(
                new List<ChatMessage> { ChatMessage.User("hello") },
                CancellationToken.None);

            var requestPath = await responseTask;
            Assert.EndsWith("/v1/chat/completions", requestPath);
            Assert.Equal("{\"action\":\"final\",\"content\":\"pong\"}", result.Content);
            Assert.Equal(3, result.TotalTokens);
        }

        [Fact]
        public async Task Streams_reasoning_deltas_incrementally()
        {
            // 流式 SSE 响应：三段 reasoning_content 增量 + 内容 dE 与 usage。客户端必须：
            //  1) 每个 reasoning 增量都触发一次 onProgress(ReasoningDelta)；
            //  2) 累计后的最终 Reasoning 为三段拼接。
            var sse = "data: {\"choices\":[{\"delta\":{\"reasoning_content\":\"TokenA \"}}]}\n\n" +
                      "data: {\"choices\":[{\"delta\":{\"reasoning_content\":\"TokenB \"}}]}\n\n" +
                      "data: {\"choices\":[{\"delta\":{\"reasoning_content\":\"TokenC\"}}]}\n\n" +
                      "data: {\"choices\":[{\"delta\":{\"content\":\"done\"}}]}\n\n" +
                      "data: {\"choices\":[],\"usage\":{\"total_tokens\":7}}\n\n" +
                      "data: [DONE]\n\n";
            _ = RespondSse(sse);

            var deltas = new List<string>();
            Action<AgentProgress> onProgress = p =>
            {
                if (p.Kind == AgentProgressKind.ReasoningDelta)
                {
                    deltas.Add(p.Reasoning);
                }
            };

            using var client = CreateClient(baseUrl);
            var result = await client.CompleteAsync(
                new List<ChatMessage> { ChatMessage.User("hi") },
                null,
                CancellationToken.None,
                onProgress);

            // 三次增量都实时上报（流式核心）。
            Assert.Equal(new[] { "TokenA ", "TokenB ", "TokenC" }, deltas);
            // 累计结果正确带完整推理 + 正文。
            Assert.Equal("TokenA TokenB TokenC", result.Reasoning);
            Assert.Equal("done", result.Content);
            Assert.Equal(7, result.TotalTokens);
        }

        [Fact]
        public async Task Test_connection_reports_http_failure()
        {
            _ = RespondOnce("{\"error\":\"bad key\"}", HttpStatusCode.Unauthorized);
            using var client = CreateClient(baseUrl);

            Assert.False(await client.TestConnectionAsync(CancellationToken.None));
        }

        [Fact]
        public async Task Bare_domain_base_url_appends_v1_chat_completions()
        {
            var captured = new TaskCompletionSource<string>();
            var handler = new CaptureHandler(request =>
            {
                captured.TrySetResult(request.RequestUri.ToString());
                return new HttpResponseMessage(System.Net.HttpStatusCode.OK)
                {
                    Content = new StringContent(
                        "{\"choices\":[{\"message\":{\"content\":\"ok\"}}],\"usage\":{\"total_tokens\":1}}",
                        Encoding.UTF8, "application/json")
                };
            });
            using var http = new HttpClient(handler, disposeHandler: true);
            using var client = new OpenAiCompatibleClient(new ModelProfile
            {
                Name = "Bare",
                BaseUrl = "https://note3-prev-api.askdiandian.com",
                ApiKey = "ak-test",
                ModelName = "dots3-note-prev",
                AuthScheme = AuthScheme.ApiKey
            }, http);

            var result = await client.CompleteAsync(
                new List<ChatMessage> { ChatMessage.User("hi") },
                CancellationToken.None);

            Assert.Equal("ok", result.Content);
            Assert.Equal(
                "https://note3-prev-api.askdiandian.com/v1/chat/completions",
                await captured.Task);
        }

        [Fact]
        public void ApiKey_auth_scheme_uses_api_key_header()
        {
            using var http = new HttpClient();
            using var client = new OpenAiCompatibleClient(new ModelProfile
            {
                Name = "Dots",
                BaseUrl = "https://note3-prev-api.askdiandian.com",
                ApiKey = "ak-secret",
                ModelName = "dots3-note-prev",
                AuthScheme = AuthScheme.ApiKey
            }, http);

            Assert.True(http.DefaultRequestHeaders.Contains("api-key"));
            Assert.Equal("ak-secret", string.Join(",", http.DefaultRequestHeaders.GetValues("api-key")));
            Assert.Null(http.DefaultRequestHeaders.Authorization);
        }

        [Fact]
        public void Bearer_auth_scheme_is_default_and_uses_authorization()
        {
            using var http = new HttpClient();
            using var client = new OpenAiCompatibleClient(new ModelProfile
            {
                Name = "OpenAI",
                BaseUrl = "https://api.openai.com/v1",
                ApiKey = "sk-x",
                ModelName = "gpt-4o-mini",
                AuthScheme = AuthScheme.Bearer
            }, http);

            Assert.False(http.DefaultRequestHeaders.Contains("api-key"));
            Assert.Equal("Bearer", http.DefaultRequestHeaders.Authorization.Scheme);
            Assert.Equal("sk-x", http.DefaultRequestHeaders.Authorization.Parameter);
        }

        private sealed class CaptureHandler : HttpMessageHandler
        {
            private readonly Func<HttpRequestMessage, HttpResponseMessage> responder;

            public CaptureHandler(Func<HttpRequestMessage, HttpResponseMessage> responder)
            {
                this.responder = responder;
            }

            protected override Task<HttpResponseMessage> SendAsync(
                HttpRequestMessage request, CancellationToken cancellationToken)
            {
                return Task.FromResult(responder(request));
            }
        }

        private async Task<string> RespondOnce(string body, HttpStatusCode status = HttpStatusCode.OK)
        {
            var context = await listener.GetContextAsync();
            context.Response.StatusCode = (int)status;
            context.Response.ContentType = "application/json";
            var bytes = Encoding.UTF8.GetBytes(body);
            await context.Response.OutputStream.WriteAsync(bytes, 0, bytes.Length);
            context.Response.Close();
            return context.Request.Url.AbsolutePath;
        }

        private Task<string> RespondSse(string body, HttpStatusCode status = HttpStatusCode.OK)
        {
            return RespondOnce(body, status);
        }

        private OpenAiCompatibleClient CreateClient(string url)
        {
            return new OpenAiCompatibleClient(new ModelProfile
            {
                Name = "Test",
                BaseUrl = url,
                ApiKey = "test-key",
                ModelName = "test-model"
            });
        }

        private static int GetFreePort()
        {
            var socket = new System.Net.Sockets.Socket(
                System.Net.Sockets.AddressFamily.InterNetwork,
                System.Net.Sockets.SocketType.Stream,
                System.Net.Sockets.ProtocolType.Tcp);
            socket.Bind(new IPEndPoint(IPAddress.Loopback, 0));
            var port = ((IPEndPoint)socket.LocalEndPoint).Port;
            socket.Dispose();
            return port;
        }

        public void Dispose()
        {
            listener.Close();
        }
    }
}
