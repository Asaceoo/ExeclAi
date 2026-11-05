using System;
using System.Collections.Generic;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using ExcelAiAssistant.Core.Agent;
using ExcelAiAssistant.Core.Configuration;
using ExcelAiAssistant.Core.Models;

internal static class Program
{
    private static int Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        var prompt = args.Length > 0 ? args[0]
            : "直接回答：你的模型名称是什么？一句话自我介绍。";
        var cfgPath = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelAiAssistant", "models.json");

        var store = new ConfigurationStore(
            cfgPath,
            apiKey => ProtectedData.Protect(
                Encoding.UTF8.GetBytes(apiKey),
                Encoding.UTF8.GetBytes("ExcelAiAssistant:v1"),
                DataProtectionScope.CurrentUser),
            cipher => Encoding.UTF8.GetString(
                ProtectedData.Unprotect(
                    cipher,
                    Encoding.UTF8.GetBytes("ExcelAiAssistant:v1"),
                    DataProtectionScope.CurrentUser)));

        var profiles = store.Load();
        Console.WriteLine("配置文件条目: " + profiles.Count);
        if (profiles.Count == 0) { Console.WriteLine("FAIL: 无配置"); return 2; }

        var profile = profiles[0];
        Console.WriteLine("模型:   " + profile.ModelName);
        Console.WriteLine("BaseUrl:" + profile.BaseUrl);
        Console.WriteLine("协议:   " + profile.ProtocolMode);
        Console.WriteLine("有密钥: " + (profile.ApiKey.Length > 0));
        Console.WriteLine("密钥前缀: " + (profile.ApiKey.Length >= 5 ? profile.ApiKey.Substring(0, 5) + "****" : "短"));

        using (var client = new OpenAiCompatibleClient(profile))
        {
            var system = "你是嵌入 Excel 的操作助手，名叫 XAI。只通过工具操作当前 Excel 工作簿。";
            var messages = new List<ChatMessage>
            {
                ChatMessage.System(system),
                ChatMessage.User(prompt)
            };

            Console.WriteLine();
            Console.WriteLine("=== [A] 通过 OpenAiCompatibleClient 调用 ===");
            var start = DateTime.UtcNow;
            try
            {
                var result = client.CompleteAsync(messages, CancellationToken.None).GetAwaiter().GetResult();
                var elapsed = DateTime.UtcNow - start;
                Console.WriteLine("=== 响应 ===");
                Console.WriteLine((result.Content ?? "(空)").Trim());
                Console.WriteLine();
                Console.WriteLine("工具调用数: " + result.ToolCalls.Count);
                Console.WriteLine("耗时: " + elapsed.TotalSeconds.ToString("0.0") + " 秒");
                bool ok = !string.IsNullOrWhiteSpace(result.Content);
                Console.WriteLine(ok ? "PASS: 模型对话返回正常" : "FAIL: 内容为空");
                return ok ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAIL(client): " + ex.Message);
                Console.WriteLine(ex.GetType().Name);

                Console.WriteLine();
                Console.WriteLine("=== [B] 用 HttpClient 最小请求复现（对比）===");
                ProbeHttp(profile);
                return 1;
            }
        }
    }

    private static void ProbeHttp(ModelProfile profile)
    {
        using (var hc = new System.Net.Http.HttpClient())
        {
            hc.Timeout = Timeout.InfiniteTimeSpan;
            hc.DefaultRequestHeaders.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", profile.ApiKey);
            var body = "{\"model\":\"" + profile.ModelName +
                       "\",\"temperature\":" + profile.Temperature.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                       ",\"messages\":[{\"role\":\"system\",\"content\":\"你是嵌入Excel的操作助手，名叫XAI。\"},{\"role\":\"user\",\"content\":\"hi\"}]}";
            Console.WriteLine("请求体: " + body);
            var req = new System.Net.Http.HttpRequestMessage(
                System.Net.Http.HttpMethod.Post,
                "https://token.sensenova.cn/v1/chat/completions");
            var encBom = new UTF8Encoding(true);
            var encNoBom = new UTF8Encoding(false);
            var bytesBom = encBom.GetBytes(body);
            var bytesNoBom = encNoBom.GetBytes(body);
            Console.WriteLine("BOM?   " + (bytesBom.Length != bytesNoBom.Length ? "带BOM" : "无BOM"));
            Console.WriteLine("正文长(带BOM): " + bytesBom.Length + " (无BOM): " + bytesNoBom.Length);
            var content = new System.Net.Http.ByteArrayContent(bytesNoBom);
            content.Headers.ContentType =
                new System.Net.Http.Headers.MediaTypeHeaderValue("application/json");
            req.Content = content;
            Console.WriteLine("--- 实际发送头 (合并 Default+Request) ---");
            foreach (var h in hc.DefaultRequestHeaders) Console.WriteLine("D " + h.Key + ": " + string.Join(",", h.Value));
            foreach (var h in req.Headers) Console.WriteLine("R " + h.Key + ": " + string.Join(",", h.Value));
            foreach (var h in req.Content.Headers) Console.WriteLine("C " + h.Key + ": " + string.Join(",", h.Value));
            var resp = hc.SendAsync(req).GetAwaiter().GetResult();
            Console.WriteLine("HTTP " + (int)resp.StatusCode);
            Console.WriteLine(resp.Content.ReadAsStringAsync().GetAwaiter().GetResult());

            Console.WriteLine();
            Console.WriteLine("=== [D] 用 HttpWebRequest（老栈，等价 curl --data-binary）===");
            var request = (HttpWebRequest)WebRequest.Create("https://token.sensenova.cn/v1/chat/completions");
            request.Method = "POST";
            request.ContentType = "application/json";
            request.ContentLength = bytesNoBom.Length;
            request.Headers["Authorization"] = "Bearer " + profile.ApiKey;
            request.UserAgent = "XAI-Excel/1.0";
            using (var s = request.GetRequestStream())
            {
                s.Write(bytesNoBom, 0, bytesNoBom.Length);
            }
            try
            {
                using (var response = (HttpWebResponse)request.GetResponse())
                using (var reader = new System.IO.StreamReader(response.GetResponseStream()))
                {
                    Console.WriteLine("HTTP " + (int)response.StatusCode);
                    Console.WriteLine(reader.ReadToEnd());
                }
            }
            catch (WebException we)
            {
                var failResp = (HttpWebResponse)we.Response;
                Console.WriteLine("HttpWebRequest FAIL: " + (int)failResp.StatusCode);
                using (var reader = new System.IO.StreamReader(failResp.GetResponseStream()))
                {
                    Console.WriteLine(reader.ReadToEnd());
                }
            }
        }
    }
}