using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
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
        var newKey = args.Length > 0 ? args[0] : string.Empty;
        if (string.IsNullOrWhiteSpace(newKey))
        {
            Console.WriteLine("ERROR 缺少新 API key 参数。用法: UpdateXhsKey <new-key>");
            return 2;
        }

        var cfgPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelAiAssistant", "models.json");
        Console.WriteLine("配置文件: " + cfgPath);

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

        var profiles = new List<ModelProfile>(store.Load());
        Console.WriteLine("现有配置数: " + profiles.Count);

        var target = profiles.FirstOrDefault(p =>
            p.BaseUrl != null &&
            p.BaseUrl.IndexOf("askdiandian", StringComparison.OrdinalIgnoreCase) >= 0);
        if (target == null)
        {
            Console.WriteLine("ERROR 未找到小红书(askdiandian)模型方案。现有方案：");
            foreach (var p in profiles)
            {
                Console.WriteLine("  - " + p.Name + " | " + p.ModelName + " | " + p.BaseUrl);
            }
            return 1;
        }

        Console.WriteLine("目标方案: " + target.Name + " | " + target.ModelName + " | " + target.BaseUrl);
        Console.WriteLine("旧 key 掩码: " + Mask(target.ApiKey));

        // 1) 用存储的旧 key 实测端点，确认根因
        var oldStatus = Probe(target, target.ApiKey);
        Console.WriteLine("旧 key 连接结果: " + oldStatus);

        // 2) 写回新 key
        target.ApiKey = newKey;
        store.Save(profiles);
        Console.WriteLine("已写回新 key，掩码: " + Mask(newKey));

        // 3) 用新 key 实测端点，闭环验证
        var newStatus = Probe(target, newKey);
        Console.WriteLine("新 key 连接结果: " + newStatus);

        return newStatus == "OK" ? 0 : 1;
    }

    private static string Probe(ModelProfile profile, string apiKey)
    {
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(40));
            var testProfile = new ModelProfile
            {
                Name = profile.Name,
                BaseUrl = profile.BaseUrl,
                ModelName = profile.ModelName,
                Temperature = profile.Temperature,
                MaxTokens = profile.MaxTokens,
                ProtocolMode = profile.ProtocolMode,
                AuthScheme = profile.AuthScheme,
                ApiKey = apiKey
            };
            using var client = new OpenAiCompatibleClient(testProfile);
            var ok = client.TestConnectionAsync(cts.Token).GetAwaiter().GetResult();
            return ok ? "OK" : "FAIL";
        }
        catch (Exception ex)
        {
            return "EXCEPTION: " + ex.Message;
        }
    }

    private static string Mask(string key)
    {
        if (string.IsNullOrEmpty(key))
        {
            return "(空)";
        }
        if (key.Length <= 8)
        {
            return "****";
        }
        return key.Substring(0, 4) + "****" + key.Substring(key.Length - 4);
    }
}
