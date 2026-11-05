using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
        var cfgPath = Path.Combine(
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
        var target = store.Load().FirstOrDefault(p =>
            p.BaseUrl != null && p.BaseUrl.IndexOf("askdiandian", StringComparison.OrdinalIgnoreCase) >= 0);
        if (target == null)
        {
            Console.WriteLine("ERROR 未找到小红书方案");
            return 1;
        }

        Console.WriteLine("当前默认 SecurityProtocol: " + ServicePointManager.SecurityProtocol);

        // 阶段1：强制仅 TLS 1.0，复现「未能创建 SSL/TLS 安全通道」
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
        Console.WriteLine("阶段1(强制 TLS1.0) SecurityProtocol: " + ServicePointManager.SecurityProtocol);
        Console.WriteLine("阶段1 结果: " + Probe(target));

        // 阶段2：加入 TLS1.2（模拟修复后），应恢复
        ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
        Console.WriteLine("阶段2(+Tls12) SecurityProtocol: " + ServicePointManager.SecurityProtocol);
        Console.WriteLine("阶段2 结果: " + Probe(target));

        return 0;
    }

    private static string Probe(ModelProfile profile)
    {
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(40));
            using var client = new OpenAiCompatibleClient(profile);
            var ok = client.TestConnectionAsync(cts.Token).GetAwaiter().GetResult();
            return ok ? "OK" : "FAIL";
        }
        catch (Exception ex)
        {
            return "EXCEPTION: " + ex.GetType().Name + ": " + ex.Message;
        }
    }
}
