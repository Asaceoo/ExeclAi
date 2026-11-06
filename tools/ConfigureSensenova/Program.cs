using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using ExcelAiAssistant.Core.Configuration;
using ExcelAiAssistant.Core.Models;

internal static class Program
{
    private static int Main(string[] args)
    {
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

        var existing = new List<ModelProfile>(store.Load());
        Console.WriteLine("现有配置数: " + existing.Count);

        existing.RemoveAll(p =>
            string.Equals(p.Name, "22", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(p.BaseUrl, "http://localhost:8000/v1/chat/completions",
                StringComparison.OrdinalIgnoreCase));

        var profile = new ModelProfile
        {
            Name = "sensenova",
            BaseUrl = "https://token.sensenova.cn/v1",
            ApiKey = args.Length > 0 ? args[0] : "sk-LWxQ5BTj8qy5Bu3KReYmNwSW6CA4M3jD",
            ModelName = "sensenova-6.8-flash-lite",
            Temperature = 0.1,
            ProtocolMode = AgentProtocolMode.Auto
        };

        existing.Add(profile);
        store.Save(existing);
        Console.WriteLine("已保存 sensenova 配置，共 " + existing.Count + " 条。");
        return 0;
    }
}