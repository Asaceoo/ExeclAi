using Newtonsoft.Json.Linq;

namespace ExcelAiAssistant.Core.Search
{
    /// <summary>
    /// 提供方服务端搜索参数注入（纯逻辑可单测）：按探测到的 provider 把对应搜索参数
    /// 写入 chat/completions 请求体。OpenAiCompatibleClient 在发送前调用；
    /// Fallback(JSON) 降级重试时不注入（由调用方控制）。
    /// </summary>
    public static class SearchRequestInjector
    {
        public static void Apply(JObject request, string provider)
        {
            if (request == null || string.IsNullOrWhiteSpace(provider))
            {
                return;
            }

            switch (provider)
            {
                case "qwen":
                    // 通义 DashScope compatible-mode：enable_search 开关。
                    request["enable_search"] = true;
                    break;

                case "zhipu":
                    // 智谱开放平台：web_search 工具声明（独立 tool 类型，与 function 工具并存）。
                    var zhipuTools = request["tools"] as JArray ?? new JArray();
                    zhipuTools.Add(new JObject
                    {
                        ["type"] = "web_search",
                        ["web_search"] = new JObject
                        {
                            ["enable"] = true,
                            ["search_engine"] = "search_std_jina"
                        }
                    });
                    request["tools"] = zhipuTools;
                    break;

                case "moonshot":
                    // Kimi/Moonshot：builtin_function 形态的 $web_search；若服务端未自动执行，
                    // AgentRunner 会拦截该调用并走本地 bing 抓取兜底回传真实结果。
                    var moonshotTools = request["tools"] as JArray ?? new JArray();
                    moonshotTools.Add(new JObject
                    {
                        ["type"] = "builtin_function",
                        ["function"] = new JObject { ["name"] = "$web_search" }
                    });
                    request["tools"] = moonshotTools;
                    break;

                case "openrouter":
                    // OpenRouter：模型名 :online 后缀启用联网插件。
                    var model = request.Value<string>("model");
                    if (!string.IsNullOrEmpty(model) &&
                        model.IndexOf(":online", System.StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        request["model"] = model + ":online";
                    }
                    break;

                case "qianfan":
                    // 百度千帆 v2：web_search 组件开关。
                    request["web_search"] = new JObject { ["enable"] = true };
                    break;

                default:
                    break;
            }
        }
    }
}
