using System;

namespace ExcelAiAssistant.Core.Models
{
    public sealed class ModelProfile
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        public string Name { get; set; }

        public string BaseUrl { get; set; }

        public string ApiKey { get; set; }

        public string ModelName { get; set; }

        public double Temperature { get; set; } = 0.1;

        public int? MaxTokens { get; set; }

        public AgentProtocolMode ProtocolMode { get; set; } = AgentProtocolMode.Auto;

        /// <summary>接口认证方式。主流模型用 <see cref="AuthScheme.Bearer"/>（Authorization: Bearer），
        /// 小红书 Dots 等自定义端点要求 <see cref="AuthScheme.ApiKey"/>（api-key 请求头）。</summary>
        public AuthScheme AuthScheme { get; set; } = AuthScheme.Bearer;

        public void Validate()
        {
            if (string.IsNullOrWhiteSpace(Name))
            {
                throw new ArgumentException("方案名称不能为空。");
            }

            if (string.IsNullOrWhiteSpace(BaseUrl) || !Uri.TryCreate(BaseUrl, UriKind.Absolute, out var uri))
            {
                throw new ArgumentException("Base URL 必须是有效绝对地址。");
            }

            if (uri.Scheme != Uri.UriSchemeHttps && uri.Scheme != Uri.UriSchemeHttp)
            {
                throw new ArgumentException("Base URL 仅支持 HTTP/HTTPS。");
            }

            if (string.IsNullOrWhiteSpace(ApiKey))
            {
                throw new ArgumentException("API Key 不能为空。");
            }

            if (string.IsNullOrWhiteSpace(ModelName))
            {
                throw new ArgumentException("模型名称不能为空。");
            }

            if (Temperature < 0 || Temperature > 2)
            {
                throw new ArgumentException("Temperature 必须在 0 到 2 之间。");
            }
        }
    }

    public enum AgentProtocolMode
    {
        Auto,
        Native,
        Fallback
    }

    public enum AuthScheme
    {
        /// <summary>Authorization: Bearer &lt;key&gt;（OpenAI/DeepSeek/通义/Kimi 等主流默认）。</summary>
        Bearer,

        /// <summary>api-key: &lt;key&gt;（小红书 Dots 等自定义端点要求）。</summary>
        ApiKey
    }
}
