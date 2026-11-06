using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelAiAssistant.Core.Models;
using ExcelAiAssistant.Core.Tools;

namespace ExcelAiAssistant.Core.Agent
{
    /// <summary>
    /// 模型补全客户端抽象：AgentRunner 只依赖此接口，不依赖具体的 OpenAI 兼容实现。
    /// 测试可注入脚本化的假客户端，驱动「工具调用 → 自动自检 → 最终回复」的完整回合编排做回归断言。
    /// </summary>
    public interface IChatCompletionClient : IDisposable
    {
        Task<ChatCompletionResult> CompleteAsync(
            IReadOnlyList<ChatMessage> messages,
            IReadOnlyList<ToolDescriptor> tools,
            CancellationToken cancellationToken,
            Action<AgentProgress> onProgress);
    }
}
