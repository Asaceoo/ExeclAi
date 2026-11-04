using System.Text;

namespace ExcelAiAssistant.Core.Agent
{
    /// <summary>
    /// 自动自检的元自检（每回合一个实例）：统计「可自检的主写操作数」与「实际执行的自检计划数」。
    /// 如果一个回合里发生了可自检的写操作，却没有一条自检计划被执行，说明自检链路本身失效
    /// （verify_* 工具未注册/非只读、富化失败、AutoVerify 被关等）——这正是最危险的静默退化：
    /// 模型照常写数据，只是没人再核对。回合结束时给出摘要/告警消息，由 AgentRunner 写入运行日志。
    /// </summary>
    public sealed class AutoVerifyMonitor
    {
        /// <summary>发生过自动自检计划的主操作个数（一次批量回写计 1 个操作、可产多条计划）。</summary>
        public int SelfCheckableOps { get; private set; }

        /// <summary>实际执行了的 verify 计划条数。</summary>
        public int ExecutedPlans { get; private set; }

        /// <summary>生成了计划但被跳过（verify 工具不可用/非只读/富化丢弃）的条数。</summary>
        public int SkippedPlans { get; private set; }

        /// <summary>主工具成功且产生了自检计划时调用（无论计划最终是否执行）。</summary>
        public void RecordSelfCheckableOp()
        {
            SelfCheckableOps++;
        }

        /// <summary>记录本条主操作的计划执行结局。</summary>
        public void RecordOutcome(int executed, int skipped)
        {
            ExecutedPlans += executed;
            SkippedPlans += skipped;
        }

        /// <summary>
        /// 回合结束时的摘要：可自检操作存在但一条自检都没执行 → 返回告警（自检链路疑似失效）；
        /// 正常执行 → 返回覆盖率摘要；本回合无可自检操作 → 返回 null（不打日志）。
        /// </summary>
        public string BuildSummary()
        {
            if (SelfCheckableOps == 0)
            {
                return null;
            }

            var builder = new StringBuilder();
            builder.Append("自动自检元监控：").Append(SelfCheckableOps).Append(" 个可自检写操作，执行 ")
                .Append(ExecutedPlans).Append(" 条自检计划");
            if (SkippedPlans > 0)
            {
                builder.Append("，跳过 ").Append(SkippedPlans).Append(" 条");
            }

            if (ExecutedPlans == 0)
            {
                builder.Append("。自检链路疑似失效：一条自检计划都没执行，请检查 verify_* 工具注册、AutoVerify 开关与结果富化逻辑，本回合的写操作全部未经核对。");
            }

            return builder.ToString();
        }
    }
}
