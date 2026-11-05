using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace ExcelAiAssistant.Core.Agent
{
    /// <summary>
    /// 面向用户展示的 token 估算。确定性近似：CJK 汉字按约 1 token，其余按约 4 字符/token。
    /// </summary>
    public static class TokenEstimator
    {
        public static int Estimate(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            var cjk = 0;
            var other = 0;
            foreach (var character in text)
            {
                if ((character >= '\u4e00' && character <= '\u9fff') ||
                    (character >= '\u3000' && character <= '\u303f'))
                {
                    cjk++;
                }
                else
                {
                    other++;
                }
            }

            return cjk + (int)Math.Ceiling(other / 4.0);
        }

        public static int Estimate(IEnumerable<ChatMessage> messages)
        {
            if (messages == null)
            {
                return 0;
            }

            return messages.Sum(message =>
                Estimate(message.Content) + (message.Role?.Length ?? 0) + 4);
        }

        public static int EstimateJson(object payload)
        {
            if (payload == null)
            {
                return 0;
            }

            return Estimate(JsonConvert.SerializeObject(payload));
        }
    }

    public sealed class SensitiveDetection
    {
        public string Kind { get; set; }

        public string Label { get; set; }

        public string Sample { get; set; }

        public int Count { get; set; }

        public string Category => SensitiveCategoryLabel(Kind);

        internal static string SensitiveCategoryLabel(string kind)
        {
            switch (kind)
            {
                case "email": return "邮箱地址";
                case "cn_mobile": return "中国大陆手机号";
                case "cn_id": return "身份证号";
                case "bank_card": return "银行卡号";
                default: return kind;
            }
        }
    }

    /// <summary>
    /// 敏感信息检测：用于发送前提示用户"将要发送的数据里含有哪些敏感内容"。
    /// 仅做启发式扫描，不保证完整覆盖；结果用于提醒而非硬性阻断。
    /// </summary>
    public static class SensitivityScanner
    {
        private static readonly Regex EmailPattern = new Regex(
            @"\b[\w.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)*\b",
            RegexOptions.Compiled);

        private static readonly Regex CnMobilePattern = new Regex(
            @"(?<!\d)1[3-9]\d{9}(?!\d)",
            RegexOptions.Compiled);

        private static readonly Regex CnIdPattern = new Regex(
            @"(?<!\d)[1-9]\d{5}(?:18|19|20)\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])\d{3}[\dXx](?!\d)",
            RegexOptions.Compiled);

        private static readonly Regex BankCardPattern = new Regex(
            @"(?<!\d)(?:[4-9]\d{11,18})(?!\d)",
            RegexOptions.Compiled);

        public static IReadOnlyList<SensitiveDetection> Scan(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return Array.Empty<SensitiveDetection>();
            }

            var results = new List<SensitiveDetection>();
            Add(results, "email", "邮箱地址", EmailPattern, text);
            Add(results, "cn_mobile", "中国大陆手机号", CnMobilePattern, text);
            Add(results, "cn_id", "身份证号", CnIdPattern, text);
            Add(results, "bank_card", "银行卡号", BankCardPattern, text);
            return results;
        }

        private static void Add(
            List<SensitiveDetection> results,
            string kind,
            string label,
            Regex pattern,
            string text)
        {
            int count = 0;
            string sample = null;
            foreach (Match match in pattern.Matches(text))
            {
                count++;
                sample ??= match.Value;
            }

            if (count > 0)
            {
                results.Add(new SensitiveDetection
                {
                    Kind = kind,
                    Label = label,
                    Sample = sample,
                    Count = count
                });
            }
        }
    }

    /// <summary>
    /// 把底层工具/接口错误码转成面向用户的中文提示。
    /// </summary>
    public static class FriendlyError
    {
        public static string For(string errorCode, string fallback)
        {
            if (string.IsNullOrWhiteSpace(errorCode))
            {
                return fallback ?? "操作失败，请重试。";
            }

            switch (errorCode)
            {
                case "USER_DENIED":
                    return "您已取消该操作，未做任何更改。";
                case "UNKNOWN_TOOL":
                    return "当前模型调用了不存在的工具，请重试或换一种表达。";
                case "UNDO_UNAVAILABLE":
                    return "本会话未启用操作历史，无法回滚。";
                case "UNDO_FAILED":
                    return "回滚失败：目标可能已被改动或记录不可用。";
                case "UNDO_WORKBOOK_MISMATCH":
                    return "该操作属于其他工作簿，无法在当前工作簿回滚。";
                case "OPERATION_NOT_FOUND":
                    return "没有找到对应的操作记录。";
                case "INVALID_ADDRESS":
                    return "地址格式不正确，请检查后再试。";
                case "SHAPE_MISMATCH":
                    return "目标区域行列数与数据不匹配，请调整后重试。";
                case "FILE_NOT_FOUND":
                    return "找不到指定的文件，请确认路径后重试。";
                case "READ_FAILED":
                    return "文件读取失败，可能是编码或权限问题。";
                case "EMPTY_FILE":
                    return "文件内容为空，无法处理。";
                case "TOO_MANY_CELLS":
                    return "数据量超过单次处理上限，请分批导入。";
                case "NO_WRITER":
                    return "当前宿主缺少写入能力，无法执行导入。";
                case "CHART_NOT_FOUND":
                    return "没有找到要导出的图表。";
                case "INVALID_COLUMN":
                    return "指定的列号超出数据范围。";
                case "EXCEL_COM_ERROR":
                    return "Excel 操作出错，请检查选区和数据后重试。";
                default:
                    return fallback ?? "操作失败，请重试。";
            }
        }
    }
}