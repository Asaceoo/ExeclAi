using System.Linq;
using ExcelAiAssistant.Core.Diagnostics;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>踩坑失败模式库：日志签名 → 已知坑识别与手册级建议。</summary>
    public sealed class PitfallLibraryTests
    {
        [Fact]
        public void Matches_vba_probe_transient_hresult()
        {
            var hits = PitfallLibrary.Match("第一轮 Application.Run 失败：0x800A9C68，第二轮成功");
            var hit = Assert.Single(hits);
            Assert.Contains("3.6", hit.Rule.Id);
            Assert.Equal("0x800A9C68", hit.MatchedSignature);
        }

        [Fact]
        public void Match_is_case_insensitive()
        {
            Assert.Contains(PitfallLibrary.Match("probe failed with 0x800a9c68"),
                h => h.Rule.Id.Contains("3.6"));
        }

        [Fact]
        public void Matches_macro_disabled_and_rcw_and_threading()
        {
            Assert.Contains(PitfallLibrary.Match("错误码 VBA_MACROS_DISABLED，请启用宏"),
                h => h.Rule.Id.Contains("3.4"));
            Assert.Contains(PitfallLibrary.Match("System.Runtime.InteropServices.COMException: COM 对象与其基础 RCW 分开"),
                h => h.Rule.Id.Contains("1.1"));
            Assert.Contains(PitfallLibrary.Match("CoreWebView2 can only be accessed from the UI thread"),
                h => h.Rule.Id.Contains("3.1"));
        }

        [Fact]
        public void Confirmed_timeout_is_not_reported_as_user_denial()
        {
            var hits = PitfallLibrary.Match("CONFIRM_TIMEOUT：确认超时");
            Assert.Contains(hits, h => h.Rule.Id.Contains("确认超时"));
        }

        [Fact]
        public void Empty_and_unknown_text_match_nothing()
        {
            Assert.Empty(PitfallLibrary.Match(null));
            Assert.Empty(PitfallLibrary.Match(string.Empty));
            Assert.Empty(PitfallLibrary.Match("一切正常 ok=True"));
        }

        [Fact]
        public void One_hit_per_rule_even_with_multiple_signatures()
        {
            var hits = PitfallLibrary.Match("cannot run the macro；may not be available");
            Assert.Equal(1, hits.Count(h => h.Rule.Id.Contains("3.6")));
        }

        [Fact]
        public void All_rules_carry_id_symptom_advice_and_section()
        {
            Assert.NotEmpty(PitfallLibrary.AllRules);
            foreach (var rule in PitfallLibrary.AllRules)
            {
                Assert.False(string.IsNullOrWhiteSpace(rule.Id));
                Assert.False(string.IsNullOrWhiteSpace(rule.Symptom));
                Assert.False(string.IsNullOrWhiteSpace(rule.Advice));
                Assert.False(string.IsNullOrWhiteSpace(rule.DocSection));
                Assert.NotEmpty(rule.Signatures);
            }
        }

        [Fact]
        public void Format_hits_render_advice_lines()
        {
            var hits = PitfallLibrary.Match("0x800A9C68");
            var text = PitfallLibrary.FormatHits(hits);
            Assert.Contains("[", text);
            Assert.Contains("建议：", text);
            Assert.Equal(string.Empty, PitfallLibrary.FormatHits(System.Array.Empty<PitfallHit>()));
        }
    }
}
