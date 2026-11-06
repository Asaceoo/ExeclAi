using ExcelAiAssistant.Excel;
using Xunit;

namespace ExcelAiAssistant.Tests
{
    /// <summary>Python 库目录（PythonEnvService.Catalog）与缺模块自愈白名单（ScriptRunnerService.MissingModulePipMap）
    /// 的登记面契约（docs/需求-Python高级数据分析.md FR-P1）：
    /// 库目录 24 库；statsmodels/wordcloud/jieba 三库必须同时可被发现（Catalog）且缺模块自愈可兜底
    /// （MissingModulePipMap），防止"目录能装、缺模块不自愈"的单边缺口。</summary>
    public class PythonLibCatalogTests
    {
        [Fact]
        public void Catalog_has_25_libraries_with_nonempty_fields()
        {
            Assert.Equal(25, PythonEnvService.Catalog.Count);
            foreach (var kv in PythonEnvService.Catalog)
            {
                Assert.False(string.IsNullOrWhiteSpace(kv.Key), "Catalog 存在空 pip 包名键");
                Assert.False(string.IsNullOrWhiteSpace(kv.Value.ImportName), $"Catalog[{kv.Key}] 导入名为空");
                Assert.False(string.IsNullOrWhiteSpace(kv.Value.Description), $"Catalog[{kv.Key}] 用途说明为空");
            }
        }

        [Fact]
        public void P1_new_libraries_present_in_catalog_and_selfheal_map()
        {
            foreach (var pkg in new[] { "statsmodels", "wordcloud", "jieba", "snownlp" })
            {
                Assert.True(PythonEnvService.Catalog.ContainsKey(pkg), $"库目录缺少 {pkg}");
                Assert.True(ScriptRunnerService.MissingModulePipMap.ContainsKey(pkg), $"缺模块自愈白名单缺少 {pkg}");
            }

            // 导入名契约：面板/教程按导入名转述，AI 生成代码用导入名 import，三者必须与官方导入名一致。
            Assert.Equal("statsmodels", PythonEnvService.Catalog["statsmodels"].ImportName);
            Assert.Equal("wordcloud", PythonEnvService.Catalog["wordcloud"].ImportName);
            Assert.Equal("jieba", PythonEnvService.Catalog["jieba"].ImportName);
            Assert.Equal("snownlp", PythonEnvService.Catalog["snownlp"].ImportName);

            // 自愈映射的 pip 包名必须与目录 pip 包名一致（避免"装了 A 报缺 B"错位）。
            Assert.Equal("statsmodels", ScriptRunnerService.MissingModulePipMap["statsmodels"]);
            Assert.Equal("wordcloud", ScriptRunnerService.MissingModulePipMap["wordcloud"]);
            Assert.Equal("jieba", ScriptRunnerService.MissingModulePipMap["jieba"]);
            Assert.Equal("snownlp", ScriptRunnerService.MissingModulePipMap["snownlp"]);
        }

        [Fact]
        public void Every_catalog_import_name_is_selfheal_resolvable()
        {
            // 目录中每个库的导入名都必须能被缺模块自愈解析兜住（P2 补齐 plotly/xlwings/openpyxl_image_loader 历史盲区后全覆盖），
            // 否则用户装了库仍会因自愈盲区反复报"缺模块"。
            foreach (var kv in PythonEnvService.Catalog)
            {
                Assert.True(
                    ScriptRunnerService.MissingModulePipMap.ContainsKey(kv.Value.ImportName),
                    $"导入名 {kv.Value.ImportName}（pip 包 {kv.Key}）不在缺模块自愈白名单，装后仍会反复报缺模块");
            }
        }
    }
}
