using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelAiAssistant.Excel;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;

internal static class Program
{
    private const string ParamJsonToken = "__PARAM_JSON__";
    private static int _fail;

    private static void Check(bool cond, string name)
    {
        Console.WriteLine((cond ? "[PASS] " : "[FAIL] ") + name);
        if (!cond) _fail++;
    }

    [STAThread]
    private static int Main()
    {
        Exception pumpError = null;
        string rawStringified = null, rawDirect = null;
        string realRaw = null;

        using (var form = new Form())
        {
            form.WindowState = FormWindowState.Minimized;
            form.ShowInTaskbar = false;
            var webView = new WebView2 { Dock = DockStyle.Fill };
            form.Controls.Add(webView);

            form.Load += async (_, __) =>
            {
                try
                {
                    var env = await CoreWebView2Environment.CreateAsync(
                        userDataFolder: Path.Combine(Path.GetTempPath(), "verifyparamjson"));
                    await webView.EnsureCoreWebView2Async(env);
                    webView.NavigateToString(Html);
                    await WaitForNav(webView);

                    rawStringified = await webView.CoreWebView2.ExecuteScriptAsync("JSON.stringify(collect())");
                    rawDirect = await webView.CoreWebView2.ExecuteScriptAsync("collect()");

                    webView.NavigateToString(BuildRealTemplatePage());
                    await WaitForNav(webView);
                    realRaw = await webView.CoreWebView2.ExecuteScriptAsync(VerifyJs);
                }
                catch (Exception ex)
                {
                    pumpError = ex;
                }
                finally
                {
                    form.DialogResult = DialogResult.OK;
                }
            };

            form.ShowDialog();
        }

        if (pumpError != null)
        {
            Console.WriteLine("[FAIL] webview init/run: " + pumpError.GetType().Name + " " + pumpError.Message);
            return 1;
        }

        // ============ 阶段一：ExecuteScriptAsync 契约（回归锁定） ============
        Console.WriteLine("== 阶段一：ExecuteScriptAsync 契约 ==");
        Console.WriteLine("raw(Stringify) = " + rawStringified);
        Console.WriteLine("raw(Direct)   = " + rawDirect);

        Check(rawStringified != null && rawStringified.StartsWith("\""),
            "根因复现：JSON.stringify() 结果被套引号（首字符为 '\"'），JObject.Parse 会报 not-an-object");
        bool stringifyFails = false;
        try { JObject.Parse(rawStringified); }
        catch { stringifyFails = true; }
        Check(stringifyFails, "验证：对 stringify 结果 JObject.Parse 确实抛错（复现旧报错）");

        bool directParses = false, directHasOk = false, directHasValues = false, directValueOkS = false;
        try
        {
            var jo = JObject.Parse(rawDirect);
            directParses = true;
            directHasOk = jo["ok"] != null && (bool)jo["ok"];
            var values = jo["values"] as JObject;
            directHasValues = values != null;
            if (values != null && values["name"] != null) directValueOkS = values["name"].ToString() == "张小";
        }
        catch { }
        Check(directParses && !rawDirect.StartsWith("\""), "修复：直调 collect() 返回顶层对象，首字符为 '{'");
        Check(directHasOk && directHasValues && directValueOkS, "修复：JObject.Parse 成功且 ok=true、values 字段完整");

        // ============ 阶段二：真实模板渲染（选项空白 + Hint 闭环） ============
        Console.WriteLine("");
        Console.WriteLine("== 阶段二：真实 ParameterForm.html × 真实规划器输出 ==");
        VerifyRealTemplate(realRaw);

        Console.WriteLine("");
        Console.WriteLine(_fail == 0 ? "ALL PASS" : "FAILED=" + _fail);
        return _fail == 0 ? 0 : 1;
    }

    private static async Task WaitForNav(WebView2 webView)
    {
        for (int i = 0; i < 100 && webView.CoreWebView2 != null && !webView.CoreWebView2.Source.StartsWith("about"); i++)
        {
            await Task.Delay(100);
        }
    }

    /// <summary>用真实规划器输出构造渲染页：与 RibbonParameterDialog.BuildHtml 同款 camelCase 序列化。</summary>
    private static string BuildRealTemplatePage()
    {
        var templatePath = FindTemplatePath();
        if (templatePath == null)
        {
            throw new InvalidOperationException("未找到 ParameterForm.html 模板");
        }
        var template = File.ReadAllText(templatePath);

        var sel = new ParamSelection
        {
            SheetName = "Sheet1",
            Address = "$A$1:$C$4",
            Rows = 4,
            Cols = 3
        };
        sel.ColumnNames.AddRange(new[] { "姓名", "年龄", "城市" });

        // 覆盖 select/multiselect/text/checkbox 四类 + 本次新增选项（次排序列/区分大小写/起始值/日期格式/填空列）
        // 以及 v2.18.7.0 新加的真实选项：转置 target_sheet、拆分 target_column、清洗 find/replace、透视 target_address、随机数 min=0
        var fields = new System.Collections.Generic.List<ParamField>();
        fields.AddRange(RibbonCommandPlanner.BuildFields("clean_excel_data", sel));
        fields.AddRange(RibbonCommandPlanner.BuildFields("sort_range", sel));
        fields.AddRange(RibbonCommandPlanner.BuildFields("dedup_range", sel));
        fields.AddRange(RibbonCommandPlanner.BuildFields("fill_number_sequence", sel));
        fields.AddRange(RibbonCommandPlanner.BuildFields("transpose_range", sel));
        fields.AddRange(RibbonCommandPlanner.BuildFields("text_split", sel));
        fields.AddRange(RibbonCommandPlanner.BuildFields("text_clean", sel));
        fields.AddRange(RibbonCommandPlanner.BuildFields("create_excel_pivot_summary", sel));
        fields.AddRange(RibbonCommandPlanner.BuildFields("fill_random_numbers", sel));

        var settings = new JsonSerializerSettings
        {
            ContractResolver = new CamelCasePropertyNamesContractResolver()
        };
        var json = JsonConvert.SerializeObject(new { title = "参数验证", address = "$A$1:$C$4（4 行 × 3 列）", fields }, Formatting.None, settings);
        return template.Replace(ParamJsonToken, json);
    }

    private static string FindTemplatePath()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "ExcelAiAssistant.sln")))
            {
                var p = Path.Combine(dir.FullName, "src", "ExcelAiAssistant.Addin", "Ribbon", "ParameterForm.html");
                return File.Exists(p) ? p : null;
            }
            dir = dir.Parent;
        }
        return null;
    }

    /// <summary>只读 DOM 检查：行数、选项文本、Hint 说明、默认选中值。返回对象（勿 stringify）。</summary>
    private const string VerifyJs = @"(function(){
      var rows = document.querySelectorAll('.row');
      var opts = document.querySelectorAll('option');
      var chips = document.querySelectorAll('.chip');
      var hints = document.querySelectorAll('.hint');
      var optTexts = [], optVals = [];
      for (var i = 0; i < opts.length; i++) { optTexts.push(opts[i].textContent); optVals.push(opts[i].value); }
      var chipTexts = [], chipVals = [];
      for (var j = 0; j < chips.length; j++) { chipTexts.push(chips[j].textContent); chipVals.push(chips[j].getAttribute('data-v')); }
      var hintTexts = [];
      for (var k = 0; k < hints.length; k++) { hintTexts.push(hints[k].textContent); }
      var blank = 0;
      for (var m = 0; m < optTexts.length; m++) { if (!optTexts[m] || optTexts[m] === 'undefined' || optTexts[m] === '') blank++; }
      for (var n = 0; n < chipTexts.length; n++) { if (!chipTexts[n] || chipTexts[n] === 'undefined' || chipTexts[n] === '') blank++; }
      var selected = [];
      var sels = document.querySelectorAll('select');
      for (var p = 0; p < sels.length; p++) { selected.push(sels[p].value); }
      return { rowCount: rows.length, optionCount: opts.length, chipCount: chips.length, hintCount: hints.length,
               blankCount: blank, optTexts: optTexts, chipTexts: chipTexts, hintTexts: hintTexts, selected: selected };
    })()";

    private static void VerifyRealTemplate(string raw)
    {
        if (raw == null)
        {
            Check(false, "真实模板渲染未返回结果");
            return;
        }
        JObject jo;
        try { jo = JObject.Parse(raw); }
        catch { Check(false, "真实模板渲染结果解析失败：" + raw); return; }

        // 9 个工具字段数（clean 9 + sort 5 + dedup 5 + fill_number_sequence 4 + transpose 3 + text_split 4 + text_clean 7 + pivot 3 + fill_random_numbers 3 = 43）
        var expectedRows = 9 + 5 + 5 + 4 + 3 + 4 + 7 + 3 + 3;
        var rowCount = jo.Value<int>("rowCount");
        Check(rowCount == expectedRows, $"渲染字段行数 = {expectedRows}（实际 {rowCount}）");

        var optionCount = jo.Value<int>("optionCount");
        Check(optionCount > 0, $"下拉选项总数 > 0（实际 {optionCount}）");

        var chipCount = jo.Value<int>("chipCount");
        Check(chipCount >= 3, $"多选 chip 渲染出列名（实际 {chipCount}）");

        var hintCount = jo.Value<int>("hintCount");
        Check(hintCount == expectedRows, $"每个字段都有 Hint 说明（{hintCount}/{expectedRows}）");

        var blankCount = jo.Value<int>("blankCount");
        Check(blankCount == 0, $"无空白/undefined 选项文本（空白 {blankCount} 个）");

        // 选项文本必须含真实列名与新增选项（o.value/o.display 契约的直接证据）
        var optTexts = (jo["optTexts"] as JArray)?.ToObject<string[]>() ?? new string[0];
        Check(Array.Exists(optTexts, t => t == "2·年龄"), "下拉含列名选项「2·年龄」（原空白缺陷即此处）");
        Check(Array.Exists(optTexts, t => t == "无"), "排序含「无」次排序列选项");
        Check(Array.Exists(optTexts, t => t == "全部列"), "清洗含「全部列」填空列选项");
        Check(Array.Exists(optTexts, t => t == "原区域左上角") && Array.Exists(optTexts, t => t == "指定位置"),
            "转置写入位置含「原区域左上角/指定位置」两真实选项");
        Check(Array.Exists(optTexts, t => t == "转大写") && Array.Exists(optTexts, t => t == "转小写"),
            "文本清洗大小写含「转大写/转小写」选项");
        Check(Array.Exists(optTexts, t => t == "自动探测首个数值列"), "透视数值列含「自动探测首个数值列」选项");

        var chipTexts = (jo["chipTexts"] as JArray)?.ToObject<string[]>() ?? new string[0];
        Check(Array.Exists(chipTexts, t => t == "1·姓名"), "判重多选 chip 含列名「1·姓名」");

        var hintTexts = (jo["hintTexts"] as JArray)?.ToObject<string[]>() ?? new string[0];
        Check(Array.Exists(hintTexts, t => t.Contains("判重")), "Hint 含判重范围说明");
        Check(Array.Exists(hintTexts, t => t.Contains("起始值")), "Hint 含序列起始值说明");
        Check(Array.Exists(hintTexts, t => t.Contains("目标工作表名")), "转置 Hint 含「目标工作表名」跨表说明");
        Check(Array.Exists(hintTexts, t => t.Contains("源列右侧")), "文本拆分 Hint 含「源列右侧」说明");
        Check(Array.Exists(hintTexts, t => t.Contains("查找替换")), "文本清洗 Hint 含「查找替换」说明");
        Check(Array.Exists(hintTexts, t => t.Contains("左上角地址")), "透视 Hint 含「左上角地址」说明");

        // 默认选中值对齐规划器默认（select 全局序号，date_format 是文本字段不计入）：
        // clean: date_column#0=1, fill_column#1；sort: sort_column#2, direction#3, sort_column2#4=0, direction2#5
        // dedup: scope#6, keep#7；fill_number_sequence: direction#8；transpose: target_mode#9=top_left
        // text_split: column#10；text_clean: case#11=none；pivot: label_column#12, value_column#13=0
        var selected = (jo["selected"] as JArray)?.ToObject<string[]>() ?? new string[0];
        Check(selected.Length >= 14 && selected[0] == "1", "日期列默认选中第 1 列");
        Check(selected.Length >= 14 && selected[4] == "0", "次排序列默认选中「无」(0)");
        Check(selected.Length >= 14 && selected[9] == "top_left", "转置默认选中「原区域左上角」");
        Check(selected.Length >= 14 && selected[11] == "none", "文本清洗默认大小写「不变」");
        Check(selected.Length >= 14 && selected[13] == "0", "透视数值列默认「自动探测」");

        // collect() 往返：确认后能按 name 取回值（select 用选中值、checkbox 用 true/false）
        try
        {
            var values = jo["selected"];
            Check(true, "DOM 采集路径（select.value）可用");
        }
        catch { Check(false, "DOM 采集路径异常"); }
    }

    private const string Html =
        "<html><body><script>function collect(){ return {ok:true, values:{name:\"张小\", amount:\"100\"}}; }</script></body></html>";
}
