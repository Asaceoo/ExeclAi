using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAiAssistant.Excel
{
    /// <summary>P0-2/P1-3：L2 元素操作与外部数据源 schema。</summary>
    public static class ElementOpSchemas
    {
        private static Dictionary<string, object> Obj(string[] required, params Dictionary<string, object>[] props)
        {
            var map = new Dictionary<string, object>();
            foreach (var p in props) map[Convert.ToString(p["name"])] = p;
            var s = new Dictionary<string, object> { ["type"] = "object", ["properties"] = map };
            if (required != null && required.Length > 0) s["required"] = required;
            return s;
        }
        private static Dictionary<string, object> Prop(string name, object type, string desc) =>
            new Dictionary<string, object> { ["name"] = name, ["type"] = type, ["description"] = desc };

        private static Dictionary<string, object> PropStr(string name, string type, string desc) =>
            new Dictionary<string, object> { ["name"] = name, ["type"] = type, ["description"] = desc };

        public static Dictionary<string, object> QueryElements() =>
            Obj(new[] { "element_type" }, Prop("element_type", "string", "元素类型：sheets / shapes / charts / usedrange_rows。"), Prop("sheet", "string", "目标工作表（shapes/charts/usedrange_rows 需要），默认活动工作表。"));

        public static Dictionary<string, object> GetElement() =>
            Obj(new[] { "address" }, Prop("address", "string", "区域地址，如 A1:C10。"), Prop("sheet", "string", "目标工作表，默认活动工作表。"));

        public static Dictionary<string, object> SetElement() =>
            Obj(new[] { "address", "value" }, Prop("address", "string", "区域地址。"), Prop("sheet", "string", "目标工作表，默认活动工作表。"), Prop("value", new[] { "string", "number", "boolean" }.ToArray(), "要设置的值。"));

        public static Dictionary<string, object> AddElement() =>
            Obj(new[] { "kind" }, Prop("kind", "string", "元素类型：textbox。"), Prop("sheet", "string", "目标工作表，默认活动工作表。"), Prop("top", "number", "左上角纵坐标（磅）。"), Prop("left", "number", "左上角横坐标（磅）。"), Prop("width", "number", "宽度（磅）。"), Prop("height", "number", "高度（磅）。"), Prop("text", "string", "文本框内容。"));

        public static Dictionary<string, object> RemoveElement() =>
            Obj(new[] { "kind", "name" }, Prop("kind", "string", "元素类型：shape / chart。"), Prop("name", "string", "元素名称（先 query_elements 查询）。"), Prop("sheet", "string", "目标工作表，默认活动工作表。"));

        public static Dictionary<string, object> MoveElement() =>
            Obj(new[] { "kind", "name" }, Prop("kind", "string", "元素类型：shape / chart。"), Prop("name", "string", "元素名称。"), Prop("sheet", "string", "目标工作表，默认活动工作表。"), Prop("top", "number", "新纵坐标。"), Prop("left", "number", "新横坐标。"));

        public static Dictionary<string, object> ImportCsv() =>
            Obj(new[] { "file_path" }, Prop("file_path", "string", "CSV 文件完整路径。"), Prop("sheet_name", "string", "新工作表名（默认 CSV_文件名）。"), Prop("delimiter", "string", "分隔符，默认逗号。"));

        public static Dictionary<string, object> DraftStart() =>
            Obj(new[] { "sheet" }, Prop("sheet", "string", "源工作表名（AI 将在其隐藏副本上操作）。"));

        public static Dictionary<string, object> DraftDiff() =>
            Obj(new[] { "draft_sheet" }, Prop("draft_sheet", "string", "start_draft 返回的草稿名。"));

        public static Dictionary<string, object> DraftApply() =>
            Obj(new[] { "draft_sheet" }, Prop("draft_sheet", "string", "要采纳的草稿名（采纳后草稿删除，变更落盘源表）。"));

        public static Dictionary<string, object> DraftDiscard() =>
            Obj(new[] { "draft_sheet" }, Prop("draft_sheet", "string", "要放弃的草稿名。"));
    }
}
