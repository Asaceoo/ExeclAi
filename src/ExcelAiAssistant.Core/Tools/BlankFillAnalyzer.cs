using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Core.Tools
{
    public enum BlankCellKind
    {
        EmptyString
    }

    /// <summary>一个待填充的空缺格。Row/Col 为 1 基（与 Excel 单元格一致）。</summary>
    public sealed class BlankCell
    {
        public int Row;
        public int Col;
        public string CellAddress;
        public BlankCellKind Kind;
    }

    /// <summary>
    /// 数据智能填充的纯逻辑：从含表头+数据的二维块中定位「空缺单元格」（仅空白格），
    /// 并按列归类、附带该列表头与非空样本，供模型做列内/列间语义推理。不接触 COM，可单测。
    /// </summary>
    public static class BlankFillAnalyzer
    {
        /// <summary>判定单元格值是否视为「空白」。仅 null / DBNull / 空字符串 / 全空格视为待填；占位文本（如"?"、"待填"）不算。</summary>
        public static bool IsLikelyBlank(object value)
        {
            if (value == null || value is DBNull) return true;
            return value is string s && (s.Length == 0 || s.Trim().Length == 0);
        }

        /// <summary>扫描整块，返回所有空白格（1 基行/列）。hasHeader=true 时表头行不参与空缺判断。</summary>
        public static List<BlankCell> FindBlanks(object[,] block, bool hasHeader)
        {
            var result = new List<BlankCell>();
            if (block == null) return result;
            int rows = block.GetLength(0);
            int cols = block.GetLength(1);
            int startRow = hasHeader && rows > 1 ? 1 : 0;
            if (hasHeader && rows == 1)
            {
                // 只有表头行，无数据行可填
                return result;
            }

            for (var r = startRow; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    var value = block[r, c];
                    if (!IsLikelyBlank(value)) continue;
                    result.Add(new BlankCell
                    {
                        Row = r + 1,
                        Col = c + 1,
                        CellAddress = CellAddress(r + 1, c + 1),
                        Kind = BlankCellKind.EmptyString
                    });
                }
            }

            return result;
        }

        /// <summary>按列号归并空缺格，并附带该列表头名（若存在）与非空样本。样本默认各列取前 sampleLimit 个非空格的实际值。</summary>
        public static Dictionary<int, ColumnBlankInfo> GroupByColumn(
            object[,] block, List<BlankCell> blanks, bool hasHeader, int sampleLimit = 5)
        {
            var groups = new Dictionary<int, ColumnBlankInfo>();
            if (blanks == null || blanks.Count == 0) return groups;
            if (block == null) return groups;

            int rows = block.GetLength(0);
            int cols = block.GetLength(1);
            int dataStart = hasHeader && rows > 1 ? 1 : 0;

            foreach (var blank in blanks)
            {
                if (blank.Col < 1 || blank.Col > cols) continue;
                if (!groups.TryGetValue(blank.Col, out var info))
                {
                    string headerName = hasHeader && rows > 1 ? CellText(block[0, blank.Col - 1]) : null;
                    info = new ColumnBlankInfo
                    {
                        Column = blank.Col,
                        ColumnName = headerName,
                        Blanks = new List<BlankCell>(),
                        Samples = CollectSamples(block, blank.Col, dataStart, sampleLimit)
                    };
                    groups[blank.Col] = info;
                }

                info.Blanks.Add(blank);
            }

            return groups;
        }

        /// <summary>返回某列前 sampleLimit 个非空样本的真实值（从数据起始行开始扫描）。</summary>
        private static List<object> CollectSamples(object[,] block, int column, int dataStart, int sampleLimit)
        {
            var samples = new List<object>();
            int rows = block.GetLength(0);
            for (var r = dataStart; r < rows && samples.Count < sampleLimit; r++)
            {
                var v = block[r, column - 1];
                if (!IsLikelyBlank(v))
                {
                    samples.Add(v);
                }
            }

            return samples;
        }

        /// <summary>按行归并空缺格的「整行邻近上下文」（D2）：对每个含空缺的行，给出同行各列值（空格列标 null+IsBlank）。</summary>
        public static List<RowContext> BuildRowContext(object[,] block, List<BlankCell> blanks, bool hasHeader)
        {
            var result = new List<RowContext>();
            if (blanks == null || blanks.Count == 0) return result;
            if (block == null) return result;

            int rows = block.GetLength(0);
            int cols = block.GetLength(1);

            var blankSet = new HashSet<int>();
            foreach (var b in blanks) blankSet.Add(b.Row);

            var ordered = new List<int>(blankSet);
            ordered.Sort();
            Dictionary<int, string> headerMap = null;
            if (hasHeader && rows > 1)
            {
                headerMap = new Dictionary<int, string>();
                for (var c = 0; c < cols; c++) headerMap[c + 1] = CellText(block[0, c]);
            }

            foreach (var row in ordered)
            {
                if (row < 1 || row > rows) continue;
                var cells = new List<RowCell>();
                for (var c = 0; c < cols; c++)
                {
                    var isBlank = IsLikelyBlank(block[row - 1, c]);
                    cells.Add(new RowCell
                    {
                        Column = c + 1,
                        ColumnName = headerMap != null && headerMap.TryGetValue(c + 1, out var h) ? h : null,
                        Value = isBlank ? null : block[row - 1, c],
                        IsBlank = isBlank
                    });
                }

                result.Add(new RowContext { Row = row, Cells = cells });
            }

            return result;
        }

        private static string CellText(object v)
        {
            return v == null || v is DBNull ? string.Empty : Convert.ToString(v);
        }

        /// <summary>1 基（row, col）→ Excel 地址，如 (1,1)→A1、(27,28)→AB27。</summary>
        public static string CellAddress(int row, int column)
        {
            return $"{ColumnName(column)}{row}";
        }

        private static string ColumnName(int column)
        {
            string name = string.Empty;
            var n = column;
            while (n > 0)
            {
                var remainder = (n - 1) % 26;
                name = (char)('A' + remainder) + name;
                n = (n - 1) / 26;
            }

            return name;
        }
    }

    /// <summary>某列的空缺信息：列号、列表头（可能为 null）、空缺格、该列非空样本。</summary>
    public sealed class ColumnBlankInfo
    {
        public int Column;
        public string ColumnName;
        public List<BlankCell> Blanks = new List<BlankCell>();
        public List<object> Samples = new List<object>();
    }

    /// <summary>某空格所在行的整行邻近上下文（D2）：同行各列值，空格列 Value=null + IsBlank。</summary>
    public sealed class RowContext
    {
        public int Row;
        public List<RowCell> Cells = new List<RowCell>();
    }

    public sealed class RowCell
    {
        public int Column;
        public string ColumnName;
        public object Value;
        public bool IsBlank;
    }
}