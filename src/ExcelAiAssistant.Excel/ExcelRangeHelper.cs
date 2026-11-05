using System;

namespace ExcelAiAssistant.Excel
{
    public static class ExcelRangeHelper
    {
        public static string Resize(string address, int rows, int columns)
        {
            // 工具分类审查批次真 bug（单测复现实锤）：0 行/0 列曾被 Math.Max(rows-1,0) 钳成 1 行地址
            // （A1:C10+0行 → A1:C1），FilterData 0 匹配/CleanData 全空行的 0 行矩阵写入该地址时，
            // Refit 清空语义会把源区域首行清空（数据破坏）。fail-closed 返 null，调用方提前拦截。
            if (rows < 1 || columns < 1)
            {
                return null;
            }

            if (!TryParseRange(address, out var sheet, out var firstRow, out var firstColumn, out _, out _))
            {
                return null;
            }

            return Format(sheet, firstRow, firstColumn, firstRow + rows - 1, firstColumn + columns - 1);
        }

        public static bool TryParseRange(
            string value,
            out string sheet,
            out int firstRow,
            out int firstColumn,
            out int lastRow,
            out int lastColumn)
        {
            sheet = null;
            firstRow = firstColumn = lastRow = lastColumn = 0;
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var rangePart = value;
            var separator = value.IndexOf('!');
            if (separator > 0)
            {
                sheet = value.Substring(0, separator).Trim('\'', '"');
                rangePart = value.Substring(separator + 1);
            }

            rangePart = rangePart.Replace("$", string.Empty);
            var parts = rangePart.Split(':');
            if (parts.Length == 0 || parts.Length > 2 || !TryParseCell(parts[0], out firstRow, out firstColumn))
            {
                return false;
            }

            if (parts.Length == 2)
            {
                return TryParseCell(parts[1], out lastRow, out lastColumn) &&
                       firstRow <= lastRow && firstColumn <= lastColumn;
            }

            lastRow = firstRow;
            lastColumn = firstColumn;
            return true;
        }

        private static bool TryParseCell(string cell, out int row, out int column)
        {
            row = column = 0;
            if (string.IsNullOrWhiteSpace(cell))
            {
                return false;
            }

            var letters = 0;
            var index = 0;
            while (index < cell.Length && char.IsLetter(cell[index]))
            {
                letters = letters * 26 + char.ToUpperInvariant(cell[index]) - 'A' + 1;
                index++;
            }

            if (letters == 0 || index >= cell.Length || !int.TryParse(cell.Substring(index), out row) || row < 1)
            {
                return false;
            }

            column = letters;
            return true;
        }

        private static string Format(string sheet, int firstRow, int firstColumn, int lastRow, int lastColumn)
        {
            var range = $"{ColumnName(firstColumn)}{firstRow}:{ColumnName(lastColumn)}{lastRow}";
            return string.IsNullOrEmpty(sheet) ? range : $"'{sheet}'!{range}";
        }

        public static string ColumnName(int column)
        {
            var name = string.Empty;
            while (column > 0)
            {
                var remainder = (column - 1) % 26;
                name = Convert.ToChar('A' + remainder) + name;
                column = (column - remainder - 1) / 26;
            }

            return name;
        }

        public static string CellAddress(int row, int column)
        {
            return $"{ColumnName(column)}{row}";
        }

        // 把相对地址转成带 $ 的绝对引用（不含工作表前缀）。
        public static string ToAbsolute(string rangePart)
        {
            var separator = rangePart.IndexOf('!');
            if (separator > 0)
            {
                rangePart = rangePart.Substring(separator + 1);
            }

            if (!TryParseRange(rangePart, out _, out var firstRow, out var firstColumn, out var lastRow, out var lastColumn))
            {
                return null;
            }

            return lastRow == firstRow && lastColumn == firstColumn
                ? $"${ColumnName(firstColumn)}${firstRow}"
                : $"${ColumnName(firstColumn)}${firstRow}:${ColumnName(lastColumn)}${lastRow}";
        }
    }
}
