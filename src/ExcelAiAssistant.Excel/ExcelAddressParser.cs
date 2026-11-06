using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelAiAssistant.Excel
{
    public static class ExcelAddressParser
    {
        private static readonly Regex Pattern = new Regex(
            @"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$",
            RegexOptions.Compiled);

        public static bool TryParseCell(string address, out int row, out int column)
        {
            row = 0;
            column = 0;
            if (string.IsNullOrWhiteSpace(address))
            {
                return false;
            }

            var match = Pattern.Match(address.Trim());
            if (!match.Success)
            {
                return false;
            }

            row = int.Parse(match.Groups[2].Value, NumberStyles.Integer, CultureInfo.InvariantCulture);
            column = ColumnLettersToNumber(match.Groups[1].Value.ToUpperInvariant());
            return row <= 1048576 && column <= 16384;
        }

        private static int ColumnLettersToNumber(string letters)
        {
            var value = 0;
            foreach (var character in letters)
            {
                value = value * 26 + (character - 'A' + 1);
            }

            return value;
        }
    }
}
