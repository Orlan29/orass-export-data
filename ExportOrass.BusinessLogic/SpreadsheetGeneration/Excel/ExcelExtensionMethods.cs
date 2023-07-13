using System;

namespace InfiSoftware.Common.DataAccess.SpreadsheetGeneration.Excel
{
    public static class ExcelExtensionMethods
    {
        public static int ToColumnNb(this string columnName)
        {
            if (string.IsNullOrEmpty(columnName))
                return -1;

            columnName = columnName.ToUpperInvariant();
            var sum = 0;
            foreach (var t in columnName)
            {
                sum *= 26;
                sum += t - 'A' + 1;
            }
            return sum;
        }

        public static string ToColumnName(this int columnNb)
        {
            if (columnNb <= 0)
                return "#N/A";
            var dividend = columnNb;
            var columnName = string.Empty;
            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }

        public static int ToColumnInt(this object column)
        {
            var s = column as string;
            var columnInt = s?.ToColumnNb() ?? (int)column;
            return columnInt;
        }

        public static bool Between(this int num, int lower, int upper, bool inclusive = true)
        {
            return inclusive
                ? lower <= num && num <= upper
                : lower < num && num < upper;
        }
    }
}
