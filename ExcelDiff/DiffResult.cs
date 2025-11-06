using System;

namespace ExcelDiff
{
    internal sealed class DiffResult
    {
        public DiffResult(string sheetName, int row, int column, string leftValue, string rightValue)
        {
            SheetName = sheetName ?? string.Empty;
            Row = row;
            Column = column;
            LeftValue = leftValue ?? string.Empty;
            RightValue = rightValue ?? string.Empty;
        }

        public string SheetName { get; }

        public int Row { get; }

        public int Column { get; }

        public string LeftValue { get; }

        public string RightValue { get; }

        public string Address => FormattableString.Invariant($"{GetColumnLabel(Column)}{Row}");

        private static string GetColumnLabel(int columnIndex)
        {
            if (columnIndex <= 0)
            {
                return "?";
            }

            var current = columnIndex;
            var label = string.Empty;
            while (current > 0)
            {
                current--;
                label = (char)('A' + (current % 26)) + label;
                current /= 26;
            }

            return label;
        }
    }
}
