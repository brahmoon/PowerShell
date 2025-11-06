using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelDiff
{
    internal static class ExcelDiffEngine
    {
        public static IReadOnlyList<DiffResult> Compare(string leftWorkbookPath, string rightWorkbookPath)
        {
            if (leftWorkbookPath is null)
            {
                throw new ArgumentNullException(nameof(leftWorkbookPath));
            }

            if (rightWorkbookPath is null)
            {
                throw new ArgumentNullException(nameof(rightWorkbookPath));
            }

            var differences = new List<DiffResult>();
            Excel.Application? excel = null;
            Excel.Workbook? leftWorkbook = null;
            Excel.Workbook? rightWorkbook = null;

            try
            {
                excel = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false,
                    ScreenUpdating = false
                };

                leftWorkbook = excel.Workbooks.Open(leftWorkbookPath, ReadOnly: true);
                rightWorkbook = excel.Workbooks.Open(rightWorkbookPath, ReadOnly: true);

                var sheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                sheetNames.UnionWith(GetSheetNames(leftWorkbook));
                sheetNames.UnionWith(GetSheetNames(rightWorkbook));

                foreach (var sheetName in sheetNames.OrderBy(n => n, StringComparer.OrdinalIgnoreCase))
                {
                    Excel.Worksheet? leftSheet = null;
                    Excel.Worksheet? rightSheet = null;

                    try
                    {
                        leftSheet = TryGetWorksheet(leftWorkbook, sheetName);
                        rightSheet = TryGetWorksheet(rightWorkbook, sheetName);

                        CompareSheet(leftSheet, rightSheet, sheetName, differences);
                    }
                    finally
                    {
                        ReleaseCom(rightSheet);
                        ReleaseCom(leftSheet);
                    }
                }
            }
            finally
            {
                if (rightWorkbook != null)
                {
                    rightWorkbook.Close(false);
                    ReleaseCom(rightWorkbook);
                }

                if (leftWorkbook != null)
                {
                    leftWorkbook.Close(false);
                    ReleaseCom(leftWorkbook);
                }

                if (excel != null)
                {
                    excel.Quit();
                    ReleaseCom(excel);
                }
            }

            return differences;
        }

        private static IEnumerable<string> GetSheetNames(Excel.Workbook? workbook)
        {
            if (workbook == null)
            {
                yield break;
            }

            foreach (var sheetObj in workbook.Worksheets)
            {
                if (sheetObj is Excel.Worksheet worksheet)
                {
                    yield return worksheet.Name;
                    ReleaseCom(worksheet);
                }
            }
        }

        private static Excel.Worksheet? TryGetWorksheet(Excel.Workbook? workbook, string sheetName)
        {
            if (workbook == null)
            {
                return null;
            }

            try
            {
                return workbook.Worksheets[sheetName] as Excel.Worksheet;
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static void CompareSheet(Excel.Worksheet? leftSheet, Excel.Worksheet? rightSheet, string sheetName, ICollection<DiffResult> output)
        {
            if (leftSheet == null && rightSheet == null)
            {
                return;
            }

            var leftValues = ReadSheetValues(leftSheet, out var leftMaxRow, out var leftMaxColumn);
            var rightValues = ReadSheetValues(rightSheet, out var rightMaxRow, out var rightMaxColumn);

            if (leftSheet == null && rightSheet != null && rightValues.Count == 0)
            {
                output.Add(new DiffResult(sheetName, 1, 1, "(シートなし)", "(空のシート)"));
                return;
            }

            if (rightSheet == null && leftSheet != null && leftValues.Count == 0)
            {
                output.Add(new DiffResult(sheetName, 1, 1, "(空のシート)", "(シートなし)"));
                return;
            }

            var maxRow = Math.Max(leftMaxRow, rightMaxRow);
            var maxColumn = Math.Max(leftMaxColumn, rightMaxColumn);

            for (var row = 1; row <= maxRow; row++)
            {
                for (var column = 1; column <= maxColumn; column++)
                {
                    var left = leftValues.TryGetValue((row, column), out var leftValue) ? leftValue : string.Empty;
                    var right = rightValues.TryGetValue((row, column), out var rightValue) ? rightValue : string.Empty;

                    if (!string.Equals(left, right, StringComparison.Ordinal))
                    {
                        output.Add(new DiffResult(sheetName, row, column, left, right));
                    }
                }
            }
        }

        private static Dictionary<(int Row, int Column), string> ReadSheetValues(Excel.Worksheet? sheet, out int maxRow, out int maxColumn)
        {
            maxRow = 0;
            maxColumn = 0;
            var values = new Dictionary<(int, int), string>();

            if (sheet == null)
            {
                return values;
            }

            Excel.Range? usedRange = null;

            try
            {
                usedRange = sheet.UsedRange;
                if (usedRange == null || usedRange.CountLarge == 0)
                {
                    return values;
                }

                var startRow = usedRange.Row;
                var startColumn = usedRange.Column;
                var rows = usedRange.Rows.Count;
                var columns = usedRange.Columns.Count;
                maxRow = Math.Max(maxRow, startRow + rows - 1);
                maxColumn = Math.Max(maxColumn, startColumn + columns - 1);

                var raw = usedRange.Value2;
                if (raw is object[,] array)
                {
                    var lowerRow = array.GetLowerBound(0);
                    var upperRow = array.GetUpperBound(0);
                    var lowerColumn = array.GetLowerBound(1);
                    var upperColumn = array.GetUpperBound(1);

                    for (var r = lowerRow; r <= upperRow; r++)
                    {
                        for (var c = lowerColumn; c <= upperColumn; c++)
                        {
                            var text = ConvertValue(array[r, c]);
                            if (text is null)
                            {
                                continue;
                            }

                            var absoluteRow = startRow + (r - lowerRow);
                            var absoluteColumn = startColumn + (c - lowerColumn);
                            values[(absoluteRow, absoluteColumn)] = text;
                        }
                    }
                }
                else if (raw != null)
                {
                    var text = ConvertValue(raw);
                    if (text != null)
                    {
                        values[(startRow, startColumn)] = text;
                    }
                }
            }
            finally
            {
                ReleaseCom(usedRange);
            }

            return values;
        }

        private static string? ConvertValue(object? value)
        {
            switch (value)
            {
                case null:
                    return null;
                case string s:
                    return s;
                case double number when double.IsNaN(number):
                    return null;
                case double number:
                    return number.ToString(CultureInfo.InvariantCulture);
                case bool boolean:
                    return boolean ? "TRUE" : "FALSE";
                case DateTime dateTime:
                    return dateTime.ToString(CultureInfo.InvariantCulture);
                default:
                    return Convert.ToString(value, CultureInfo.InvariantCulture);
            }
        }

        private static void ReleaseCom(object? comObject)
        {
            if (comObject is null)
            {
                return;
            }

            try
            {
                Marshal.FinalReleaseComObject(comObject);
            }
            catch
            {
                // Intentionally ignore failure to release COM objects. Excel will clean up when the process exits.
            }
        }
    }
}
