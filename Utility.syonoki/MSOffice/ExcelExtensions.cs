using System;
using Microsoft.Office.Interop.Excel;

namespace Utility.syonoki.MSOffice {
    public static class ExcelExtensions {
        public enum FindingCellDirection {
            Up, Down, Left, Right
        }

        public static string findValue(this Worksheet wks, string key, FindingCellDirection direction, int interval = 2) {
            Range findCell = wks.Cells.Find(key);
            if (direction == FindingCellDirection.Right) {
                return Convert.ToString(findCell[1, interval].Value);
            }
            if (direction == FindingCellDirection.Down) {
                return Convert.ToString(findCell[interval, 1].Value);
            }

            return null;
        }
    }

    public static class XlRangeExtension
    {
        public static TResult valueOfRange<TResult>(this Range targetRng, int i, int j)
            => XlRange.valueOfRange<TResult>(targetRng, i, j);

        public static int rowCount(this Range targetRng)
            => XlRange.rowCount(targetRng);

        public static int columnCount(this Range targetRng)
            => XlRange.columnCount(targetRng);
    }
}