using Microsoft.Office.Interop.Excel;

namespace Utility.syonoki.MSOffice {
    public static class XlRange {
        public static Range rightDownRange(Worksheet wks, Range startRange)
            => wks.Range[startRange, startRange.End[XlDirection.xlToRight].End[XlDirection.xlDown]];

        public static Range downRightRange(Worksheet wks, Range startRange)
            => wks.Range[startRange, startRange.End[XlDirection.xlDown].End[XlDirection.xlToRight]];

        public static Range rightAndnRowsRange(Worksheet wks, Range startRange, int n)
            => wks.Range[startRange, startRange.End[XlDirection.xlToRight].Cells[n, 1]];

        public static Range downAndnColumnsRange(Worksheet wks, Range startRange, int n)
            => wks.Range[startRange, startRange.End[XlDirection.xlDown].Cells[1, n]];

        public static TResult valueOfRange<TResult>(Range targetRng, int i, int j)
            => (TResult)(targetRng.Value as object[,])[i, j];

        public static int rowCount(Range targetRng)
            => (targetRng.Columns[1].Value as object[,]).Length;

        public static int columnCount(Range targetRng)
            => (targetRng.Rows[1].Value as object[,]).Length;
    }
}