using System;

namespace ClassToExcel.Sample.Data
{
    public class BeverageDates
    {
        [ClassToExcelRow(ColumnLetter = "B", RowNumber = 9)]
        public DateTime BeerStart { get; set; }
        [ClassToExcelRow(ColumnLetter = "C", RowNumber = 9)]
        public DateTime BeerEnd { get; set; }

        [ClassToExcelRow(ColumnLetter = "B", RowNumber = 10)]
        public DateTime WineStart { get; set; }
        [ClassToExcelRow(ColumnLetter = "C", RowNumber = 10)]
        public DateTime WineEnd { get; set; }
    }
}