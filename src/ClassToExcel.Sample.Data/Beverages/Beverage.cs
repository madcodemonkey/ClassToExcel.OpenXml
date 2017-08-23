namespace ClassToExcel.Sample.Data
{
    public class Beverage
    {
        [ClassToExcelRow(ColumnLetter = "C", RowNumber = 2)]
        public int Beer { get; set; }
        [ClassToExcelRow(ColumnLetter = "C", RowNumber = 3)]
        public int Wine { get; set; }
        [ClassToExcelRow(ColumnLetter = "C", RowNumber = 4)]
        public int Pepsi { get; set; }
        [ClassToExcelRow(ColumnLetter = "C", RowNumber = 5)]
        public int Coke { get; set; }
        [ClassToExcelRow(ColumnLetter = "C", RowNumber = 6)]
        public int DrPepper { get; set; }
    }
}