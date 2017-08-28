namespace ClassToExcel.Sample.Data
{
    public class Beverage
    {
        private int _drPepper;
        private decimal _drPepperPriceRaw;

        [ClassToExcelRow(ColumnLetter = "C", RowNumber = 2)]
        public int Beer { get; set; }

        [ClassToExcelRow(ColumnLetter = "D", RowNumber = 2, DecimalPlaces = 2)]
        public decimal BeerPrice { get; set; }

        [ClassToExcelRow(ColumnLetter = "C", RowNumber = 3)]
        public int Wine { get; set; }

        [ClassToExcelRow(ColumnLetter = "D", RowNumber = 3, DecimalPlaces = 2)]
        public decimal WinePrice { get; set; }

        [ClassToExcelRow(ColumnLetter = "C", RowNumber = 4)]
        public int Pepsi { get; set; }

        [ClassToExcelRow(ColumnLetter = "D", RowNumber = 4, DecimalPlaces = 2)]
        public decimal PepsiPrice { get; set; }


        [ClassToExcelRow(ColumnLetter = "C", RowNumber = 5)]
        public int Coke { get; set; }

        [ClassToExcelRow(ColumnLetter = "D", RowNumber = 5, DecimalPlaces = 2)]
        public decimal CokePrice { get; set; }


        [ClassToExcelRow(ColumnLetter = "C", RowNumber = 6)]
        private int DrPepperRaw
        {
            get { return _drPepper; }
            set { _drPepper = value; }
        }

        public int DrPepper
        {
            get { return _drPepper; }
            set { _drPepper = value; }
        }

        [ClassToExcelRow(ColumnLetter = "D", RowNumber = 6, DecimalPlaces = 2)]
        private decimal DrPepperPriceRaw
        {
            get { return _drPepperPriceRaw; }
            set { _drPepperPriceRaw = value; }
        }

        public decimal DrPepperPrice
        {
            get { return _drPepperPriceRaw; }
            set { _drPepperPriceRaw = value; }
        }


        [ClassToExcelRow(ColumnLetter = "C", RowNumber = 7, DecimalPlaces = 1)]
        public double AvgNumberOfLiters { get; set; }
    }
}