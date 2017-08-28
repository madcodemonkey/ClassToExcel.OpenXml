using System;

namespace ClassToExcel
{
    /// <summary>This is used by the ClassToExcelRowConverter</summary>
    public class ClassToExcelRowAttribute : Attribute
    {
        public ClassToExcelRowAttribute()
        {
            ColumnLetter = "A";
            RowNumber = 1;
            DecimalPlaces = -1; // Less than 0 means don't round so turn it off by default.
        }

        /// <summary>The column letter that should be used.</summary>
        public string ColumnLetter { get; set; }

        /// <summary>The row number that should be used.</summary>
        public int RowNumber { get; set; }

        /// <summary>The number of digits of percision you want.  Setting it to a number less than one will avoid performing any action on a number.</summary>
        public int DecimalPlaces { get; set; }

    }
}