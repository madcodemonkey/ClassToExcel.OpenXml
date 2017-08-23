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
        }

        /// <summary>The column letter that should be used.</summary>
        public string ColumnLetter { get; set; }

        /// <summary>The row number that should be used.</summary>
        public int RowNumber { get; set; }
    }
}