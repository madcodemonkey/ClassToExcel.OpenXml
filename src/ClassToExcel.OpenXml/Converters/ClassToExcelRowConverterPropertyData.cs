using System.Reflection;

namespace ClassToExcel
{
    internal class ClassToExcelRowConverterPropertyData {
        public PropertyInfo Property { get; set; }
        /// <summary>The column letter that should be used.</summary>
        public string ColumnLetter { get; set; }

        /// <summary>The row number that should be used.</summary>
        public int RowNumber { get; set; }
    }
}