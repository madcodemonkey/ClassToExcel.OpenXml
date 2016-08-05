using System.Collections.Generic;

namespace ClassToExcel
{
    public class ClassToExcelRawRow
    {
        public ClassToExcelRawRow()
        {
            Columns = new List<ClassToExcelRawColumn>();
        }

        public int RowIndex { get; set; }
        public List<ClassToExcelRawColumn> Columns { get; set; }
    }
}