using System.Reflection;

namespace ClassToExcel
{
    public class ClassToExcelColumn
    {
        public string ColumnName { get; set; }
        public string ExcelColumnLetter { get; set; }
        public bool IsBoolean { get; set; }
        public bool IsDate { get; set; }
        public bool IsNumber { get; set; }
        public bool IsOptional { get; set; }
        public int Order { get; set; }
        public PropertyInfo Property { get; set; }
        public string StyleFormat { get; set; }
        public uint StyleIndex { get; set; }
    }
}