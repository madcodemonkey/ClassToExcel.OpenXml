namespace ClassToExcel
{
    public class ClassToExcelMessage
    {
        public ClassToExcelMessage(ClassToExcelMessageType messageType, string message) : this(messageType, 0, null, message){}
        public ClassToExcelMessage(ClassToExcelMessageType messageType, int rowNumber, string columnName, string columnLetter, string message)
        {
            RowNumber = rowNumber;
            ColumnName = columnName;
            ColumnLetter = columnLetter;
            Message = message;
            MessageType = messageType;
        }
        public ClassToExcelMessage(ClassToExcelMessageType messageType, int rowNumber, ClassToExcelColumn column, string message)
        {
            RowNumber = rowNumber;
            Message = message;
            MessageType = messageType;

            if (column != null)
            {
                ColumnName = column.ColumnName;
                ColumnLetter = string.IsNullOrWhiteSpace(column.ExcelColumnLetter)
                    ? UnknownColumnLetter
                    : column.ExcelColumnLetter;
            }
            else
            {
                ColumnLetter =  UnknownColumnLetter;
            }
        }
        public int RowNumber { get; set; }
        public string ColumnName { get; set; }
        public string ColumnLetter { get; set; }
        public string Message { get; set; }

        public ClassToExcelMessageType MessageType { get; set; }

        public const string UnknownColumnLetter = "?";

        public override string ToString()
        {
            if (RowNumber > 0 || string.IsNullOrWhiteSpace(ColumnName) == false || ColumnLetter != UnknownColumnLetter)
               return string.Format("{0} -> Row {1} Column[{2}]: '{3}' : {4}", MessageType, RowNumber, ColumnLetter, ColumnName,  Message);
            return string.Format("{0} -> {1}", MessageType, Message); 
        }
    }
}
