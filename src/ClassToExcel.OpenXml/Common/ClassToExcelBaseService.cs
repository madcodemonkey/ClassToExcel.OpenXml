using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClassToExcel
{
    public class ClassToExcelBaseService
    {
        public ClassToExcelBaseService(Action<ClassToExcelMessage> loggingCallback)
        {
            _loggingCallback = loggingCallback;
            WorksheetColumns = new Dictionary<string, List<ClassToExcelColumn>>();
        }
        public Dictionary<string, List<ClassToExcelColumn>> WorksheetColumns { get; private set; }

        private readonly Action<ClassToExcelMessage> _loggingCallback;

        /// <summary>Used when reading a file to get the string value from a cell.</summary>
        protected string GetCellText(Cell c, SharedStringTable sst)
        {
            if (sst != null && c.DataType != null && c.DataType == CellValues.SharedString)
            {
                int ssid = int.Parse(c.CellValue.Text);
                return sst.ChildElements[ssid].InnerText;
            }

            if (c == null || c.CellValue == null)
                return null;

            return c.CellValue.Text;
        }

        /// <summary>Used when reading a file to get the column/cell letter.</summary>
        protected string GetColumnLetter(Cell oneCell)
        {
            string cellReference = oneCell != null && oneCell.CellReference != null
                ? oneCell.CellReference.ToString()
                : string.Empty;

            // XFD largest column ID for Excel 2013
            if (string.IsNullOrWhiteSpace(cellReference))
                return string.Empty;
            if (char.IsNumber(cellReference, 0))
                return string.Empty;
            if (char.IsNumber(cellReference, 1))
                return cellReference.Substring(0, 1);
            if (char.IsNumber(cellReference, 2))
                return cellReference.Substring(0, 2);

            return cellReference.Substring(0, 3);
        }

        /// <summary>Get an Excel column number.  So, if you pass in 1, you will get A.  If you pass in 2, you will get B.
        /// This will continue to reach AA, AB, etc. just as Excel numbers its columns).</summary>
        protected string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        /// <summary>Using reflection, create the columns and pull the attribute information off the class so
        /// that we know how to format the columns later.</summary>
        protected List<ClassToExcelColumn> CreateColumns(string worksheetName, Type typeOfClass)
        {
            List<ClassToExcelColumn> columns = new List<ClassToExcelColumn>();

            foreach (PropertyInfo property in typeOfClass.GetProperties())
            {
                if ((property.PropertyType != typeof(string) && property.PropertyType.IsClass) || property.PropertyType.IsArray)
                    continue;
                if (property.CanWrite == false)
                {
                    LogMessage(ClassToExcelMessageType.Info, string.Format("Ignoring property because it is read-only: {0}", property.Name));
                    continue;
                }
                var newData = new ClassToExcelColumn { ColumnName = property.Name, Order = 99999, Property = property };

                // If the property has a dislay name attribute, use it as the column name.
                object firstAttribute = property.GetCustomAttributes(typeof(ClassToExcelAttribute), false).FirstOrDefault();
                if (firstAttribute != null)
                {
                    ClassToExcelAttribute displayAttribute = firstAttribute as ClassToExcelAttribute;
                    if (displayAttribute != null)
                    {
                        if (displayAttribute.Ignore)
                        {
                            continue; // do not add this column
                        }

                        // Does it have a column name?
                        if (string.IsNullOrWhiteSpace(displayAttribute.ColumnName) == false)
                            newData.ColumnName = displayAttribute.ColumnName;
                    
                        // Is it optional?
                        newData.IsOptional = displayAttribute.IsOptional;

                        // Does it have an order?
                        if (displayAttribute.Order > 0)
                            newData.Order = displayAttribute.Order;

                        // Is it a number, date or boolean?
                        if (property.PropertyType != typeof(string))
                        {
                            if (property.PropertyType == typeof(int) || property.PropertyType == typeof(int?))
                                newData.IsInteger = true;
                            else if (property.PropertyType == typeof(double) || property.PropertyType == typeof(double?))
                                newData.IsDouble = true;
                            else if (property.PropertyType == typeof(decimal) || property.PropertyType == typeof(decimal?))
                                newData.IsDecimal = true;
                            else if (property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(DateTime?))
                                newData.IsDate = true;
                            else if (property.PropertyType == typeof(bool) || property.PropertyType == typeof(bool?))
                                newData.IsBoolean = true;
                        }
                        
                        // Any special styling?
                        if (string.IsNullOrWhiteSpace(displayAttribute.StyleFormat) == false && (newData.IsNumber() || newData.IsDate))
                        {
                            newData.StyleFormat = displayAttribute.StyleFormat.Trim();
                        }
                    }
                }

                columns.Add(newData);
            }

            columns = columns.OrderBy(o => o.Order).ThenBy(o => o.ColumnName).ToList();

            for (int i = 1; i <= columns.Count; i++)
            {
                columns[i - 1].ExcelColumnLetter = GetExcelColumnName(i);
            }

            if (WorksheetColumns.ContainsKey(worksheetName))
            {
                WorksheetColumns[worksheetName] = columns.OrderBy(o => o.Order).ToList();
            }
            else
            {
                WorksheetColumns.Add(worksheetName, columns);
            }

            return WorksheetColumns[worksheetName];
        }

        protected void LogMessage(ClassToExcelMessage message)
        {
            if (_loggingCallback != null)
            {
                _loggingCallback(message);
            }
        }

        protected void LogMessage(ClassToExcelMessageType messageType, string message)
        {
            if (_loggingCallback != null)
            {
                _loggingCallback(new ClassToExcelMessage(messageType, message));
            }
        }
    }
}