using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClassToExcel
{
    /// <summary>Reads only primitives and maps them ont a given class type.</summary>
    /// <typeparam name="T"></typeparam>
    public class ClassToExcelReaderService<T> : ClassToExcelBaseService, IClassToExcelReaderService<T> where T : class, new()
    {
        public ClassToExcelReaderService() : base(null) { }
        public ClassToExcelReaderService(Action<ClassToExcelMessage> loggingCallback) : base(loggingCallback){}

        /// <summary>Read a worksheet from an Excel file.</summary>
        /// <param name="fileName">Name and path of the file</param>
        /// <param name="worksheetName">Tab / Worksheet name</param>
        /// <param name="hasHeaderRow">Is there a header row with the names of the columns on the first row?</param>
        /// <param name="startRow">Starting row. This is a ONE based number, but keep in mind that if hasHeaderRow is true, the header row IS row 1.
        /// If the value is null, it is not used.</param>
        /// <param name="endRow">Ending row.  This is a ONE based number, but keep in mind that if hasHeaderRow is true, the header row IS row 1.
        /// If the value is null, it is not used.</param>
        /// <returns>List of objects</returns>
        public List<T> ReadWorksheet(string fileName, string worksheetName, bool hasHeaderRow = true, int? startRow = null, int? endRow = null)
        {
            if (File.Exists(fileName) == false)
                throw new ArgumentException("The file does not exists: " + fileName);

            using (var fs = File.OpenRead(fileName))
            {
                return ReadWorksheet(fs, worksheetName, hasHeaderRow, startRow, endRow);
            }
        }

        /// <summary>Read a worksheet from an Excel file.</summary>
        /// <param name="data">Byte array containing Excel file data.  This is typically what you will have when reading data from a database.</param>
        /// <param name="worksheetName">Tab / Worksheet name</param>
        /// <param name="hasHeaderRow">Is there a header row with the names of the columns on the first row?</param>
        /// <param name="startRow">Starting row.  This is a ONE based number, but keep in mind that if hasHeaderRow is true, the header row IS row 1.
        /// If the value is null, it is not used.</param>
        /// <param name="endRow">Ending row.  This is a ONE based number, but keep in mind that if hasHeaderRow is true, the header row IS row 1.
        /// If the value is null, it is not used.</param>
        /// <returns>List of objects</returns>
        public List<T> ReadWorksheet(byte[] data, string worksheetName, bool hasHeaderRow = true, int? startRow = null, int? endRow = null)
        {
            using (var ms = new MemoryStream(data))
            {
                return ReadWorksheet(ms, worksheetName, hasHeaderRow, startRow, endRow);
            }
        }

        /// <summary>Reads data from a work sheet.</summary>
        /// <param name="dataStream">The stream that contains the worksheet</param>
        /// <param name="worksheetName">The worksheet name/tab that you want to read.</param>
        /// <param name="hasHeaderRow">Indicates if the first row in the worksheet is a header row.  If it does, we will use it
        /// to determine how things should map to the class; otherwise, we will just assue the class is indexed properly with
        /// the ExcelMappperAtttribute (Order property) and try to pull the data out that way.</param>
        /// <param name="startRow">Starting row.  This is a ONE based number, but keep in mind that if hasHeaderRow is true, the header row IS row 1.  
        /// If the value is null, it is not used.</param>
        /// <param name="endRow">Ending row.  This is a ONE based number, but keep in mind that if hasHeaderRow is true, the header row IS row 1.  
        /// If the value is null, it is not used.</param>
        /// <returns>List of objects</returns>
        public List<T> ReadWorksheet(Stream dataStream, string worksheetName, bool hasHeaderRow = true, int? startRow = null, int? endRow = null)
        {
            var result = new List<T>();
            
            if (dataStream == null)
                throw new ArgumentException("You must specify data!");

            if (string.IsNullOrEmpty(worksheetName))
                throw new ArgumentException("You must specify a worksheet name!");

            List<ClassToExcelColumn> columns = CreateColumns(worksheetName, typeof(T));
            
            //Open the Excel file
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(dataStream, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                SharedStringTable sst = sstpart != null ? sstpart.SharedStringTable : null;

                Sheet theSheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == worksheetName);
                if (theSheet == null)
                {
                    return result;
                }

                WorksheetPart theWorksheetPart = (WorksheetPart)workbookPart.GetPartById(theSheet.Id);
                Worksheet theWorksheet = theWorksheetPart.Worksheet;
                var rows = theWorksheet.Descendants<Row>();

                //Loop through the Worksheet rows.
                int rowIndex = 0;
                foreach (Row oneRow in rows)
                {
                    rowIndex++;

                    // Use the first row to add columns to DataTable.
                    if (hasHeaderRow && rowIndex == 1)
                    {
                        RemapExcelColumnLetters(oneRow, columns, sst);
                    }

                    if (startRow.HasValue && startRow.Value > rowIndex)
                        continue;
                    
                    if (hasHeaderRow == false || rowIndex > 1)
                    {
                        //Add rows to DataTable.
                        var newItem = new T();

                        foreach (Cell oneCell in oneRow.Elements<Cell>())
                        {
                            string columnLetter = GetColumnLetter(oneCell);
                            ClassToExcelColumn column = columns.FirstOrDefault(w => w.ExcelColumnLetter == columnLetter);
                            AssignValue(newItem, column, oneCell, sst, rowIndex);
                        }

                        result.Add(newItem);
                    }

                    if (endRow.HasValue && endRow.Value == rowIndex)
                        break;
                }
            }

            return result;
        }

        /// <summary>Use PropertyInfo and column information assign a property a value based on incoming cell data.</summary>
        /// <param name="newItem">The entire object</param>
        /// <param name="column">Information about the property that corresponds to the column (cell)</param>
        /// <param name="oneCell">A single OpenXML cell with data in it</param>
        /// <param name="sst">The shared strings table</param>
        /// <param name="rowIndex">The row number</param>
        private void AssignValue(T newItem, ClassToExcelColumn column, Cell oneCell, SharedStringTable sst, int rowIndex)
        {
            if (column == null || column.Property == null || oneCell.CellValue == null)
                return;

            string stringValue = GetCellText(oneCell, sst);
            if (string.IsNullOrEmpty(stringValue))
                return;

            if (column.Property.PropertyType == typeof(string))
                column.Property.SetValue(newItem, stringValue, null);
            else if (column.Property.PropertyType == typeof(DateTime) || column.Property.PropertyType == typeof(DateTime?))
                AssignDateTimeValue(newItem, column, rowIndex, stringValue);
            else
            {
                if (column.Property.PropertyType == typeof(int) || column.Property.PropertyType == typeof(int?))
                    AssignIntValue(newItem, column, rowIndex, stringValue);
                else if (column.Property.PropertyType == typeof(double) || column.Property.PropertyType == typeof(double?))
                    AssignDoubleValue(newItem, column, rowIndex, stringValue);
                else if (column.Property.PropertyType == typeof(decimal) || column.Property.PropertyType == typeof(decimal?))
                    AssignDecimalValue(newItem, column, rowIndex, stringValue);
                else if (column.Property.PropertyType == typeof(bool) || column.Property.PropertyType == typeof(bool?))
                    AssignBoolValue(newItem, column, rowIndex, stringValue);
                else LogMessage(new ClassToExcelMessage(ClassToExcelMessageType.TypeNotSupportedProblem, rowIndex, column, 
                    string.Format("Unknown property type '{0}' seen in AssignValue", column.Property.PropertyType)));
            }
        }

        private void AssignBoolValue(T newItem, ClassToExcelColumn column, int rowIndex, string stringValue)
        {
            if (string.IsNullOrWhiteSpace(stringValue) == false)
            {
                bool booleanValue;
                if (bool.TryParse(stringValue, out booleanValue))
                {
                    column.Property.SetValue(newItem, booleanValue, null);
                }
                else
                {
                    var lower = stringValue.Trim().ToLower();
                    if (lower == "y" || lower == "true" || lower == "t" || lower == "1")
                        column.Property.SetValue(newItem, true, null);
                    else
                    {
                        LogMessage(new ClassToExcelMessage(ClassToExcelMessageType.ParseProblem, 
                            rowIndex, column, string.Format("Assuming '{0}' is false since parsing failed.", stringValue)));
                        column.Property.SetValue(newItem, false, null);
                    }
                }
            }
        }

        private void AssignDateTimeValue(T newItem, ClassToExcelColumn column, int rowIndex, string stringValue)
        {
            if (string.IsNullOrWhiteSpace(stringValue) == false)
            {
                double value;
                if (double.TryParse(stringValue, out value))
                {
                    DateTime fromOaDate = DateTime.FromOADate(value);
                    column.Property.SetValue(newItem, fromOaDate, null);
                }
                else
                {
                    LogMessage(new ClassToExcelMessage(ClassToExcelMessageType.ParseProblem, 
                        rowIndex, column, string.Format("Could not parse '{0}' as an DateTime!", stringValue)));
                }
            }
        }

        private void AssignDecimalValue(T newItem, ClassToExcelColumn column, int rowIndex, string stringValue)
        {
            if (string.IsNullOrWhiteSpace(stringValue) == false)
            {
                decimal number;
                if (decimal.TryParse(stringValue, out number))
                {
                    column.Property.SetValue(newItem, number, null);
                }
                else
                {
                    string convertSpecialStrings = ConvertSpecialStrings(stringValue);
                    if (string.CompareOrdinal(convertSpecialStrings, stringValue) == 0)
                    {
                        LogMessage(new ClassToExcelMessage(ClassToExcelMessageType.ParseProblem, 
                            rowIndex, column, string.Format("Could not parse '{0}' as a decimal!", stringValue)));
                    }
                    else AssignDecimalValue(newItem, column, rowIndex, convertSpecialStrings);
                }
            }
        }

        private void AssignDoubleValue(T newItem, ClassToExcelColumn column, int rowIndex, string stringValue)
        {
            if (string.IsNullOrWhiteSpace(stringValue) == false)
            {
                double number;
                if (double.TryParse(stringValue, out number))
                {
                    column.Property.SetValue(newItem, number, null);
                }
                else
                {
                    string convertSpecialStrings = ConvertSpecialStrings(stringValue);
                    if (string.CompareOrdinal(convertSpecialStrings, stringValue) == 0)
                    {
                        LogMessage(new ClassToExcelMessage(ClassToExcelMessageType.ParseProblem, 
                            rowIndex, column, string.Format("Could not parse '{0}' as a double!", stringValue)));
                    }
                    else AssignDoubleValue(newItem, column, rowIndex, convertSpecialStrings);
                }
            }
        }

        private void AssignIntValue(T newItem, ClassToExcelColumn column, int rowIndex, string stringValue)
        {
            if (string.IsNullOrWhiteSpace(stringValue) == false)
            {
                int number;
                if (int.TryParse(stringValue, out number))
                {
                    column.Property.SetValue(newItem, number, null);
                }
                else
                {
                    LogMessage(new ClassToExcelMessage(ClassToExcelMessageType.ParseProblem, 
                        rowIndex, column, string.Format("Could not parse '{0}' as an int!", stringValue)));
                }
            }
        }

        private string ConvertSpecialStrings(string stringValue)
        {
            if (stringValue.Contains("E-"))
            {
                double number;
                if (double.TryParse(stringValue, NumberStyles.Float, null, out number) == false)
                    return stringValue;

                return number.ToString(CultureInfo.InvariantCulture);
            }

            if (stringValue.Contains("%"))
            {
                string replace = stringValue.Remove('%');
                double number;
                if (double.TryParse(replace, out number) == false)
                    return stringValue;

                number = number / 100.0;

                return number.ToString(CultureInfo.InvariantCulture);
            }

            return stringValue;
        }

        /// <summary>If the user has specified that there is a header row, remap the class properties based 
        /// on the columns in the row</summary>
        /// <param name="headerRow">The top (first) row with header text</param>
        /// <param name="columns">Information about the property that corresponds to the column (cell)</param>
        /// <param name="sst"></param>
        private void RemapExcelColumnLetters(Row headerRow, List<ClassToExcelColumn> columns, SharedStringTable sst)
        {
            // Erase all existing column letters
            foreach (var column in columns)
            {
                column.ExcelColumnLetter = string.Empty;
            }

            // Assign new column letter
            foreach (Cell oneCell in headerRow.Elements<Cell>())
            {
                if (oneCell.CellValue == null)
                    continue;
                
                string columnName = GetCellText(oneCell, sst);
                ClassToExcelColumn column = columns.FirstOrDefault(w => string.CompareOrdinal(w.ColumnName, columnName) == 0);
                if (column != null)
                {
                    column.ExcelColumnLetter = GetColumnLetter(oneCell);
                }
            }

            // Log problems
            foreach (var column in columns)
            {
                if (string.IsNullOrWhiteSpace(column.ExcelColumnLetter) && column.IsOptional == false)
                    LogMessage(new ClassToExcelMessage(ClassToExcelMessageType.HeaderProblem, 1, column, "Could not find column in the header row!"));
            }
        }
    }
}
