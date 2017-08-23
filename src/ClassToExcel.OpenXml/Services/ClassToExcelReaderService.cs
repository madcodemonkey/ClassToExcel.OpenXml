using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClassToExcel
{
    /// <summary>Reads only primitives and maps them ont a given class type.  Assumes that columns map to properties.</summary>
    /// <typeparam name="T"></typeparam>
    public class ClassToExcelReaderService<T> : ClassToExcelBaseService, IClassToExcelReaderService<T> where T : class, new()
    {
        private readonly StringToPropertyConverter<T> _stringToPropertyConverter = new StringToPropertyConverter<T>();

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

            var resultType = _stringToPropertyConverter.AssignValue(column.Property, newItem, stringValue);
            if (resultType == StringToPropertyConverterEnum.Error || resultType == StringToPropertyConverterEnum.Warning)
            {
                var classToExcelMessageType = resultType == StringToPropertyConverterEnum.Error ? ClassToExcelMessageType.Error : ClassToExcelMessageType.Warning;
                LogMessage(new ClassToExcelMessage(classToExcelMessageType, rowIndex, column, _stringToPropertyConverter.LastMessage));
            }
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
