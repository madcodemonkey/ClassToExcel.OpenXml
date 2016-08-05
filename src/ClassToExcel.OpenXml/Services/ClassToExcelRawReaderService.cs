using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClassToExcel
{
    /// <summary>Reads everything as a string with no coversion at all.  Keep in mind that Excel does NOT require 
    /// that BLANK cells be written so rows will not always have every column in their data!</summary>
    public class ClassToExcelRawReaderService : ClassToExcelBaseService, IClassToExcelRawReaderService
    {
        public ClassToExcelRawReaderService() : base(null) { }
        public ClassToExcelRawReaderService(Action<ClassToExcelMessage> loggingCallback) : base(loggingCallback){}

        /// <summary>Read a worksheet from an Excel file.</summary>
        /// <param name="fileName">Name and path of the file</param>
        /// <param name="worksheetName">Tab / Worksheet name</param>
        /// <param name="startRow">Starting row (this is a ONE based number)</param>
        /// <param name="endRow">Ending row (this is a ONE based number)</param>
        /// <returns>List of objects</returns>
        public List<ClassToExcelRawRow> ReadWorksheet(string fileName, string worksheetName, int? startRow = null, int? endRow = null)
        {
            if (File.Exists(fileName) == false)
                throw new ArgumentException("The file does not exists: " + fileName);

            using (var fs = File.OpenRead(fileName))
            {
                return ReadWorksheet(fs, worksheetName, startRow, endRow);
            }
        }

        /// <summary>Read a worksheet from an Excel file.</summary>
        /// <param name="data">Byte array containing Excel file data.  This is typically what you will have when reading data from a database.</param>
        /// <param name="worksheetName">Tab / Worksheet name</param>
        /// <param name="startRow">Starting row (this is a ONE based number)</param>
        /// <param name="endRow">Ending row (this is a ONE based number)</param>
        /// <returns>List of rows</returns>
        public List<ClassToExcelRawRow> ReadWorksheet(byte[] data, string worksheetName, int? startRow = null, int? endRow = null)
        {
            using (var ms = new MemoryStream(data))
            {
                return ReadWorksheet(ms, worksheetName, startRow, endRow);
            }
        }

        /// <summary>Reads data from a work sheet.</summary>
        /// <param name="dataStream">The stream that contains the worksheet</param>
        /// <param name="worksheetName">The worksheet name/tab that you want to read.</param>
        /// <param name="startRow">Starting row (this is a ONE based number)</param>
        /// <param name="endRow">Ending row (this is a ONE based number)</param>
        /// <returns>List of rows</returns>
        public List<ClassToExcelRawRow> ReadWorksheet(Stream dataStream, string worksheetName, int? startRow = null, int? endRow = null)
        {
            var result = new List<ClassToExcelRawRow>();
            
            if (dataStream == null)
                throw new ArgumentException("You must specify data!");

            if (string.IsNullOrEmpty(worksheetName))
                throw new ArgumentException("You must specify a worksheet name!");
            
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
                    if (startRow.HasValue && startRow.Value > rowIndex)
                        continue;

                    //Add rows to DataTable.
                    var newRow = new ClassToExcelRawRow {RowIndex = rowIndex};

                    foreach (Cell oneCell in oneRow.Elements<Cell>())
                    {
                        var newColumn = new ClassToExcelRawColumn();
                        newColumn.ColumnLetter = GetColumnLetter(oneCell);
                        newColumn.Data = GetCellText(oneCell, sst);
                        newRow.Columns.Add(newColumn);
                    }

                    result.Add(newRow);

                    if (endRow.HasValue && endRow.Value == rowIndex)
                        break;
                }
            }

            return result;
        }
    }
}
