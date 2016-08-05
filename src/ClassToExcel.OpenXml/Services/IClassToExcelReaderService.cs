using System.Collections.Generic;
using System.IO;

namespace ClassToExcel
{
    public interface IClassToExcelReaderService<T> where T : class, new()
    {
        /// <summary>Read a worksheet from an Excel file.</summary>
        /// <param name="fileName">Name and path of the file</param>
        /// <param name="worksheetName">Tab / Worksheet name</param>
        /// <param name="hasHeaderRow">Is there a header row with the names of the columns on the first row?</param>
        /// <param name="startRow">Starting row (this is a ONE based number)</param>
        /// <param name="endRow">Ending row (this is a ONE based number)</param>
        /// <returns>List of objects</returns>
        List<T> ReadWorksheet(string fileName, string worksheetName, bool hasHeaderRow = true, int? startRow = null, int? endRow = null);

        /// <summary>Read a worksheet from an Excel file.</summary>
        /// <param name="data">Byte array containing Excel file data.  This is typically what you will have when reading data from a database.</param>
        /// <param name="worksheetName">Tab / Worksheet name</param>
        /// <param name="hasHeaderRow">Is there a header row with the names of the columns on the first row?</param>
        /// <param name="startRow">Starting row (this is a ONE based number)</param>
        /// <param name="endRow">Ending row (this is a ONE based number)</param>
        /// <returns>List of objects</returns>
        List<T> ReadWorksheet(byte[] data, string worksheetName, bool hasHeaderRow = true, int? startRow = null, int? endRow = null);

        /// <summary>Reads data from a work sheet.</summary>
        /// <param name="dataStream">The stream that contains the worksheet</param>
        /// <param name="worksheetName">The worksheet name/tab that you want to read.</param>
        /// <param name="hasHeaderRow">Indicates if the first row in the worksheet is a header row.  If it does, we will use it
        /// to determine how things should map to the class; otherwise, we will just assue the class is indexed properly with
        /// the ExcelMappperAtttribute (Order property) and try to pull the data out that way.</param>
        /// <param name="startRow">Starting row (this is a ONE based number)</param>
        /// <param name="endRow">Ending row (this is a ONE based number)</param>
        /// <returns>List of objects</returns>
        List<T> ReadWorksheet(Stream dataStream, string worksheetName, bool hasHeaderRow = true, int? startRow = null, int? endRow = null);

        /// <summary>The columns created when worksheets were added.  These are based on the Class of T that you passed into the ReadWorksheet method and 
        /// NOT the column header row from the Excel file!  The dictionary key is the worksheet name.</summary>
        Dictionary<string, List<ClassToExcelColumn>> WorksheetColumns { get; }
    }
}