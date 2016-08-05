using System.Collections.Generic;
using System.IO;

namespace ClassToExcel
{
    public interface IClassToExcelRawReaderService
    {
        /// <summary>Read a worksheet from an Excel file.</summary>
        /// <param name="fileName">Name and path of the file</param>
        /// <param name="worksheetName">Tab / Worksheet name</param>
        /// <param name="startRow">Starting row (this is a ONE based number)</param>
        /// <param name="endRow">Ending row (this is a ONE based number)</param>
        /// <returns>List of objects</returns>
        List<ClassToExcelRawRow> ReadWorksheet(string fileName, string worksheetName, int? startRow = null, int? endRow = null);

        /// <summary>Read a worksheet from an Excel file.</summary>
        /// <param name="data">Byte array containing Excel file data.  This is typically what you will have when reading data from a database.</param>
        /// <param name="worksheetName">Tab / Worksheet name</param>
        /// <param name="startRow">Starting row (this is a ONE based number)</param>
        /// <param name="endRow">Ending row (this is a ONE based number)</param>
        /// <returns>List of rows</returns>
        List<ClassToExcelRawRow> ReadWorksheet(byte[] data, string worksheetName, int? startRow = null, int? endRow = null);

        /// <summary>Reads data from a work sheet.</summary>
        /// <param name="dataStream">The stream that contains the worksheet</param>
        /// <param name="worksheetName">The worksheet name/tab that you want to read.</param>
        /// <param name="startRow">Starting row (this is a ONE based number)</param>
        /// <param name="endRow">Ending row (this is a ONE based number)</param>
        /// <returns>List of rows</returns>
        List<ClassToExcelRawRow> ReadWorksheet(Stream dataStream, string worksheetName, int? startRow = null, int? endRow = null);
    }
}