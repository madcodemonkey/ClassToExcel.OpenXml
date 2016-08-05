using System;
using System.Collections.Generic;
using System.IO;

namespace ClassToExcel
{
    public interface IClassToExcelWriterService : IDisposable
    {
        /// <summary>Saves one worksheet in the spreadsheet.  You can add one or more.</summary>
        /// <typeparam name="T">The class type that contains data that will be added to the worksheet</typeparam>
        /// <param name="dataList">The data that will be added to the worksheet</param>
        /// <param name="sheetName">The name of the worksheet</param>
        /// <param name="createHeaderRow">Indicates if the worksheet has a header row.</param>
        /// <returns>A worksheet</returns>
        void AddWorksheet<T>(List<T> dataList, string sheetName, bool createHeaderRow = true) where T : class;

        /// <summary>Closes the spreadsheet off to editing and saves the data to a file.</summary>
        /// <param name="filePath">The file name and path to save it.</param>
        void SaveToFile(string filePath);

        /// <summary>Closes the spreadsheet off to editing and saves the data to a MemoryStream that the caller is responsible for disposing.</summary>
        /// <returns>A MemoryStream that the caller is responsible for disposing.</returns>
        MemoryStream SaveToMemoryStream();

        /// <summary>The columns created when worksheets were added.  These are based on the Class of T that you passed into the AddWorksheet method.
        /// The dictionary key is the worksheet name.</summary>
        Dictionary<string, List<ClassToExcelColumn>> WorksheetColumns { get; }
    }
}