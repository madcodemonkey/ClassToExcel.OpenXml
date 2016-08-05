using System;
using System.Collections.Generic;
using System.IO;

namespace ClassToExcel.Tests
{
    public class SaveAndReadHelper<T,U> where T: class where U : class, new()
    {
        public List<U> SaveAndRead(List<T> expectedList, bool hasHeaderRow, 
            Action<ClassToExcelMessage> writeCallback = null, 
            Action<ClassToExcelMessage> readCallback = null) 
        {
            const string workSheetName = "Test";

            // WRITE FIRST
            IClassToExcelWriterService writer = new ClassToExcelWriterService(writeCallback);
            writer.AddWorksheet(expectedList, workSheetName, hasHeaderRow);
            MemoryStream ms = writer.SaveToMemoryStream();


            // READ AFTERWARDS
            List<U> actualList;
            using (ms)
            {
                var reader = new ClassToExcelReaderService<U>(readCallback);
                actualList = reader.ReadWorksheet(ms, workSheetName, hasHeaderRow);
            }

            return actualList;
        }

    }
}