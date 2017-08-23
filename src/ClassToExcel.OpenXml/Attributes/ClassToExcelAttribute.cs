using System;

namespace ClassToExcel
{
    public class ClassToExcelAttribute : Attribute
    {
        public ClassToExcelAttribute()
        {
            Order = 999999;
            Ignore = false;
        }

        /// <summary>The name of the column in the Excel spreadsheet.  
        /// This is used when reading and writing files.</summary>
        public string ColumnName { get; set; }

        /// <summary>Ignore this property.  
        /// This is used when reading and writing files.</summary>
        public bool Ignore { get; set; }

        /// <summary>Is an optional property.  If specified, we will not report this as a missing header column 
        /// if you don't find it in the Excel file.  If you aren't listening for messages when reading data 
        /// or your data has no header row, you don't need to use this at all.
        /// This is used only when reading files.</summary> 
        public bool IsOptional { get; set; }

        /// <summary>The order that the column should be written to the excel spread sheet.  This is used when reading if there is NOT a header row.    
        /// This is used when reading and writing files.</summary>
        public int Order { get; set; }

        /// <summary>Used for writing Excel spreadsheets.  Use it when a special number or date format is desired. 
        /// This will only be used if the type is DateTime, int, decimal or double.    
        /// This is used only when writing files.</summary>
        public string StyleFormat { get; set; }

    }
}