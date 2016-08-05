using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NumberingFormat = DocumentFormat.OpenXml.Spreadsheet.NumberingFormat;

namespace ClassToExcel
{
    public class ClassToExcelWriterService : ClassToExcelBaseService, IClassToExcelWriterService
    {
        public ClassToExcelWriterService() : base(null) { }
        public ClassToExcelWriterService(Action<ClassToExcelMessage> loggingCallback) : base(loggingCallback){ }

        /// <summary>Saves one worksheet in the spreadsheet.  You can add one or more.</summary>
        /// <typeparam name="T">The class type that contains data that will be added to the worksheet</typeparam>
        /// <param name="dataList">The data that will be added to the worksheet</param>
        /// <param name="sheetName">The name of the worksheet</param>
        /// <param name="createHeaderRow">Indicates if the worksheet has a header row.</param>
        /// <returns>A worksheet</returns>
        public void AddWorksheet<T>(List<T> dataList, string sheetName, bool createHeaderRow = true) where T : class
        {
            if (dataList == null)
                throw new ArgumentException("You must specify a list of data!");

            if (string.IsNullOrEmpty(sheetName))
                sheetName = string.Format("sheet{0}", _spreadSheetId);
 
            if (_spreadSheet == null)
            {
                _ms = new MemoryStream();
                _spreadSheet = SpreadsheetDocument.Create(_ms, SpreadsheetDocumentType.Workbook);

                // Add a WorkbookPart to the document.
                _workbookpart = _spreadSheet.AddWorkbookPart();
                _workbookpart.Workbook = new Workbook();
                SharedStringTablePart shareStringPart = _workbookpart.AddNewPart<SharedStringTablePart>();
                shareStringPart.SharedStringTable = new SharedStringTable();

                // Add Sheets to the Workbook.
                _sheets = _spreadSheet.WorkbookPart.Workbook.AppendChild(new Sheets());

                // Add minimal Stylesheet
                // http://stackoverflow.com/questions/2792304/how-to-insert-a-date-to-an-open-xml-worksheet
                var stylesPart = _workbookpart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet
                {
                    Fonts = new Fonts(new Font()),
                    Fills = new Fills(new Fill()),
                    Borders = new Borders(new Border()),
                    CellStyleFormats = new CellStyleFormats(new CellFormat()),
                    CellFormats = new CellFormats(
                            // empty one for index 0, seems to be required
                            new CellFormat()),
                    NumberingFormats = new NumberingFormats()
                };
            }


            List<ClassToExcelColumn> columns = CreateColumns(sheetName, typeof(T));
            CreateAdditionalCellFormats(columns);


            // For each Excel sheet (that has separate data) 
            // - a separate WorkSheetPart object is needed
            // - a separate WorkSheet object is needed
            // - a separate SheetData object is needed
            // - a separate Sheet object is needed
            // Ref: http://stackoverflow.com/a/22230453/97803
            // Also see: https://msdn.microsoft.com/en-us/library/office/cc881781.aspx

            // Add a WorkSheetPart 
            WorksheetPart worksheetPart = _workbookpart.AddNewPart<WorksheetPart>();

            // Append Sheet
            Sheet sheet = new Sheet
            {
                Id = _spreadSheet.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = _spreadSheetId++,
                Name = sheetName
            };
            _sheets.Append(sheet);

            // Add a Worksheet
            Worksheet worksheet = new Worksheet();
            worksheetPart.Worksheet = worksheet;

            // Create and add SheetData
            SheetData sheetData = CreateSheetData(dataList, createHeaderRow, columns);
            worksheet.AppendChild(sheetData);

            worksheet.Save();
        }

        public void Dispose()
        {
            if (_ms != null)
            {
                _ms.Dispose();
                _ms = null;
            }

            DisposeOfSpreadSheet(false);
        }

        /// <summary>Closes the spreadsheet off to editing and saves the data to a file.</summary>
        /// <param name="filePath">The file name and path to save it.</param>
        public void SaveToFile(string filePath)
        {
            using (MemoryStream ms = SaveToMemoryStream())
            using(FileStream fs = File.Create(filePath))
            {
                ms.WriteTo(fs);
            }
        }

        /// <summary>Closes the spreadsheet off to editing and saves the data to a MemoryStream that the caller is responsible for disposing.</summary>
        /// <returns>A MemoryStream that the caller is responsible for disposing.</returns>
        public MemoryStream SaveToMemoryStream()
        {
            DisposeOfSpreadSheet(true);
            
            MemoryStream msForCallerToDispose = _ms;
            msForCallerToDispose.Position = 0;
            _ms = null;
            return msForCallerToDispose;
        }


        private const uint NumberFormatDoesNotExist = 99999;
        private uint _numberFormatId = 164;  // Built in codes are below 164
        private MemoryStream _ms;
        private SpreadsheetDocument _spreadSheet;
        private WorkbookPart _workbookpart;
        private Sheets _sheets;
        private uint _spreadSheetId = 1;

        private void CreateAdditionalCellFormats(List<ClassToExcelColumn> columns)
        {
            foreach (ClassToExcelColumn column in columns)
            {
                if (string.IsNullOrWhiteSpace(column.StyleFormat))
                {
                    if (column.IsDate)
                        column.StyleFormat = "d/m/yyyy";
                    else continue;
                }

                column.StyleIndex = FindOrCreateStyleIndex(column);
            }
        }

        private static SheetData CreateSheetData<T>(List<T> dataList, bool createHeaderRow, List<ClassToExcelColumn> columns) where T : class
        {
            var sheetData = new SheetData();

            // create row
            uint rowIndex = 0;
            if (createHeaderRow)
            {
                rowIndex++;

                Row headeRow = sheetData.AppendChild(new Row { RowIndex = rowIndex });
                foreach (var dataColumn in columns)
                {
                    // Assign COLUMN NAME to the cell
                    Cell newCell = new Cell {
                        CellReference = String.Format("{0}1", dataColumn.ExcelColumnLetter) ,
                        DataType = CellValues.String,
                        CellValue = new CellValue(dataColumn.ColumnName)};
                    headeRow.Append(newCell);
                }
            }
            
            foreach (var row in dataList)
            {
                rowIndex++;
                Row newRow = sheetData.AppendChild(new Row {RowIndex = rowIndex });
                
                foreach (ClassToExcelColumn dataColumn in columns)
                {
                    object value = dataColumn.Property.GetValue(row, null);
                    if (value != null)
                    {
                        Cell newCell;
                        if (dataColumn.IsDate)
                        {
                            DateTime valueAsDate = (DateTime)value;

                            newCell = new Cell
                            {
                                CellReference = String.Format("{0}{1}", dataColumn.ExcelColumnLetter, rowIndex),
                                DataType = CellValues.Number,
                                StyleIndex = dataColumn.StyleIndex, 
                                CellValue = new CellValue(valueAsDate.ToOADate().ToString(CultureInfo.InvariantCulture))
                            };
                        }
                        else if (dataColumn.IsNumber)
                        {
                            newCell = new Cell
                            {
                                CellReference = String.Format("{0}{1}", dataColumn.ExcelColumnLetter, rowIndex),
                                DataType = CellValues.Number,
                                StyleIndex = dataColumn.StyleIndex,
                                CellValue = new CellValue(value.ToString())
                            };
                        }
                        else if (dataColumn.IsBoolean)
                        {
                            newCell = new Cell
                            {
                                CellReference = String.Format("{0}{1}", dataColumn.ExcelColumnLetter, rowIndex),
                                DataType = CellValues.Boolean,
                                StyleIndex = dataColumn.StyleIndex,
                                CellValue = new CellValue(value.ToString())
                            };
                        }
                        else
                        {
                            newCell = new Cell
                            {
                                CellReference = String.Format("{0}{1}", dataColumn.ExcelColumnLetter, rowIndex),
                                DataType = CellValues.String,
                                StyleIndex = dataColumn.StyleIndex,
                                CellValue = new CellValue(value.ToString())
                            };
                        }

                        newRow.Append(newCell);
                    }
                }
            }

            return sheetData;
        }


        private void DisposeOfSpreadSheet(bool closeItFirst)
        {
            if (_spreadSheet != null)
            {
                if (closeItFirst)
                    _spreadSheet.Close();
                _spreadSheet.Dispose();
                _spreadSheet = null;
                _workbookpart = null;
                _sheets = null;
                _spreadSheetId = 0;
            }
        }

        /// <summary>Finds or creates a style index.  Style indices point to the INDEX position of a cell format.  In other words,
        /// if I return 5 from this method, it is the 6th element (yes, 6 it is zero based array) in the CellFormats array that will tell Excel 
        /// how to format the column.</summary>
        private uint FindOrCreateStyleIndex(ClassToExcelColumn column)
        {
            uint styleIndex = 0;

            // There are a bunch of built in numeric (includes dates) formatting styles in Excel.  Look to see if the user is trying
            // to use one of them.  If they are, we don't need to create a custome numeric format; otherwise, we do.
            // Creating numeric formats helpful links:
            // http://stackoverflow.com/questions/16607989/numbering-formats-in-openxml-c-sharp
            uint numberFormatId = GetStandardNumericStyle(column.StyleFormat);
            
            if (numberFormatId == NumberFormatDoesNotExist)
            {
                for (int i = 0; i < _spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats.ChildElements.Count; i++)
                {
                    var data = (NumberingFormat) _spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats.ChildElements[i];
                    if (data.FormatCode == column.StyleFormat)
                    {
                        numberFormatId = data.NumberFormatId;
                        break;
                    }
                }

                if (numberFormatId == NumberFormatDoesNotExist)
                {
                    numberFormatId = _numberFormatId++;
                    var newNumberingFormat = new NumberingFormat
                    {
                        NumberFormatId = UInt32Value.FromUInt32(numberFormatId),
                        FormatCode = StringValue.FromString(column.StyleFormat)
                    };
                    _spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats.AppendChild(newNumberingFormat);
                }
            }

            // Ok, at this point we know the numberFormatId.  Is there already an existing CellFormat record point to it?
            // If so, great return its index in the CellFormats array; otherwise, create it.
            bool foundCellFormat = false;
            for (int i = 0; i < _spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements.Count; i++)
            {
                var data = (CellFormat)_spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements[i];

                if (data.NumberFormatId != null && data.NumberFormatId.Value == numberFormatId)
                {
                    styleIndex = (uint) i;
                    foundCellFormat = true;
                    break;
                }
            }

            if (foundCellFormat == false)
            {
                var newCellFormat = new CellFormat
                {
                    NumberFormatId = numberFormatId,
                    ApplyNumberFormat = true
                };

                _spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild(newCellFormat);
                styleIndex = (uint) _spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements.Count - 1;
            }
            
            return styleIndex;
        }

        private uint GetStandardNumericStyle(string styleFormat)
        {
            switch (styleFormat)
            {
                case "General":
                    return 0;
                case "0":
                    return 1;
                case "0.00":
                    return 2;
                case "#,##0":
                    return 3;
                case "#,##0.00":
                    return 4;
                case "0%":
                    return 9;
                case "0.00%":
                    return 10;
                case "0.00E+00":
                    return 11;
                case "# ?/?":
                    return 12;
                case "# ??/??":
                    return 13;
                case "d/m/yyyy":
                    return 14;
                case "d-mmm-yy":
                    return 15;
                case "d-mmm":
                    return 16;
                case "mmm-yy":
                    return 17;
                case "h:mm tt":
                    return 18;
                case "h:mm:ss tt":
                    return 19;
                case "H:mm":
                    return 20;
                case "H:mm:ss":
                    return 21;
                case "m/d/yyyy H:mm":
                    return 22;
                case "#,##0 ;(#,##0)":
                    return 37;
                case "#,##0 ;[Red](#,##0)":
                    return 38;
                case "#,##0.00;(#,##0.00)":
                    return 39;
                case "#,##0.00;[Red](#,##0.00)":
                    return 40;
                case "mm:ss":
                    return 45;
                case "[h]:mm:ss":
                    return 46;
                case "mmss.0":
                    return 47;
                case "##0.0E+0":
                    return 48;
                case "@":
                    return 49;

            }

            return 99999; // General
        }

        // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        // From https://msdn.microsoft.com/en-us/library/office/cc861607.aspx
        private int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }
    }
}
