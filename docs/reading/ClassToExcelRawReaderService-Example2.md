# ClassToExcelRawReaderService - EXAMPLE 2

The ClassToExcelRawReaderService service is used for doing a raw dump of an Excel spreadsheet where every column's data is expressed as a string.

First, create an ClassToExcelRawReaderService object and then call the ReadWorksheet method with an input (file path, stream or byte array) and the name of the tab/worksheet.

```c#
// Passing in LogServiceMessage is optional.  I'm using it here to help debug any issues while reading a file.
var rawReaderService = new ClassToExcelRawReaderService(LogServiceMessage);
List<ClassToExcelRawRow> rawRows = rawReaderService.ReadWorksheet("c:\\temp\\Example1.xlsx", "People");
if (rawRows.Count > 0)
{
	foreach (var row in rawRows)
	{
		StringBuilder sb = new StringBuilder();
		sb.AppendFormat("Row: {0} --> ", row.RowIndex);
		foreach (var column in row.Columns)
		{
			sb.AppendFormat("[Column: {0} Data: {1}]", column.ColumnLetter, column.Data);
		}
		Console.WriteLine(sb.ToString());
	}
}
else Console.WriteLine("No data found.");
```

Notes
* The ClassToExcelRawReaderService service returns a list of row objects and each row object contains a list of column objects.  The column object contains the column letter and the data in the column represented as a RAW (unparsed) string.  In other words, things like dates will come out as numbers, which is how they are represented behind the scene.  
* If you need to convert dates from this number format into a real date, you can use  DateTime.FromOADate(someNumber) to convert it.
* ClassToExcelWriterService is disposable, but the ClassToExcelRawReaderService is not.
* Sometimes the OpenXML framework will remove a blank line and other times it will not (perhaps there is space or something that it recognizes).

