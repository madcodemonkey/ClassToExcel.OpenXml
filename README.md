# ClassToExcel
A simple .NET library for writing and reading Excel worksheets using a List of T as the input or received output of the operation.

Notes
* Since these reading and writing classes are performing operations in memory, I do NOT recommend using this code to produce large Excel files with lots of rows.
* Requires the use of Microsoft's OpenXML library:  DocumentFormat.OpenXML

---
---
# ClassToExcelWriterService and ClassToExcelReaderService
The ClassToExcelWriterService service can be used to turn a List\<T\> into an Excel spreadsheet where each class represents a row and certain properties on the class are mapped to columns in the Excel spreadsheet. The ClassToExcelReaderService service can be used to do the reverse action.  It takes each row in the Excel spreadsheet and maps the columnd to properties on a class which results in a List\<T\>.

---
## Step 1: Create instructions
Create a class that will hold your data.

---
## Step 2: Decorate instructions
Decorate the properties with the ClassToExcelAttribute. The ClassToExcelAttribute has the following properties:
* ColumnName - This is used when reading or writing files where a header row is present.
* Order - I highly recommend specifying the Order.  When writing, it specifies the order that the column is
written to the file (i.e., The property marked with Order = 1 would be in Column A, etc.). If you do not specify order when 
writing a file, the data will written alphabetically using ColumnName, if specified, or property name.  When reading files, 
Order is used if you specify that no header column is present in the file. Again, if order isn't specified, the properties 
are sorted by ColunName, if specified, or property name and then read in that order.
* SyleFormat - This is used only when writing files.  It is used to style the column in the spreadsheet.
* Ignore - This is used when reading or writing files.  It means you don't want data read into the property and 
you don't want it written to the file.
* IsOptional - This is used only when reading files.  It is an optional property.  If specified, we will not report this as a missing header column if you don't find it in the Excel file.  If you aren't listening for messages when reading data or your data has no header row, you don't need to use this at all.

Here is an example:
```c#
public class Chicken
{
    public Chicken()
    {
        Owner = new Person { FirstName = "Farmer", LastName = "Brown", Age = 56};
    }

    [ClassToExcel(ColumnName = "Bird Name", Order = 1)]
    public string Name { get; set; }

    [ClassToExcel(Order = 2)]
    public int Age { get; set; }

    [ClassToExcel(ColumnName = "Date of Birth", Order = 3)]
    public DateTime DOB { get; set; }

    [ClassToExcel(StyleFormat = "$ #,###0.00", Order = 4)]
    public decimal Value { get; set; }

    [ClassToExcel(ColumnName = "Size of pen (sq. ft.)", StyleFormat = "#,###0", Order = 5)]
    public int SizeOfPenInSquareFeet { get; set; }
        
    /// <summary>Will be ignored.</summary>
    public Person Owner { get; set; }
}
```

Here is a list of built in style format supported by Excel:
* General
* 0
* 0.00
* #,##0
* #,##0.00
* 0%
* 0.00%
* 0.00E+00
* \# ?/?
* \# ??/??
* d/m/yyyy
* d-mmm-yy
* d-mmm
* mmm-yy
* h:mm tt
* h:mm:ss tt
* H:mm
* H:mm:ss
* m/d/yyyy H:mm
* #,##0 ;(#,##0)
* #,##0 ;\[Red\](#,##0)
* #,##0.00;(#,##0.00)
* #,##0.00;\[Red\](#,##0.00)
* mm:ss
* \[h\]:mm:ss
* mmss.0
* ##0.0E+0
* @

Other examples that work with Excel
* mm/dd/yy
* _($* #,##0.00_);_($* (#,##0.00);_($* "" - ""??_);_(@_)   

---
## Step 3a: Write instructions
Create a ClassToExcelWriterService object, add your worksheet data (list of instantiated objects from step 1) 
and then save it as a file or request the data in a MemoryStream.
```c#
List<Chicken> chickens = new List<Chicken>();
List<Person> people = new List<Person>();
// ..... populate chickens and people with data......
// File example
// Passing in LogServiceMessage is optional.  I'm using it here to help debug any issues while writing a file.
using (var writerService = new ClassToExcelWriterService(LogServiceMessage))
{
	writerService.AddWorksheet(people, "People", true);
	writerService.AddWorksheet(chickens, "Chickens", true);
	writerService.SaveToFile("c:\\temp\\Example1.xlsx");
}
```
OR
```c#
List<Chicken> chickens = new List<Chicken>();
List<Person> people = new List<Person>();
// ..... populate chickens and people with data......
// Stream example
using (var writerService = new ClassToExcelWriterService())
{
	writerService.AddWorksheet(people, "People", true);
	writerService.AddWorksheet(chickens, "Chickens", true);
	// YOU the caller must dispose of the MemoryStream you receive!
	MemoryStream ms = writerService.SaveToMemoryStream();
}
```

Notes
* If you call SaveToMemoryStream, YOU are responsible for disposing of the MemoryStream returned.
* ClassToExcelWriterService is disposable, but the ClassToExcelReaderService is not.  
Disposing of the writer will clean up OpenXML SpreadsheetDocument and other objects. 
* Objects and arrays will NOT be written to the file.

---
## Step 3b: Read instructions
Create a ClassToExcelReaderService object specifying the class that will be used.  Call ReadWorksheet with an 
input (file path, stream or byte array), the name of the tab/worksheet, and specify if the first row is a header row 
(warning, if you say "no" the Order attribute will be used to find data if order is not specified on the class
alphabetical order by ColumnName \[property name if ColumnName not specified\] will be assumed and data may be mapped incorrectly).

```c#
// File example
var readerService = new ClassToExcelReaderService<Person>();
List<Person> people = readerService.ReadWorksheet("c:\\temp\\Example1.xlsx", "People", true);
```

Notes
* ClassToExcelWriterService is disposable, but the ClassToExcelReaderService is not.
* If you pass in a stream, you still have dispose of it yourself.
* Objects and Arrays will be IGNORED on the class you created in Step 1.
* This code uses reflection, so it can set a value on a private setter in the class you created in Step 1.

---
---
# Optional: ClassToExcelRawReaderService (Raw Read instructions)
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
* The ClassToExcelRawReaderService service returns a list of row objects and each row object contains a list of column objects.  
The column object contains the column letter and the data in the column represented as a RAW (unparsed) string.  
In other words, things like dates will come out as numbers, which is how they are represented behind the scene.  
* If you need to convert dates from this number format into a real date, you can use  DateTime.FromOADate(someNumber) to convert it.
* ClassToExcelWriterService is disposable, but the ClassToExcelRawReaderService is not.


---
---
# Optional: ClassToExcelRowConverter
This coverter can be used in conjunction with the ClassToExcelRawReaderService and the ClassToExcelRowAttribute to map several rows onto a single class.  I don't expect this to be used very often, but there are a few edge cases where I needed to do somethhing like read rows 1 - 4 to one class and 5 - 8 to another.

## Step 1 
Run the example and do a Raw read to see how the data is going to look and take note that blank lines are eliminated by the OpenXML framework.
```
Row: 1 --> [Column: A Data: Beverages][Column: B Data: Column1][Column: C Data: Quantity]
Row: 2 --> [Column: A Data: Beer][Column: C Data: 1]
Row: 3 --> [Column: A Data: Wine][Column: C Data: 2]
Row: 4 --> [Column: A Data: Pepsi][Column: C Data: 3]
Row: 5 --> [Column: A Data: Coke][Column: C Data: 4]
Row: 6 --> [Column: A Data: Dr. Pepper][Column: C Data: 5]
Row: 7 --> [Column: A Data: Dates][Column: B Data: Start Date][Column: C Data: End Date]
Row: 8 --> [Column: A Data: Beer availability date][Column: B Data: 41883][Column: C Data: 41896]
Row: 9 --> [Column: A Data: Wine availability date][Column: B Data: 42798][Column: C Data: 42827]
```

## Step 2
Create your classes and decorate it with them ClassToExcelRowAttribute.

```c#
// Row 1 is a header
public class Beverage
{
	[ClassToExcelRow(ColumnLetter = "C", RowNumber = 2)]
	public int Beer { get; set; }
	[ClassToExcelRow(ColumnLetter = "C", RowNumber = 3)]
	public int Wine { get; set; }
	[ClassToExcelRow(ColumnLetter = "C", RowNumber = 4)]
	public int Pepsi { get; set; }
	[ClassToExcelRow(ColumnLetter = "C", RowNumber = 5)]
	public int Coke { get; set; }
	[ClassToExcelRow(ColumnLetter = "C", RowNumber = 6)]
	public int DrPepper { get; set; }
}

// Row 7 is a header and the blank row between the data is eliminated by OpenXML framework.
public class BeverageDates
{
	[ClassToExcelRow(ColumnLetter = "B", RowNumber = 8)]
	public DateTime BeerStart { get; set; }
	[ClassToExcelRow(ColumnLetter = "C", RowNumber = 8)]
	public DateTime BeerEnd { get; set; }

	[ClassToExcelRow(ColumnLetter = "B", RowNumber = 9)]
	public DateTime WineStart { get; set; }
	[ClassToExcelRow(ColumnLetter = "C", RowNumber = 9)]
	public DateTime WineEnd { get; set; }
}
```

## Step 3
```c#
string filePath = "C:\\Temp\\Beverages.xlsx";
if (File.Exists(filePath))
{
	var worksheetName = "Test";

	// Passing in LogServiceMessage in optional.  I'm using it here to help debug in issues while reading a file.
	var rawReaderService = new ClassToExcelRawReaderService(LogServiceMessage);
	List<ClassToExcelRawRow> rawRows = rawReaderService.ReadWorksheet(filePath, worksheetName);
	if (rawRows.Count > 0)
	{
		// Rows 2-6 (Column C) have the Beverage class data
		var converter1 = new ClassToExcelRowConverter<Beverage>(LogServiceMessage);
		Beverage beverage = converter1.Convert(rawRows);
		Console.WriteLine(SerializationHelper.SerializeToXml(beverage));

		// Rows 8-9 (Columns B & C) have the BeverageDates class data
		var converter2 = new ClassToExcelRowConverter<BeverageDates>(LogServiceMessage);
		BeverageDates dates = converter2.Convert(rawRows);
		Console.WriteLine(SerializationHelper.SerializeToXml(dates));

	}
	else Console.WriteLine("No data found.");
 }
 else Console.WriteLine($"The file does not exists: {filePath}");	
 ```