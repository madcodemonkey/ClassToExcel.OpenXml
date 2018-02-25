# ClassToExcelRowConverter

The ClassToExcelRowConverter can be used in conjunction with the ClassToExcelRawReaderService and the ClassToExcelRowAttribute to map several rows onto a single class.  I don't expect this to be used very often, but there are a few edge cases where I needed to do somethhing like read rows 1 - 4 to one class and rows 5 - 8 into another class.

## Step 1 
Run the example program included here in Github and examime the Raw data to see how the data is going to look.  

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

Notes
* SOMETIMES blank lines are eliminated by the OpenXML framework.  This can make this type of parsing very fragile!!!  In these cases, I usually read a label or some other addition text from a different column to double check that I'm reading from the proper location.

## Step 2
Create your classes and decorate it with them ClassToExcelRowAttribute.

```c#
// Row 1 is a header, so start on row 2
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