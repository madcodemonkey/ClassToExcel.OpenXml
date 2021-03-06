# ClassToExcel
A simple .NET library for writing and reading Excel worksheets using a List of T as the input or received output of the operation.

Notes
* Since these reading and writing classes are performing operations in memory, I do NOT recommend using this code to produce large Excel files with lots of rows.
* Requires the use of Microsoft's OpenXML library:  DocumentFormat.OpenXML

---
Follow these links to learn more about **commonly** used services:
- [Reading Excel files with ClassToExcelReaderService](./docs/reading/ClassToExcelReaderService.md)
- [Writing Excel files with ClassToExcelWriterService](./docs/writing/ClassToExcelWriterService.md)

---
Follow these links to learn more about services you'll **rarely** use:
- [Reading Excel files with the ClassToExcelRawReaderService](./docs/reading/ClassToExcelRawReaderService.md) - Reads everything as a string with no coversion at all.  Keep in mind that Excel does NOT require  that BLANK cells be written so rows will not always have every column in the Excel data!
- [Reading Excel files with the ClassToExcelRawReaderService and ClassToExcelRowConverter](./docs/reading/ClassToExcelRowConverter.md) - The ClassToExcelRowConverter can be used in conjunction with the ClassToExcelRawReaderService and the ClassToExcelRowAttribute to map several rows onto a single class.  I don't expect this to be used very often, but there are a few edge cases where I needed to do somethhing like read rows 1 - 4 to one class and rows 5 - 8 into another class.
