using System;
using System.Collections.Generic;
using ClassToExcel.Sample.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClassToExcel.Tests.Converters
{
    [TestClass]
    public class ClassToExcelRowConverterTests
    {
        [TestMethod]
        public void Convert_CanMapRowsToPropeties_ValuesTransferCorrectly()
        {
            // Arrange
            var rows = CreateRows();

            // Act
            var sut = new ClassToExcelRowConverter<Beverage>(null);
            var actualResult = sut.Convert(rows);

            // Assert
            Assert.AreEqual(1, actualResult.Beer);
            Assert.AreEqual(2, actualResult.Wine);
            Assert.AreEqual(3, actualResult.Pepsi);
            Assert.AreEqual(4, actualResult.Coke);
            Assert.AreEqual(5, actualResult.DrPepper);
        }

        [TestMethod]
        public void Convert_CanParseDifferentRowsToDifferentClass_ValuesTransferCorrectly()
        {
            // Arrange
            var rows = CreateRows();

            // Act
            var sut = new ClassToExcelRowConverter<BeverageDates>(null);
            var actualResult = sut.Convert(rows);

            // Assert
            Assert.AreEqual(new DateTime(2014, 8, 31), actualResult.BeerStart);
            Assert.AreEqual(new DateTime(2014, 9, 1), actualResult.BeerEnd);
            Assert.AreEqual(new DateTime(2014, 10, 31), actualResult.WineStart);
            Assert.AreEqual(new DateTime(2014, 11, 1), actualResult.WineEnd);
        }



        private List<ClassToExcelRawRow> CreateRows()
        {
            int rowNumber = 1;

            var rows = new List<ClassToExcelRawRow>();
            rows.Add(new ClassToExcelRawRow { RowNumber = rowNumber++, Columns = CreateColumns("Beverages", string.Empty, "Quantity") });
            rows.Add(new ClassToExcelRawRow { RowNumber = rowNumber++, Columns = CreateColumns("Beer", string.Empty, "1") });
            // Blank lines are ignored by the framework:  rows.Add(new ClassToExcelRawRow { RowNumber = rowNumber++, Columns = CreateColumns(string.Empty, string.Empty, string.Empty) });
            rows.Add(new ClassToExcelRawRow { RowNumber = rowNumber++, Columns = CreateColumns("Wine", string.Empty, "2") });
            rows.Add(new ClassToExcelRawRow { RowNumber = rowNumber++, Columns = CreateColumns("Pepsi", string.Empty, "3") });
            rows.Add(new ClassToExcelRawRow { RowNumber = rowNumber++, Columns = CreateColumns("Coke", string.Empty, "4") });
            rows.Add(new ClassToExcelRawRow { RowNumber = rowNumber++, Columns = CreateColumns("Dr. Pepper", string.Empty, "5") });
            // Blank lines are ignored by the framework: rows.Add(new ClassToExcelRawRow { RowNumber = rowNumber++, Columns = CreateColumns(string.Empty, string.Empty, string.Empty) });
            rows.Add(new ClassToExcelRawRow { RowNumber = rowNumber++, Columns = CreateColumns("Dates", "Start Date", "End Date") });
            rows.Add(new ClassToExcelRawRow { RowNumber = rowNumber++, Columns = CreateColumns("Beer availabilty Date", "41882", "41883") });
            // Blank lines are ignored by the framework: rows.Add(new ClassToExcelRawRow { RowNumber = rowNumber++, Columns = CreateColumns(string.Empty, string.Empty, string.Empty) });
            rows.Add(new ClassToExcelRawRow { RowNumber = rowNumber++, Columns = CreateColumns("Wine availabilty Date", "41943", "41944") });
            //
            return rows;
        }

        private List<ClassToExcelRawColumn> CreateColumns(string a, string b, string c)
        {
            var result = new List<ClassToExcelRawColumn>();
            result.Add(new ClassToExcelRawColumn {ColumnLetter = "A", Data = a});
            result.Add(new ClassToExcelRawColumn { ColumnLetter = "B", Data = b });
            result.Add(new ClassToExcelRawColumn { ColumnLetter = "C", Data = c });
            return result;
        }
    }
}

