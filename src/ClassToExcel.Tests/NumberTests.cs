using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClassToExcel.Tests
{
    [TestClass]
    public class NumberTests
    {
        [TestMethod]
        public void CanReadDecimals()
        {
            // Arrange
            // Arrange
            // Arrange
            var originalData = new List<NumberTestExampe1>();
            originalData.Add(new NumberTestExampe1 { Id = 1, DecimalValue = "123.30" });
            originalData.Add(new NumberTestExampe1 { Id = 2, DecimalValue = "-123.54" });
            originalData.Add(new NumberTestExampe1 { Id = 3, DecimalValue = Decimal.MaxValue.ToString() });
            originalData.Add(new NumberTestExampe1 { Id = 4, DecimalValue = Decimal.MinValue.ToString() });

            // Act
            // Act
            // Act
            // Using a different class with similar properties WITHOUT the ignore attribute to prove that nothing was truly written
            // if we used the same class the reader would not even look for the Name field.
            var saveAndReadHelper = new SaveAndReadHelper<NumberTestExampe1, NumberTestExampe2>();
            List<NumberTestExampe2> actualList = saveAndReadHelper.SaveAndRead(originalData, false);


            // Assert
            // Assert
            // Assert
            CompareDecimalValue(actualList, 1, 123.30m);
            CompareDecimalValue(actualList, 2, -123.54m);
            CompareDecimalValue(actualList, 3, decimal.MaxValue);
            CompareDecimalValue(actualList, 4, decimal.MinValue);
        }

        [TestMethod]
        public void CanReadDecimalWithExponentNotation()
        {
            // Arrange
            // Arrange
            // Arrange
            var originalData = new List<NumberTestExampe1>();
            originalData.Add(new NumberTestExampe1 { Id = 1, DecimalValue = "123.30" });
            originalData.Add(new NumberTestExampe1 { Id = 2, DecimalValue = "-1.6530975699424744E-8" });
            originalData.Add(new NumberTestExampe1 { Id = 3, DecimalValue = "1.6530E+5" });

            // Act
            // Act
            // Act
            // Using a different class with similar properties WITHOUT the ignore attribute to prove that nothing was truly written
            // if we used the same class the reader would not even look for the Name field.
            var saveAndReadHelper = new SaveAndReadHelper<NumberTestExampe1, NumberTestExampe2>();
            List<NumberTestExampe2> actualList = saveAndReadHelper.SaveAndRead(originalData, false);


            // Assert
            // Assert
            // Assert
            CompareDecimalValue(actualList, 1, 123.30m);
            CompareDecimalValue(actualList, 2, -.000000016530975699424744m);
            CompareDecimalValue(actualList, 3, 165300.00m);
        }


        [TestMethod]
        public void CanReadDoubles()
        {
            // Arrange
            // Arrange
            // Arrange
            var originalData = new List<NumberTestExampe1>();
            originalData.Add(new NumberTestExampe1 { Id = 1, DoubleValue = "123.30" });
            originalData.Add(new NumberTestExampe1 { Id = 2, DoubleValue = "-124.67" });
            // To round-trip the double, using "R" option
            // http://www.hanselman.com/blog/WhyYouCantDoubleParseDoubleMaxValueToStringOrSystemOverloadExceptionsWhenUsingDoubleParse.aspx
            originalData.Add(new NumberTestExampe1 { Id = 3, DoubleValue = double.MaxValue.ToString("R") });
            originalData.Add(new NumberTestExampe1 { Id = 4, DoubleValue = double.MinValue.ToString("R") });

            // Act
            // Act
            // Act
            // Using a different class with similar properties WITHOUT the ignore attribute to prove that nothing was truly written
            // if we used the same class the reader would not even look for the Name field.
            var saveAndReadHelper = new SaveAndReadHelper<NumberTestExampe1, NumberTestExampe2>();
            List<NumberTestExampe2> actualList = saveAndReadHelper.SaveAndRead(originalData, false);


            // Assert
            // Assert
            // Assert
            CompareDoubleValue(actualList, 1, 123.30d);
            CompareDoubleValue(actualList, 2, -124.67d);
            CompareDoubleValue(actualList, 3, double.MaxValue);
            CompareDoubleValue(actualList, 4, double.MinValue);
        }

        [TestMethod]
        public void CanReadDoubleWithPercentSign()
        {
            // Arrange
            // Arrange
            // Arrange
            var originalData = new List<NumberTestExampe1>();
            originalData.Add(new NumberTestExampe1 { Id = 1, DoubleValue = "56.1%" });
            originalData.Add(new NumberTestExampe1 { Id = 2, DoubleValue = "-45.2%" });
            originalData.Add(new NumberTestExampe1 { Id = 3, DoubleValue = "123%" });

            // Act
            // Act
            // Act
            // Using a different class with similar properties WITHOUT the ignore attribute to prove that nothing was truly written
            // if we used the same class the reader would not even look for the Name field.
            var saveAndReadHelper = new SaveAndReadHelper<NumberTestExampe1, NumberTestExampe2>();
            List<NumberTestExampe2> actualList = saveAndReadHelper.SaveAndRead(originalData, false);


            // Assert
            // Assert
            // Assert
            CompareDoubleValue(actualList, 1, .561d);
            CompareDoubleValue(actualList, 2, -.452d);
            CompareDoubleValue(actualList, 3, 1.23d);
        }



        [TestMethod]
        public void CanReadDoubleWithExponentNotation()
        {
            // Arrange
            // Arrange
            // Arrange
            var originalData = new List<NumberTestExampe1>();
            originalData.Add(new NumberTestExampe1 { Id = 1, DoubleValue = "123.30" });
            originalData.Add(new NumberTestExampe1 { Id = 2, DoubleValue = "-1.6530975699424744E-8" });
            originalData.Add(new NumberTestExampe1 { Id = 3, DoubleValue = "1.6530E+5" });
 
            // Act
            // Act
            // Act
            // Using a different class with similar properties WITHOUT the ignore attribute to prove that nothing was truly written
            // if we used the same class the reader would not even look for the Name field.
            var saveAndReadHelper = new SaveAndReadHelper<NumberTestExampe1, NumberTestExampe2>();
            List<NumberTestExampe2> actualList = saveAndReadHelper.SaveAndRead(originalData, false);


            // Assert
            // Assert
            // Assert
            CompareDoubleValue(actualList, 1, 123.30d);
            CompareDoubleValue(actualList, 2, -.000000016530975699424744d);
            CompareDoubleValue(actualList, 3, 165300.00d);
        }




        private static void CompareDecimalValue(List<NumberTestExampe2> actualList, int id, decimal expectedDecimal)
        {
            NumberTestExampe2 actual = actualList.FirstOrDefault(w => w.Id == id);
            Assert.IsNotNull(actual, $"Unable to find Id == {id} so data was not written properly");
            Assert.AreEqual(expectedDecimal, actual.DecimalValue, $"Expected decimal value for id {id} to be a be equal!");
        }

        private static void CompareDoubleValue(List<NumberTestExampe2> actualList, int id, double expectedDouble)
        {
            NumberTestExampe2 actual = actualList.FirstOrDefault(w => w.Id == id);
            Assert.IsNotNull(actual, $"Unable to find Id == {id} so data was not written properly");
            Assert.AreEqual(expectedDouble, actual.DoubleValue, $"Expected double value for id {id} to be a be equal!");
        }

        public class NumberTestExampe1
        {
            [ClassToExcel(Order = 1, ColumnName = "Id")]
            public int Id { get; set; }

            [ClassToExcel(Order = 2, ColumnName = "Dec Value")]
            public string DecimalValue { get; set; }

            [ClassToExcel(Order = 3, ColumnName = "Dou Value")]
            public string DoubleValue { get; set; }
        }

        public class NumberTestExampe2
        {
            [ClassToExcel(Order = 1, ColumnName = "Id")]
            public int Id { get; set; }

      
            [ClassToExcel(Order = 2, ColumnName = "Dec Value")]
            public decimal DecimalValue { get; set; }

            [ClassToExcel(Order = 3, ColumnName = "Dec Value")]
            public double DoubleValue { get; set; }
        }

    }
}
