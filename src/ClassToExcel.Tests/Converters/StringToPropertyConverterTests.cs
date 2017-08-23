using System;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClassToExcel.Tests.Converters
{
    [TestClass]
    public class StringToPropertyConverterTests
    {
        #region string
        [TestMethod]
        public void String_CanAssignValues_ValuesAssigned()
        {
            PropertyTester(StringToPropertyConverterEnum.Good, "SomeString", "12345", "12345");
            PropertyTester(StringToPropertyConverterEnum.Good, "SomeString", "Text with spaces", "Text with spaces");
            PropertyTester(StringToPropertyConverterEnum.Good, "SomeString", "", "");
            PropertyTester(StringToPropertyConverterEnum.Good, "SomeString", string.Empty, String.Empty);
            PropertyTester<string>(StringToPropertyConverterEnum.Good, "SomeString", null, null);
        }
        #endregion

        #region int and int?
        [TestMethod]
        public void Integer_CanAssignValues_ValuesAssigned()
        {
            PropertyTester(StringToPropertyConverterEnum.Good, "SomeInteger", 12345, "12345");
            PropertyTester(StringToPropertyConverterEnum.Default, "SomeInteger", 0, "");
            PropertyTester(StringToPropertyConverterEnum.Default, "SomeInteger", 0, String.Empty);
            PropertyTester(StringToPropertyConverterEnum.Default, "SomeInteger", 0, null);
        }

        [TestMethod]
        public void Integer_CanHandleNonIntegerStrings_ErrorReturnedAndZeroIsDefault()
        {
            PropertyTester<int>(StringToPropertyConverterEnum.Error, "SomeInteger", 0, "abc");
            PropertyTester<int>(StringToPropertyConverterEnum.Error, "SomeInteger", 0, "5488e4$#@#");
        }

        [TestMethod]
        public void NullableInteger_CanAssignValues_ValuesAssigned()
        {
            PropertyTester<int?>(StringToPropertyConverterEnum.Good, "SomeNullableInteger", 12345, "12345");
            PropertyTester<int?>(StringToPropertyConverterEnum.Default, "SomeNullableInteger", null, "");
            PropertyTester<int?>(StringToPropertyConverterEnum.Default, "SomeNullableInteger", null, String.Empty);
            PropertyTester<int?>(StringToPropertyConverterEnum.Default, "SomeNullableInteger", null, null);
        }

        [TestMethod]
        public void NullableInteger_CanHandleNonIntegerStrings_ErrorReturnedAndNullIsDefault()
        {
            PropertyTester<int?>(StringToPropertyConverterEnum.Error, "SomeNullableInteger", null, "abc");
            PropertyTester<int?>(StringToPropertyConverterEnum.Error, "SomeNullableInteger", null, "5488e4$#@#");
        }
        #endregion

        #region double and double?
        [TestMethod]
        public void Double_CanAssignValues_ValuesAssigned()
        {
            PropertyTester<double>(StringToPropertyConverterEnum.Good, "SomeDouble", 12345.25, "12345.25");
            PropertyTester<double>(StringToPropertyConverterEnum.Good, "SomeDouble", 0.023, "2.3%");
            PropertyTester<double>(StringToPropertyConverterEnum.Default, "SomeDouble", 0.0, "");
            PropertyTester<double>(StringToPropertyConverterEnum.Default, "SomeDouble", 0.0, String.Empty);
            PropertyTester<double>(StringToPropertyConverterEnum.Default, "SomeDouble", 0.0, null);
        }

        [TestMethod]
        public void Double_CanHandleNonIntegerStrings_ErrorReturnedAndZeroIsDefault()
        {
            PropertyTester<double>(StringToPropertyConverterEnum.Error, "SomeDouble", 0, "abc");
            PropertyTester<double>(StringToPropertyConverterEnum.Error, "SomeDouble", 0, "5488e4$#@#");
        }

        [TestMethod]
        public void NullableDouble_CanAssignValues_ValuesAssigned()
        {
            PropertyTester<double?>(StringToPropertyConverterEnum.Good, "SomeNullableDouble", 12345.254, "12345.254");
            PropertyTester<double?>(StringToPropertyConverterEnum.Good, "SomeNullableDouble", 0.023, "2.3%");
            PropertyTester<double?>(StringToPropertyConverterEnum.Default, "SomeNullableDouble", null, "");
            PropertyTester<double?>(StringToPropertyConverterEnum.Default, "SomeNullableDouble", null, String.Empty);
            PropertyTester<double?>(StringToPropertyConverterEnum.Default, "SomeNullableDouble", null, null);
        }

        [TestMethod]
        public void NullableDouble_CanHandleNonIntegerStrings_ErrorReturnedAndNullIsDefault()
        {
            PropertyTester<int?>(StringToPropertyConverterEnum.Error, "SomeNullableDouble", null, "abc");
            PropertyTester<int?>(StringToPropertyConverterEnum.Error, "SomeNullableDouble", null, "5488e4$#@#");
        }
        #endregion
        
        #region decimal and decimal?
        [TestMethod]
        public void Decimal_CanAssignValues_ValuesAssigned()
        {
            PropertyTester<decimal>(StringToPropertyConverterEnum.Good, "SomeDecimal", 12345.25m, "12345.25");
            PropertyTester<decimal>(StringToPropertyConverterEnum.Good, "SomeDecimal", 0.023m, "2.3%");
            PropertyTester<decimal>(StringToPropertyConverterEnum.Default, "SomeDecimal", 0.0m, "");
            PropertyTester<decimal>(StringToPropertyConverterEnum.Default, "SomeDecimal", 0.0m, String.Empty);
            PropertyTester<decimal>(StringToPropertyConverterEnum.Default, "SomeDecimal", 0.0m, null);
        }

        [TestMethod]
        public void Decimal_CanHandleNonIntegerStrings_ErrorReturnedAndZeroIsDefault()
        {
            PropertyTester(StringToPropertyConverterEnum.Error, "SomeDecimal", 0m, "abc");
            PropertyTester(StringToPropertyConverterEnum.Error, "SomeDecimal", 0m, "5488e4$#@#");
        }

        [TestMethod]
        public void NullableDecimal_CanAssignValues_ValuesAssigned()
        {
            PropertyTester<decimal?>(StringToPropertyConverterEnum.Good, "SomeNullableDecimal", 12345.254m, "12345.254");
            PropertyTester<decimal?>(StringToPropertyConverterEnum.Good, "SomeNullableDecimal", 0.023m, "2.3%");
            PropertyTester<decimal?>(StringToPropertyConverterEnum.Default, "SomeNullableDecimal", null, "");
            PropertyTester<decimal?>(StringToPropertyConverterEnum.Default, "SomeNullableDecimal", null, String.Empty);
            PropertyTester<decimal?>(StringToPropertyConverterEnum.Default, "SomeNullableDecimal", null, null);
        }

        [TestMethod]
        public void NullableDecimal_CanHandleNonIntegerStrings_ErrorReturnedAndNullIsDefault()
        {
            PropertyTester<int?>(StringToPropertyConverterEnum.Error, "SomeNullableDecimal", null, "abc");
            PropertyTester<int?>(StringToPropertyConverterEnum.Error, "SomeNullableDecimal", null, "5488e4$#@#");
        }
        #endregion

        #region bool and bool?
        [TestMethod]
        public void Bool_CanAssignValues_ValuesAssigned()
        {
            PropertyTester<bool>(StringToPropertyConverterEnum.Good, "SomeBoolean", true, "TRUE");
            PropertyTester<bool>(StringToPropertyConverterEnum.Good, "SomeBoolean", true, "True");
            PropertyTester<bool>(StringToPropertyConverterEnum.Good, "SomeBoolean", true, "true");
            PropertyTester<bool>(StringToPropertyConverterEnum.Good, "SomeBoolean", true, "t");
            PropertyTester<bool>(StringToPropertyConverterEnum.Good, "SomeBoolean", true, "y");
            PropertyTester<bool>(StringToPropertyConverterEnum.Good, "SomeBoolean", true, "1");

            PropertyTester<bool>(StringToPropertyConverterEnum.Good, "SomeBoolean", false, "FALSE");
            PropertyTester<bool>(StringToPropertyConverterEnum.Good, "SomeBoolean", false, "False");
            PropertyTester<bool>(StringToPropertyConverterEnum.Good, "SomeBoolean", false, "false");
            PropertyTester<bool>(StringToPropertyConverterEnum.Good, "SomeBoolean", false, "f");
            PropertyTester<bool>(StringToPropertyConverterEnum.Good, "SomeBoolean", false, "n");
            PropertyTester<bool>(StringToPropertyConverterEnum.Good, "SomeBoolean", false, "0");
        }

        [TestMethod]
        public void Bool_CanHandleNonIntegerStrings_ErrorReturnedAndZeroIsDefault()
        {
            PropertyTester(StringToPropertyConverterEnum.Error, "SomeBoolean", false, "abc");
            PropertyTester(StringToPropertyConverterEnum.Error, "SomeBoolean", false, "5488e4$#@#");
        }

        [TestMethod]
        public void NullableBool_CanAssignValues_ValuesAssigned()
        {
            PropertyTester<bool?>(StringToPropertyConverterEnum.Good, "SomeNullableBoolean", true, "TRUE");
            PropertyTester<bool?>(StringToPropertyConverterEnum.Good, "SomeNullableBoolean", true, "True");
            PropertyTester<bool?>(StringToPropertyConverterEnum.Good, "SomeNullableBoolean", true, "true");
            PropertyTester<bool?>(StringToPropertyConverterEnum.Good, "SomeNullableBoolean", true, "t");
            PropertyTester<bool?>(StringToPropertyConverterEnum.Good, "SomeNullableBoolean", true, "y");
            PropertyTester<bool?>(StringToPropertyConverterEnum.Good, "SomeNullableBoolean", true, "1");

            PropertyTester<bool?>(StringToPropertyConverterEnum.Good, "SomeNullableBoolean", false, "FALSE");
            PropertyTester<bool?>(StringToPropertyConverterEnum.Good, "SomeNullableBoolean", false, "False");
            PropertyTester<bool?>(StringToPropertyConverterEnum.Good, "SomeNullableBoolean", false, "false");
            PropertyTester<bool?>(StringToPropertyConverterEnum.Good, "SomeNullableBoolean", false, "f");
            PropertyTester<bool?>(StringToPropertyConverterEnum.Good, "SomeNullableBoolean", false, "n");
            PropertyTester<bool?>(StringToPropertyConverterEnum.Good, "SomeNullableBoolean", false, "0");
        }

        [TestMethod]
        public void NullableBool_CanHandleNonIntegerStrings_ErrorReturnedAndNullIsDefault()
        {
            PropertyTester<bool?>(StringToPropertyConverterEnum.Error, "SomeNullableBoolean", false, "abc");
            PropertyTester<bool?>(StringToPropertyConverterEnum.Error, "SomeNullableBoolean", false, "5488e4$#@#");
        }
        #endregion


        public class StringPropertTesting
        {
            public string SomeString { get; set; }
            public int SomeInteger { get; set; }
            public int? SomeNullableInteger { get; set; }
            public double SomeDouble { get; set; }
            public double? SomeNullableDouble { get; set; }
            public decimal SomeDecimal { get; set; }
            public decimal? SomeNullableDecimal { get; set; }
            public bool SomeBoolean { get; set; }
            public bool? SomeNullableBoolean { get; set; }
        }

        private void PropertyTester<T>(StringToPropertyConverterEnum expectedStatus, string propertyName, T expectedValue, string stringInputValue)
        {
            // Arrange
            PropertyInfo propertyInfo = typeof(StringPropertTesting).GetProperties().FirstOrDefault(w => w.Name == propertyName);
            Assert.IsNotNull(propertyInfo, $"Unable to find a property named '{propertyName}'");

            StringPropertTesting dataObject = new StringPropertTesting();
            StringToPropertyConverter<StringPropertTesting> converter = new StringToPropertyConverter<StringPropertTesting>();

            // Act
            StringToPropertyConverterEnum actualStatus = converter.AssignValue(propertyInfo, dataObject, stringInputValue);

            // Assert
            Assert.AreEqual(expectedStatus, actualStatus);
            Assert.AreEqual(expectedValue, propertyInfo.GetValue(dataObject));
        }



    }
}
