using System.Collections.Generic;
using ClassToExcel.Sample.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClassToExcel.Tests
{
    [TestClass]
    public class BasicReadAndWriteTests
    {
        [TestMethod]
        public void CanReadAndWriteRandomDataWithHeaders()
        {
            const bool hasHeaderRow = true;

            // Arrange
            List<Person> expectedList = PersonCreator.Work(20);

            // Act
            var saveAndReadHelper = new SaveAndReadHelper<Person, Person>();
            var actualList = saveAndReadHelper.SaveAndRead(expectedList, hasHeaderRow);

            // Assert
            PersonHelper.CompareExpectedWithActualAndAssertIfNotEqual(expectedList, actualList);
        }

        [TestMethod]
        public void CanReadAndWriteRandomDataWithNoHeaders()
        {
            const bool hasHeaderRow = false;

            // Arrange
            List<Person> expectedList = PersonCreator.Work(20);

            // Act
            var saveAndReadHelper = new SaveAndReadHelper<Person, Person>();
            var actualList = saveAndReadHelper.SaveAndRead(expectedList, hasHeaderRow);

            // Assert
            PersonHelper.CompareExpectedWithActualAndAssertIfNotEqual(expectedList, actualList);
        }
    }
}
