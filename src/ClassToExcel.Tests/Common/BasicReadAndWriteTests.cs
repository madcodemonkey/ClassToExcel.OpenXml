using System.Collections.Generic;
using ClassToExcel.Sample.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClassToExcel.Tests
{
    [TestClass]
    public class BasicReadAndWriteTests
    {
        [TestMethod]
        public void CanReadAndWriteRandomPersonDataWithHeaders()
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
        public void CanReadAndWriteRandomChickenDataWithHeaders()
        {
            const bool hasHeaderRow = true;

            // Arrange
            List<Chicken> expectedList = ChickenCreator.Work(20);

            // Act
            var saveAndReadHelper = new SaveAndReadHelper<Chicken, Chicken>();
            var actualList = saveAndReadHelper.SaveAndRead(expectedList, hasHeaderRow);

            // Assert
            ChickenHelper.CompareExpectedWithActualAndAssertIfNotEqual(expectedList, actualList);
        }

        [TestMethod]
        public void CanReadAndWriteRandomPersonDataWithNoHeaders()
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

        [TestMethod]
        public void CanReadAndWriteRandomChickenDataWithNoHeaders()
        {
            const bool hasHeaderRow = false;

            // Arrange
            List<Chicken> expectedList = ChickenCreator.Work(20);

            // Act
            var saveAndReadHelper = new SaveAndReadHelper<Chicken, Chicken>();
            var actualList = saveAndReadHelper.SaveAndRead(expectedList, hasHeaderRow);

            // Assert
            ChickenHelper.CompareExpectedWithActualAndAssertIfNotEqual(expectedList, actualList);
        }
    }
}
