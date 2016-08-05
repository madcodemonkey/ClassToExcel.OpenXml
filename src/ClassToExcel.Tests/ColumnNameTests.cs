using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClassToExcel.Tests
{
    /// <summary>
    /// Summary description for IgnoreTests
    /// </summary>
    [TestClass]
    public class ColumnNameTests
    {
        [TestMethod]
        public void PropertyNames_do_not_need_to_match_if_ClassToExcelAttribute_ColumnName_is_specified_and_the_file_contains_a_header_row()
        {
            // Arrange
            // Arrange
            // Arrange
            var originalData = new List<ColumnNameTestsExampe1>();
            originalData.Add(new ColumnNameTestsExampe1 { Id = 1, FName = "Tom", LName = "McMasters" });
            originalData.Add(new ColumnNameTestsExampe1 { Id = 2, FName = "David", LName = "Jones" });

            // Act
            // Act
            // Act
            // Using a different class with similar properties WITHOUT the ignore attribute to prove that nothing was truly written
            // if we used the same class the reader would not even look for the Name field.
            var saveAndReadHelper = new SaveAndReadHelper<ColumnNameTestsExampe1, ColumnNameTestsExampe2>();
            List<ColumnNameTestsExampe2> actualList = saveAndReadHelper.SaveAndRead(originalData, true);


            // Assert
            // Assert
            // Assert
            ColumnNameTestsExampe1 expected1 = originalData.FirstOrDefault(w => w.Id == 1);
            ColumnNameTestsExampe2 actual1 = actualList.FirstOrDefault(w => w.Id == 1);
            Assert.IsNotNull(actual1, "Unable to find Id == 1 so data was not written properly");
            Assert.AreEqual(expected1.FName, actual1.FirstName, "Expected first name to be equal!");
            Assert.AreEqual(expected1.LName, actual1.LastName, "Expected last name to be equal!");

            ColumnNameTestsExampe1 expected2 = originalData.FirstOrDefault(w => w.Id == 2);
            ColumnNameTestsExampe2 actual2 = actualList.FirstOrDefault(w => w.Id == 2);
            Assert.IsNotNull(actual2, "Unable to find Id == 2 so data was not written properly");
            Assert.AreEqual(expected2.FName, actual2.FirstName, "Expected first name to be equal!");
            Assert.AreEqual(expected2.LName, actual2.LastName, "Expected last name to be equal!");
        }
    }

    public class ColumnNameTestsExampe1
    {
        [ClassToExcel(Order = 1, ColumnName = "Id")]
        public int Id { get; set; }

        [ClassToExcel(Order = 2, ColumnName = "First Name")]
        public string FName { get; set; }

        [ClassToExcel(Order = 3, ColumnName = "Last Name")]
        public string LName { get; set; }
    }

    public class ColumnNameTestsExampe2
    {
        [ClassToExcel(Order = 1, ColumnName = "Id")]
        public int Id { get; set; }

        [ClassToExcel(Order = 3, ColumnName = "First Name")]
        public string FirstName { get; set; }

        [ClassToExcel(Order = 2, ColumnName = "Last Name")]
        public string LastName { get; set; }
    }


}
