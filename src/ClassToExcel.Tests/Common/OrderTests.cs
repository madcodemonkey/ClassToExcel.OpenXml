using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClassToExcel.Tests
{
    /// <summary>
    /// Summary description for IgnoreTests
    /// </summary>
    [TestClass]
    public class OrderTests
    {
        [TestMethod]
        public void When_no_header_is_present_the_ClassToExcelAttribute_Order_is_respected()
        {
            // Arrange
            // Arrange
            // Arrange
            var originalData = new List<OrderTestsExampe1>();
            originalData.Add(new OrderTestsExampe1 { Id = 1, FirstName = "Tom", LastName = "McMasters"});
            originalData.Add(new OrderTestsExampe1 { Id = 2, FirstName = "David", LastName = "Jones" });

            // Act
            // Act
            // Act
            // Using a different class with similar properties WITHOUT the ignore attribute to prove that nothing was truly written
            // if we used the same class the reader would not even look for the Name field.
            var saveAndReadHelper = new SaveAndReadHelper<OrderTestsExampe1, OrderTestsExampe2>();
            List<OrderTestsExampe2> actualList = saveAndReadHelper.SaveAndRead(originalData, false);


            // Assert
            // Assert
            // Assert
            OrderTestsExampe1 expected1 = originalData.FirstOrDefault(w => w.Id == 1);
            OrderTestsExampe2 actual1 = actualList.FirstOrDefault(w => w.Id == 1);
            Assert.IsNotNull(actual1, "Unable to find Id == 1 so data was not written properly");
            Assert.AreEqual(expected1.FirstName, actual1.LastName, "Expected first and last name to be inverted due to NO HEADER and order on each should be respected!");
            Assert.AreEqual(expected1.LastName, actual1.FirstName, "Expected first and last name to be inverted due to NO HEADER and order on each should be respected!");

            OrderTestsExampe1 expected2 = originalData.FirstOrDefault(w => w.Id == 2);
            OrderTestsExampe2 actual2 = actualList.FirstOrDefault(w => w.Id == 2);
            Assert.IsNotNull(actual2, "Unable to find Id == 2 so data was not written properly");
            Assert.AreEqual(expected2.FirstName, actual2.LastName, "Expected first and last name to be inverted due to NO HEADER and order on each should be respected!");
            Assert.AreEqual(expected2.LastName, actual2.FirstName, "Expected first and last name to be inverted due to NO HEADER and order on each should be respected!");
        }
        
        [TestMethod]
        public void When_a_header_is_present_the_ClassToExcelAttribute_ColumnName_takes_priority_over_Order()
        {
            // Arrange
            // Arrange
            // Arrange
            var originalData = new List<OrderTestsExampe1>();
            originalData.Add(new OrderTestsExampe1 { Id = 1, FirstName = "Tom", LastName = "McMasters" });
            originalData.Add(new OrderTestsExampe1 { Id = 2, FirstName = "David", LastName = "Jones" });

            // Act
            // Act
            // Act
            // Using a different class with similar properties WITHOUT the ignore attribute to prove that nothing was truly written
            // if we used the same class the reader would not even look for the Name field.
            var saveAndReadHelper = new SaveAndReadHelper<OrderTestsExampe1, OrderTestsExampe2>();
            List<OrderTestsExampe2> actualList = saveAndReadHelper.SaveAndRead(originalData, true);


            // Assert
            // Assert
            // Assert
            OrderTestsExampe1 expected1 = originalData.FirstOrDefault(w => w.Id == 1);
            OrderTestsExampe2 actual1 = actualList.FirstOrDefault(w => w.Id == 1);
            Assert.IsNotNull(actual1, "Unable to find Id == 1 so data was not written properly");
            Assert.AreEqual(expected1.FirstName, actual1.FirstName, "Expected first name to be equal!");
            Assert.AreEqual(expected1.LastName, actual1.LastName, "Expected last name to be equal!");

            OrderTestsExampe1 expected2 = originalData.FirstOrDefault(w => w.Id == 2);
            OrderTestsExampe2 actual2 = actualList.FirstOrDefault(w => w.Id == 2);
            Assert.IsNotNull(actual2, "Unable to find Id == 2 so data was not written properly");
            Assert.AreEqual(expected2.FirstName, actual2.FirstName, "Expected first name to be equal!");
            Assert.AreEqual(expected2.LastName, actual2.LastName, "Expected last name to be equal!");
        }
    }

    public class OrderTestsExampe1
    {
        [ClassToExcel(Order = 1, ColumnName = "Id")]
        public int Id { get; set; }

        [ClassToExcel(Order = 2, ColumnName = "First Name")]
        public string FirstName { get; set; }

        [ClassToExcel(Order = 3, ColumnName = "Last Name")]
        public string LastName { get; set; }
    }

    public class OrderTestsExampe2
    {
        [ClassToExcel(Order = 1, ColumnName = "Id")]
        public int Id { get; set; }

        [ClassToExcel(Order = 3, ColumnName = "First Name")]
        public string FirstName { get; set; }

        [ClassToExcel(Order = 2, ColumnName = "Last Name")]
        public string LastName { get; set; }
    }


}
