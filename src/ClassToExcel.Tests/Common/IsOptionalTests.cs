using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClassToExcel.Tests
{
    /// <summary>
    /// Summary description for IgnoreTests
    /// </summary>
    [TestClass]
    public class IsOptionalTests
    {
        private List<ClassToExcelMessage> _messages = new List<ClassToExcelMessage>();


        [TestMethod]
        public void We_get_header_a_column_missing_errors_when_we_do_not_use_ClassToExcelAttribute_IsOptional()
        {
            // Arrange
            // Arrange
            // Arrange
            _messages.Clear();
            var originalData = new List<OptionalTestsExampe1>();
            originalData.Add(new OptionalTestsExampe1 { Id = 1, LastName = "McMasters" });
            originalData.Add(new OptionalTestsExampe1 { Id = 2, LastName = "Jones" });

            // Act
            // Act
            // Act
            // Using a different class with similar properties WITHOUT the ignore attribute to prove that nothing was truly written
            // if we used the same class the reader would not even look for the Name field.
            var saveAndReadHelper = new SaveAndReadHelper<OptionalTestsExampe1, OptionalTestsExampe3>();
            List<OptionalTestsExampe3> actualList = saveAndReadHelper.SaveAndRead(originalData, true, LogMessage, LogMessage);


            // Assert
            // Assert
            // Assert
            Assert.IsTrue(
                _messages.Any(
                    w => w.MessageType == ClassToExcelMessageType.HeaderProblem && w.ColumnName == "First Name"),
                "We should get a header problem because the 'First Name' column is NOT marked as optional!");
        }

        [TestMethod]
        public void We_get_do_not_get_a_header_column_missing_errors_when_we_use_ClassToExcelAttribute_IsOptional()
        {
            // Arrange
            // Arrange
            // Arrange
            _messages.Clear();
            var originalData = new List<OptionalTestsExampe1>();
            originalData.Add(new OptionalTestsExampe1 { Id = 1, LastName = "McMasters"});
            originalData.Add(new OptionalTestsExampe1 { Id = 2, LastName = "Jones" });

            // Act
            // Act
            // Act
            // Using a different class with similar properties WITHOUT the ignore attribute to prove that nothing was truly written
            // if we used the same class the reader would not even look for the Name field.
            var saveAndReadHelper = new SaveAndReadHelper<OptionalTestsExampe1, OptionalTestsExampe2>();
            List<OptionalTestsExampe2> actualList = saveAndReadHelper.SaveAndRead(originalData, true, LogMessage, LogMessage);


            // Assert
            // Assert
            // Assert
            Assert.IsFalse(
                _messages.Any(
                    w => w.MessageType == ClassToExcelMessageType.HeaderProblem && w.ColumnName == "First Name"),
                "We should not have any header problems listed because the 'First Name' column is marked as optional!");
        }

        private void LogMessage(ClassToExcelMessage message)
        {
            if (message.MessageType != ClassToExcelMessageType.Info)
            {
                _messages.Add(message);
            }
            
        }
    }

    public class OptionalTestsExampe1
    {
        [ClassToExcel(Order = 1, ColumnName = "Id")]
        public int Id { get; set; }

        [ClassToExcel(Order = 2, ColumnName = "Last Name")]
        public string LastName { get; set; }
    }

    public class OptionalTestsExampe2
    {
        [ClassToExcel(Order = 1, ColumnName = "Id")]
        public int Id { get; set; }

        [ClassToExcel(Order = 3, ColumnName = "First Name", IsOptional = true)]
        public string FirstName { get; set; }

        [ClassToExcel(Order = 2, ColumnName = "Last Name")]
        public string LastName { get; set; }
    }

    public class OptionalTestsExampe3
    {
        [ClassToExcel(Order = 1, ColumnName = "Id")]
        public int Id { get; set; }

        [ClassToExcel(Order = 3, ColumnName = "First Name")]
        public string FirstName { get; set; }

        [ClassToExcel(Order = 2, ColumnName = "Last Name")]
        public string LastName { get; set; }
    }

}
