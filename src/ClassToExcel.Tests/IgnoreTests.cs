using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClassToExcel.Tests
{
    /// <summary>
    /// Summary description for IgnoreTests
    /// </summary>
    [TestClass]
    public class IgnoreTests
    {
        [TestMethod]
        public void When_Ignore_is_used_with_the_ClassToExcelAttribute_data_is_not_written_to_the_file()
        {
            // Arrange
            // Arrange
            // Arrange
            var originalData = new List<IgnoreTestsExampe1>();
            originalData.Add(new IgnoreTestsExampe1 { Id = 1, Name = "Some Name"});
            originalData.Add(new IgnoreTestsExampe1 { Id = 2, Name = "Some Other Name" });

            // Act
            // Act
            // Act
            // Using a different class with similar properties WITHOUT the ignore attribute to prove that nothing was truly written
            // if we used the same class the reader would not even look for the Name field.
            var saveAndReadHelper = new SaveAndReadHelper<IgnoreTestsExampe1, IgnoreTestsExampe2>();
            List<IgnoreTestsExampe2> actualList = saveAndReadHelper.SaveAndRead(originalData, true);


            // Assert
            // Assert
            // Assert
            IgnoreTestsExampe2 data1 = actualList.FirstOrDefault(w => w.Id == 1);
            Assert.IsNotNull(data1, "Unable to find Id == 1 so data was not written properly");
            Assert.IsTrue(string.IsNullOrEmpty(data1.Name));
            IgnoreTestsExampe2 data2 = actualList.FirstOrDefault(w => w.Id == 2);
            Assert.IsNotNull(data2, "Unable to find Id == 2 so data was not written properly");
            Assert.IsTrue(string.IsNullOrEmpty(data2.Name));
        }

        [TestMethod]
        public void When_data_has_a_private_setter_data_is_written_to_the_file()
        {
            // If you want to ignore a property, users should use the Ignore property on the ClassToExcelAttribute
            // I'm testing this because the reader and writer use the same base class method for reflecting over properties.
            // Arrange
            // Arrange
            // Arrange
            var originalData = new List<IgnoreTestsExampe3>();
            originalData.Add(new IgnoreTestsExampe3 { Id = 1 });
            originalData.Add(new IgnoreTestsExampe3 { Id = 2 });

            // Act
            // Act
            // Act
            // Using a different class with similar properties WITHOUT the ignore attribute to prove that nothing was truly written
            // if we used the same class the reader would not even look for the Name field.
            var saveAndReadHelper = new SaveAndReadHelper<IgnoreTestsExampe3, IgnoreTestsExampe2>();
            List<IgnoreTestsExampe2> actualList = saveAndReadHelper.SaveAndRead(originalData, true);


            // Assert
            // Assert
            // Assert
            IgnoreTestsExampe2 data1 = actualList.FirstOrDefault(w => w.Id == 1);
            Assert.IsNotNull(data1, "Unable to find Id == 1 so data was not written properly");
            Assert.AreEqual("Some Name", data1.Name);
            IgnoreTestsExampe2 data2 = actualList.FirstOrDefault(w => w.Id == 2);
            Assert.IsNotNull(data2, "Unable to find Id == 2 so data was not written properly");
            Assert.AreEqual("Some Name", data2.Name);
        }

        [TestMethod]
        public void When_Ignore_is_used_with_the_ClassToExcelAttribute_data_is_not_read_from_the_file()
        {
            // Arrange
            // Arrange
            // Arrange
            var originalData = new List<IgnoreTestsExampe2>();
            originalData.Add(new IgnoreTestsExampe2 { Id = 1, Name = "Some Name" });
            originalData.Add(new IgnoreTestsExampe2 { Id = 2, Name = "Some Other Name" });

            // Act
            // Act
            // Act
            // Using a different class with similar properties WITHOUT the ignore attribute to prove that nothing was truly written
            // if we used the same class the reader would not even look for the Name field.
            var saveAndReadHelper = new SaveAndReadHelper<IgnoreTestsExampe2, IgnoreTestsExampe1>();
            List<IgnoreTestsExampe1> actualList = saveAndReadHelper.SaveAndRead(originalData, true);


            // Assert
            // Assert
            // Assert
            IgnoreTestsExampe1 data1 = actualList.FirstOrDefault(w => w.Id == 1);
            Assert.IsNotNull(data1, "Unable to find Id == 1 so data was not written properly");
            Assert.IsTrue(string.IsNullOrEmpty(data1.Name));
            IgnoreTestsExampe1 data2 = actualList.FirstOrDefault(w => w.Id == 2);
            Assert.IsNotNull(data2, "Unable to find Id == 2 so data was not written properly");
            Assert.IsTrue(string.IsNullOrEmpty(data2.Name));
        }

        [TestMethod]
        public void When_data_has_a_private_setter_data_is_not_read_from_the_file()
        {
            // If you want to ignore a property, users should use the Ignore property on the ClassToExcelAttribute
            // Arrange
            // Arrange
            // Arrange
            var originalData = new List<IgnoreTestsExampe2>();
            originalData.Add(new IgnoreTestsExampe2 { Id = 1, Name = "Random Name 1"});
            originalData.Add(new IgnoreTestsExampe2 { Id = 2, Name = "Random Name 2" });

            // Act
            // Act
            // Act
            // Using a different class with similar properties WITHOUT the ignore attribute to prove that nothing was truly written
            // if we used the same class the reader would not even look for the Name field.
            var saveAndReadHelper = new SaveAndReadHelper<IgnoreTestsExampe2, IgnoreTestsExampe3>();
            List<IgnoreTestsExampe3> actualList = saveAndReadHelper.SaveAndRead(originalData, true);


            // Assert
            // Assert
            // Assert
            IgnoreTestsExampe3 data1 = actualList.FirstOrDefault(w => w.Id == 1);
            Assert.IsNotNull(data1, "Unable to find Id == 1 so data was not written properly");
            Assert.AreEqual("Random Name 1", data1.Name);
            IgnoreTestsExampe3 data2 = actualList.FirstOrDefault(w => w.Id == 2);
            Assert.IsNotNull(data2, "Unable to find Id == 2 so data was not written properly");
            Assert.AreEqual("Random Name 2", data2.Name);
        }

        [TestMethod]
        public void Object_properties_are_ignored()
        {
            // Arrange
            // Arrange
            // Arrange
            var originalData = new List<IgnoreTestsExampe4>();
            originalData.Add(new IgnoreTestsExampe4 { Id = 1, Name = "Random Name 1", Owner = new IgnoreTestsExampe1() { Id = 4, Name = "Owner dude" }});
            originalData.Add(new IgnoreTestsExampe4 { Id = 2, Name = "Random Name 2", Owner = new IgnoreTestsExampe1() { Id = 5, Name = "Owner dude3"}});

            // Act
            // Act
            // Act
            // Using a different class with similar properties WITHOUT the ignore attribute to prove that nothing was truly written
            // if we used the same class the reader would not even look for the Name field.
            var saveAndReadHelper = new SaveAndReadHelper<IgnoreTestsExampe4, IgnoreTestsExampe4>();
            List<IgnoreTestsExampe4> actualList = saveAndReadHelper.SaveAndRead(originalData, true);


            // Assert
            // Assert
            // Assert
            IgnoreTestsExampe4 data1 = actualList.FirstOrDefault(w => w.Id == 1);
            Assert.IsNotNull(data1, "Unable to find Id == 1 so data was not written properly");
            Assert.AreEqual("Random Name 1", data1.Name);
            Assert.IsNull(data1.Owner);
            IgnoreTestsExampe4 data2 = actualList.FirstOrDefault(w => w.Id == 2);
            Assert.IsNotNull(data2, "Unable to find Id == 2 so data was not written properly");
            Assert.AreEqual("Random Name 2", data2.Name);
            Assert.IsNull(data2.Owner);
        }

        [TestMethod]
        public void Array_properties_are_ignored()
        {
            // Arrange
            // Arrange
            // Arrange
            var originalData = new List<IgnoreTestsExampe5>();
            originalData.Add(new IgnoreTestsExampe5 {Id = 1, Name = "Random Name 1", Numbers = new[] {1, 2, 3, 4}});
            originalData.Add(new IgnoreTestsExampe5 {Id = 2, Name = "Random Name 2", Numbers = new[] {5, 6, 6, 7}});

            // Act
            // Act
            // Act
            // Using a different class with similar properties WITHOUT the ignore attribute to prove that nothing was truly written
            // if we used the same class the reader would not even look for the Name field.
            var saveAndReadHelper = new SaveAndReadHelper<IgnoreTestsExampe5, IgnoreTestsExampe5>();
            List<IgnoreTestsExampe5> actualList = saveAndReadHelper.SaveAndRead(originalData, true);


            // Assert
            // Assert
            // Assert
            IgnoreTestsExampe5 data1 = actualList.FirstOrDefault(w => w.Id == 1);
            Assert.IsNotNull(data1, "Unable to find Id == 1 so data was not written properly");
            Assert.AreEqual("Random Name 1", data1.Name);
            Assert.IsNull(data1.Numbers);
            IgnoreTestsExampe5 data2 = actualList.FirstOrDefault(w => w.Id == 2);
            Assert.IsNotNull(data2, "Unable to find Id == 2 so data was not written properly");
            Assert.AreEqual("Random Name 2", data2.Name);
            Assert.IsNull(data2.Numbers);
        }
    }

    public class IgnoreTestsExampe1
    {
        public int Id { get; set; }

       [ClassToExcel(Ignore = true)]
        public string Name { get; set; }
    }

    public class IgnoreTestsExampe2
    {
        public int Id { get; set; }
        
        public string Name { get; set; }
    }

    public class IgnoreTestsExampe3
    {
        public IgnoreTestsExampe3()
        {
            Name = "Some Name";
        }

        public int Id { get; set; }

        public string Name { get; private set; }
    }

    public class IgnoreTestsExampe4
    {
        public int Id { get; set; }

        public string Name { get; set; }

        public IgnoreTestsExampe1 Owner { get; set; }
    }

    public class IgnoreTestsExampe5
    {
        public int Id { get; set; }

        public string Name { get; set; }

        public int[] Numbers { get; set; }
    }

}
