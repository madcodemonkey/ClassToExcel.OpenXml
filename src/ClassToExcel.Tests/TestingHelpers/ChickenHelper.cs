using System.Collections.Generic;
using System.Linq;
using ClassToExcel.Sample.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClassToExcel.Tests
{
    public class ChickenHelper
    {
        public static void CompareExpectedWithActualAndAssertIfNotEqual(List<Chicken> expectedList, List<Chicken> actualList)
        {
            foreach (Chicken expectedChicken in expectedList)
            {
                Chicken actualChicken = actualList.FirstOrDefault(w => w.Name == expectedChicken.Name);
                if (actualChicken == null)
                    Assert.Fail("Could not find a chicken with the following name {0}", expectedChicken.Name);

                string name = expectedChicken.Name;

                Assert.AreEqual(expectedChicken.Age, actualChicken.Age, "The expected Age {0} is not equal to the actual age {1} for chicken: {2}",
                    expectedChicken.Age, actualChicken.Age, name);
                Assert.AreEqual(expectedChicken.DOB, actualChicken.DOB, "The expected DOB {0} is not equal to the actual DOB {1} for chicken: {2}",
                    expectedChicken.DOB, actualChicken.DOB, name);
                Assert.AreEqual(expectedChicken.Value, actualChicken.Value, "The expected Value {0} is not equal to the actual Value {1} for chicken: {2}",
                    expectedChicken.Value, actualChicken.Value, name);
                Assert.AreEqual(expectedChicken.SizeOfPenInSquareFeet, actualChicken.SizeOfPenInSquareFeet, "The expected SizeOfPenInSquareFeet {0} is not equal to the actual SizeOfPenInSquareFeet {1} for chicken: {2}",
                    expectedChicken.SizeOfPenInSquareFeet, actualChicken.SizeOfPenInSquareFeet, name);
                Assert.AreEqual(expectedChicken.IsActive, actualChicken.IsActive, "The expected IsActive {0} is not equal to the actual IsActive {1} for chicken: {2}",
                    expectedChicken.IsActive, actualChicken.IsActive, name);
                Assert.AreEqual(expectedChicken.IsTall, actualChicken.IsTall, "The expected IsTall {0} is not equal to the actual IsTall {1} for chicken: {2}",
                    expectedChicken.IsTall, actualChicken.IsTall, name);
            }

        }
    }
}