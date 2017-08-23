using System;
using System.Collections.Generic;
using System.Linq;
using ClassToExcel.Sample.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClassToExcel.Tests
{
    public class PersonHelper
    {
        public static void CompareExpectedWithActualAndAssertIfNotEqual(List<Person> expectedList, List<Person> actualList)
        {
            foreach (Person expectedPerson in expectedList)
            {
                Person actualPerson = actualList.FirstOrDefault(w => w.FirstName == expectedPerson.FirstName && w.LastName == expectedPerson.LastName);
                if (actualPerson == null)
                    Assert.Fail("Could not find a person with the following name {0} {1}", expectedPerson.FirstName, expectedPerson.LastName);

                string name = String.Format("{0} {1}", expectedPerson.FirstName, expectedPerson.LastName);

                Assert.AreEqual(expectedPerson.Age, actualPerson.Age, "The expected Age {0} is not equal to the actual age {1} for person: {2}",
                    expectedPerson.Age, actualPerson.Age, name);
                Assert.AreEqual(expectedPerson.CheckingBalance, actualPerson.CheckingBalance, "The expected CheckingBalance {0} is not equal to the actual CheckingBalance {1} for person: {2}",
                    expectedPerson.CheckingBalance, actualPerson.CheckingBalance, name);
                Assert.AreEqual(expectedPerson.CrazyAboutJellyBeans, actualPerson.CrazyAboutJellyBeans, "The expected CrazyAboutJellyBeans {0} is not equal to the actual CrazyAboutJellyBeans {1} for person: {2}",
                    expectedPerson.CrazyAboutJellyBeans, actualPerson.CrazyAboutJellyBeans, name);
                Assert.AreEqual(expectedPerson.DOB, actualPerson.DOB, "The expected DOB {0} is not equal to the actual DOB {1} for person: {2}",
                    expectedPerson.DOB, actualPerson.DOB, name);
                Assert.AreEqual(expectedPerson.IraBalance, actualPerson.IraBalance, "The expected IraBalance {0} is not equal to the actual IraBalance {1} for person: {2}",
                    expectedPerson.IraBalance, actualPerson.IraBalance, name);
                Assert.AreEqual(expectedPerson.LastTimeVisitedAustralia, actualPerson.LastTimeVisitedAustralia, "The expected LastTimeVisitedAustralia {0} is not equal to the actual LastTimeVisitedAustralia {1} for person: {2}",
                    expectedPerson.LastTimeVisitedAustralia, actualPerson.LastTimeVisitedAustralia, name);
                Assert.AreEqual(expectedPerson.Married, actualPerson.Married, "The expected Married {0} is not equal to the actual Married {1} for person: {2}",
                    expectedPerson.Married, actualPerson.Married, name);
                Assert.AreEqual(expectedPerson.NumberOfStepsTakenToday, actualPerson.NumberOfStepsTakenToday, "The expected NumberOfStepsTakenToday {0} is not equal to the actual NumberOfStepsTakenToday {1} for person: {2}",
                    expectedPerson.NumberOfStepsTakenToday, actualPerson.NumberOfStepsTakenToday, name);
                Assert.AreEqual(expectedPerson.NumberOfTimesArrested, actualPerson.NumberOfTimesArrested, "The expected NumberOfTimesArrested {0} is not equal to the actual NumberOfTimesArrested {1} for person: {2}",
                    expectedPerson.NumberOfTimesArrested, actualPerson.NumberOfTimesArrested, name);
            }

        }
    }
}
