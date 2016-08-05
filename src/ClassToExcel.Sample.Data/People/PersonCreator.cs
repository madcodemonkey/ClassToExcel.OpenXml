using System;
using System.Collections.Generic;

namespace ClassToExcel.Sample.Data
{
    public class PersonCreator
    {
        public static List<Person> Work(int numberOfPeople)
        {
            var rand = new Random();
            var result = new List<Person>();

            for (int i = 1; i <= numberOfPeople; i++)
            {
                var person = new Person
                {
                    Age = rand.Next(12, 65),
                    FirstName = "First" + i,
                    LastName = "Last" + i,
                    NumberOfStepsTakenToday = rand.Next(1000, 30000),
                    Married = rand.Next(1, 100) > 50,
                    CheckingBalance = GetMoneyValue(2000, 5000, rand)
                };

                DateTime dob = new DateTime(DateTime.Now.Year - person.Age, rand.Next(1, 12), rand.Next(1, 29));
                person.DOB = dob;


                person.LastTimeVisitedAustralia = rand.Next(1, 100) > 50
                    ? (DateTime?) new DateTime(DateTime.Now.Year - rand.Next(1, 4), rand.Next(1, 12), rand.Next(1, 29))
                    : null;

                person.NumberOfTimesArrested = rand.Next(1, 100) > 50 ? rand.Next(1, 5) : (int?) null;

                person.CrazyAboutJellyBeans = rand.Next(1, 100) > 50 ? rand.Next(1, 100) > 50 : (bool?)null;

                person.IraBalance = rand.Next(1, 100) > 50 ? GetMoneyValue(100000, 300000, rand) : (decimal?)null;
                
                result.Add(person);
            }
            
            return result;
        }

        public static decimal GetMoneyValue(int start, int stop, Random rand)
        {
            decimal dollars = rand.Next(start, stop) * 1.00m;
            decimal cents = rand.Next(1, 100) * 0.01m;

            return dollars + cents;
        }
    }
}