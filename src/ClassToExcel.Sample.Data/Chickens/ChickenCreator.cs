using System;
using System.Collections.Generic;

namespace ClassToExcel.Sample.Data
{
    public class ChickenCreator
    {
        public static List<Chicken> Work(int numberOfChickens)
        {
            var rand = new Random();
            var result = new List<Chicken>();

            for (int i = 1; i <= numberOfChickens; i++)
            {
                bool ageIsKnown = rand.Next(1, 100) > 50;

                var chicken = new Chicken
                {
                    Age = ageIsKnown ? rand.Next(1, 5) : (int?) null,
                    Name = "Bird" + i,
                    SizeOfPenInSquareFeet = rand.Next(1000, 3000),
                };

                DateTime? dob = ageIsKnown ? new DateTime(DateTime.Now.Year - chicken.Age.Value, rand.Next(1, 12), rand.Next(1, 29)) : (DateTime?) null;
                chicken.DOB = dob;

                decimal dollars = rand.Next(5, 50);
                decimal cents = (decimal) (rand.Next(1, 99) * .01);
                chicken.Value = dollars + cents;

                chicken.IsTall = rand.Next(1, 100) > 50;
                chicken.IsActive = rand.Next(1, 100) > 50 ? rand.Next(1, 100) > 50 : (bool?)null;

                result.Add(chicken);
            }

            return result;
        }
    }
}