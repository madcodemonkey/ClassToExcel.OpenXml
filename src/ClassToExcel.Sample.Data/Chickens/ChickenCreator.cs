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
                var chicken = new Chicken
                {
                    Age = rand.Next(1, 5),
                    Name = "Bird" + i,
                    SizeOfPenInSquareFeet = rand.Next(1000, 3000),
                };

                DateTime dob = new DateTime(DateTime.Now.Year - chicken.Age, rand.Next(1, 12), rand.Next(1, 29));
                chicken.DOB = dob;

                decimal dollars = rand.Next(5, 50);
                decimal cents = (decimal) (rand.Next(1, 99) * .01);
                chicken.Value = dollars + cents;

                result.Add(chicken);
            }

            return result;
        }
    }
}