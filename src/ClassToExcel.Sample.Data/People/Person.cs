using System;

namespace ClassToExcel.Sample.Data
{
    public class Person
    {
        [ClassToExcel(ColumnName = "First Name", Order = 2)]
        public string FirstName { get; set; }

        [ClassToExcel(ColumnName = "Last Name", Order = 1)]
        public string LastName { get; set; }

        [ClassToExcel(Order = 5)]
        public int Age { get; set; }

        [ClassToExcel(StyleFormat = "$ #,###0.00")]
        public decimal CheckingBalance { get; set; }

        public bool? CrazyAboutJellyBeans { get; set; }

        [ClassToExcel(ColumnName = "Date of Birth", StyleFormat = "yyyy/mm/dd")]
        public DateTime DOB { get; set; }

        [ClassToExcel(ColumnName = "IRA Balance", StyleFormat = "$ #,###0.00")]
        public decimal? IraBalance { get; set; }

        [ClassToExcel(ColumnName = "Date visted Australia", StyleFormat = "d/m/yyyy")]
        public DateTime? LastTimeVisitedAustralia { get; set; }

        public bool Married { get; set; }

        [ClassToExcel(StyleFormat = "#,##0")]
        public int NumberOfStepsTakenToday { get; set; }

        [ClassToExcel(StyleFormat = "#,##0")]
        public int? NumberOfTimesArrested { get; set; }
        
        public override string ToString()
        {
            return String.Format("{0} {1}  Age: {2}   DOB: {3}  Steps: {4} Australia: {5}  Arrested {6}  Married: {7}  Jelly Beans: {8} Checking: {9}  IRA: {10}",
                FirstName, LastName, Age, DOB, NumberOfStepsTakenToday, LastTimeVisitedAustralia, 
                NumberOfTimesArrested, Married, CrazyAboutJellyBeans, CheckingBalance, IraBalance);
        }
    }
}
