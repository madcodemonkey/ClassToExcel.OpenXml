using System;
using ClassToExcel;

namespace ClassToExcel.Sample.Data
{
    public class Chicken
    {
        public Chicken()
        {
            Owner = new Person { FirstName = "Farmer", LastName = "Brown", Age = 56};
        }

        [ClassToExcel(ColumnName = "Bird Name", Order = 1)]
        public string Name { get; set; }

        [ClassToExcel(Order = 2)]
        public int? Age { get; set; }

        [ClassToExcel(ColumnName = "Date of Birth", Order = 3)]
        public DateTime? DOB { get; set; }

        [ClassToExcel(StyleFormat = "$ #,###0.00", Order = 4)]
        public decimal Value { get; set; }

        [ClassToExcel(ColumnName = "Size of pen (sq. ft.)", StyleFormat = "#,###0", Order = 5)]
        public int SizeOfPenInSquareFeet { get; set; }
        
        /// <summary>Will be ignored.</summary>
        public Person Owner { get; set; }

        [ClassToExcel(Order = 7)]
        public bool? IsActive { get; set; }

        [ClassToExcel(Order = 8)]
        public bool IsTall { get; set; }

        public override string ToString()
        {
            return String.Format("{0} Age: {1}   DOB: {2}  Pen size (sq ft): {3}  Value: {4}  IsActive: {5}  IsTall: {6}", 
                Name, Age, DOB, SizeOfPenInSquareFeet, Value, IsActive, IsTall);
        }

        public bool IsEqual(Chicken chicken)
        {
            return (chicken.Name == Name && (chicken.SizeOfPenInSquareFeet == SizeOfPenInSquareFeet) && (chicken.Age == Age)
                 && chicken.DOB == DOB && chicken.Value == Value);
        }

    }
}
