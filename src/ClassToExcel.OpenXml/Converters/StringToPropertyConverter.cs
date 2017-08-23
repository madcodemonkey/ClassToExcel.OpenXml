using System;
using System.Globalization;
using System.Reflection;

namespace ClassToExcel
{
    /// <summary>Converts a string based on the a property's type and assigns the converted value to the property.</summary>
    /// <remarks>I was originally going to write this as an extension and return an object from each call to AssignValue,
    /// but upon considering the thousands of calls that would be made to the AssignValue method I didn't think it 
    /// wise to create that many objects especially when most would be a "good" result and not used.</remarks>
    public class StringToPropertyConverter<T> where T: class 
    {
        /// <summary>If you got an error when calling AssignValue, you can read this property to figure out what went wrong.  
        /// The next call; however, will errase the value so read it immediately after calling AssignValue.</summary>
        public string LastMessage { get; set; }

        /// <summary>Use PropertyInfo to determine how stringValue should be converted.  Afterwards, it assigns the value to the object's property.</summary>
        /// <param name="property">Property information.  It will be used to determine the type and then assign the value to the object</param>
        /// <param name="obj">The object that will will receive the assignment.</param>
        /// <param name="stringValue">Value to assign to the property</param>
        public StringToPropertyConverterEnum AssignValue(PropertyInfo property, T obj, string stringValue)
        {
            if (property == null)
                throw new ArgumentException("You must specify a property information!");
            if (obj == null)
                throw new ArgumentException("You must specify an object to receive the new value!");

            // Reset LastMessage to erease anything that happened with the last call.
            LastMessage = string.Empty;

            if (property.PropertyType == typeof(string))
            {
                property.SetValue(obj, stringValue, null);
                return  StringToPropertyConverterEnum.Good;
            }

            if (string.IsNullOrEmpty(stringValue))
                return StringToPropertyConverterEnum.Default; // no issues.  Go with type default.

            if (property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(DateTime?))
                return AssignDateTimeValue(property, obj, stringValue);

            if (property.PropertyType == typeof(int) || property.PropertyType == typeof(int?))
              return AssignIntValue(property, obj, stringValue);

            if (property.PropertyType == typeof(double) || property.PropertyType == typeof(double?))
                return AssignDoubleValue(property, obj, stringValue);

            if (property.PropertyType == typeof(decimal) || property.PropertyType == typeof(decimal?))
                return AssignDecimalValue(property, obj, stringValue);

            if (property.PropertyType == typeof(bool) || property.PropertyType == typeof(bool?))
                return AssignBoolValue(property, obj, stringValue);

            LastMessage = $"Unknown property type '{property.PropertyType}' seen in AssignValue";
            return StringToPropertyConverterEnum.Error;

        }

        private StringToPropertyConverterEnum AssignBoolValue(PropertyInfo property, T obj, string stringValue)
        {
            if (string.IsNullOrWhiteSpace(stringValue))
                return StringToPropertyConverterEnum.Default;

            bool booleanValue;
            if (bool.TryParse(stringValue, out booleanValue))
            {
                property.SetValue(obj, booleanValue, null);
                return StringToPropertyConverterEnum.Good;
            }

            var lower = stringValue.Trim().ToLower();
            if (lower == "y" || lower == "true" || lower == "t" || lower == "1")
            {
                property.SetValue(obj, true, null);
                return StringToPropertyConverterEnum.Good;
            }

            property.SetValue(obj, false, null);

            if (lower == "n" || lower == "false" || lower == "f" || lower == "0")
            {
                return StringToPropertyConverterEnum.Good;
            }

            LastMessage = $"Assuming '{stringValue}' is false since parsing failed.";
            return StringToPropertyConverterEnum.Error; 
        }

        private StringToPropertyConverterEnum AssignDateTimeValue(PropertyInfo property, T obj, string stringValue)
        {
            if (string.IsNullOrWhiteSpace(stringValue))
                return StringToPropertyConverterEnum.Default;

            double value;
            if (double.TryParse(stringValue, out value))
            {
                DateTime fromOaDate = DateTime.FromOADate(value);
                property.SetValue(obj, fromOaDate, null);
                return StringToPropertyConverterEnum.Good;
            }

            LastMessage = $"Could not parse '{stringValue}' as an DateTime!";
            return StringToPropertyConverterEnum.Error;
        }

        private StringToPropertyConverterEnum AssignDecimalValue(PropertyInfo property, T obj, string stringValue)
        {
            if (string.IsNullOrWhiteSpace(stringValue))
                return StringToPropertyConverterEnum.Default;


            decimal number;
            // NumberStyles.Float used to allow parse to deal with exponents
            if (decimal.TryParse(stringValue, NumberStyles.Float, null, out number))
            {
                property.SetValue(obj, number, null);
                return StringToPropertyConverterEnum.Good;
            }

            string convertSpecialStrings = ConvertSpecialStrings(stringValue);
            // Avoid a recursive loop by comparing what was passed in to what the ConvertSpecialStrings
            // method returns. If they are the same, we need to exit and log an error!
            if (string.CompareOrdinal(convertSpecialStrings, stringValue) == 0)
            {
                LastMessage = $"Could not parse '{stringValue}' as a decimal!";
                return StringToPropertyConverterEnum.Error;
            }

            return AssignDecimalValue(property, obj, convertSpecialStrings);
        }

        private StringToPropertyConverterEnum AssignDoubleValue(PropertyInfo property, T obj, string stringValue)
        {
            if (string.IsNullOrWhiteSpace(stringValue))
                return StringToPropertyConverterEnum.Default;

            double number;
            // NumberStyles.Float used to allow parse to deal with exponents
            if (double.TryParse(stringValue, NumberStyles.Float, null, out number))
            {
                property.SetValue(obj, number, null);
                return StringToPropertyConverterEnum.Good;  
            }

            string convertSpecialStrings = ConvertSpecialStrings(stringValue);
            // Avoid a recursive loop by comparing what was passed in to what the ConvertSpecialStrings
            // method returns. If they are the same, we need to exit and log an error!
            if (string.CompareOrdinal(convertSpecialStrings, stringValue) == 0)
            {
                LastMessage = $"Could not parse '{stringValue}' as a double!";
                return StringToPropertyConverterEnum.Error;
            }

            return AssignDoubleValue(property, obj, convertSpecialStrings);
        }

        private StringToPropertyConverterEnum AssignIntValue(PropertyInfo property, T obj, string stringValue)
        {
            if (string.IsNullOrWhiteSpace(stringValue))
                return StringToPropertyConverterEnum.Default;

            int number;
            if (int.TryParse(stringValue, out number))
            {
                property.SetValue(obj, number, null);
                return StringToPropertyConverterEnum.Good;
            }

            LastMessage = $"Could not parse '{stringValue}' as an int!";
            return StringToPropertyConverterEnum.Error;
        }

        private string ConvertSpecialStrings(string stringValue)
        {
            if (stringValue.Contains("%"))
            {
                string stringWithoutPercentageSign = stringValue.Replace("%", "");
                double number;
                if (double.TryParse(stringWithoutPercentageSign, out number) == false)
                    return stringValue;

                number = number / 100.0;

                return number.ToString(CultureInfo.InvariantCulture);
            }

            return stringValue;
        }

    }
}
