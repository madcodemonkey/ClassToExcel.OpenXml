using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ClassToExcel
{
    /// <summary>Converts a bunch of rows into a class.  The class must be decorated with the ClassToExcelRowAttribute</summary>
    public class ClassToExcelRowConverter<T> where T : class, new()
    {
        private readonly Action<ClassToExcelMessage> _loggingCallback;
        private readonly List<ClassToExcelRowConverterPropertyData> _properties;
        private readonly StringToPropertyConverter<T> _stringToPropertyConverter = new StringToPropertyConverter<T>();

        public ClassToExcelRowConverter(Action<ClassToExcelMessage> loggingCallback)
        {
            _loggingCallback = loggingCallback;
            _properties = CreateColumns();
        }

        /// <summary>Converts a bunch of rows into a class.  The class must be decorated with the ClassToExcelRowAttribute</summary>
        public T Convert(List<ClassToExcelRawRow> rows)
        {
            var result = new T();

            foreach (ClassToExcelRowConverterPropertyData item in _properties)
            {
                ClassToExcelRawRow row = rows.FirstOrDefault(w => w.RowNumber == item.RowNumber);
                if (row == null)
                {
                    LogMessage(ClassToExcelMessageType.Error, $"Row number is invalid ({item.RowNumber}) for the '{item.Property.Name}' property");
                    continue;
                }

                ClassToExcelRawColumn column = row.Columns.FirstOrDefault(w => w.ColumnLetter == item.ColumnLetter);
                if (column == null)
                {
                    LogMessage(ClassToExcelMessageType.Error, $"Column letter is invalid ({item.ColumnLetter}) for the '{item.Property.Name}' property");
                    continue;
                }

                var convertResult = _stringToPropertyConverter.AssignValue(item.Property, result, column.Data, item.DecimalPlaces);

                if (convertResult == StringToPropertyConverterEnum.Error || convertResult == StringToPropertyConverterEnum.Warning)
                {
                    var classToExcelMessageType = convertResult == StringToPropertyConverterEnum.Error ? ClassToExcelMessageType.Error : ClassToExcelMessageType.Warning;
                    LogMessage(new ClassToExcelMessage(classToExcelMessageType, item.RowNumber, null, item.ColumnLetter, _stringToPropertyConverter.LastMessage));
                }
            }

            return result;
        }
        
        /// <summary>Using reflection, create the columns and pull the attribute information off the class so
        /// that we know how to format the columns later.</summary>
        private List<ClassToExcelRowConverterPropertyData> CreateColumns()
        {
           var typeOfClass = typeof(T);

            List<ClassToExcelRowConverterPropertyData> columns = new List<ClassToExcelRowConverterPropertyData>();

            foreach (PropertyInfo property in typeOfClass.GetProperties(BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance))
            {
                if ((property.PropertyType != typeof(string) && property.PropertyType.IsClass) || property.PropertyType.IsArray)
                    continue;
                if (property.CanWrite == false)
                {
                    LogMessage(ClassToExcelMessageType.Info, $"Ignoring property because it is read-only: {property.Name}");
                    continue;
                }

                // If the property has a dislay name attribute, use it as the column name.
                object firstAttribute = property.GetCustomAttributes(typeof(ClassToExcelRowAttribute), false).FirstOrDefault();
                ClassToExcelRowAttribute displayAttribute = firstAttribute as ClassToExcelRowAttribute;
                if (displayAttribute != null)
                {
                    var newData = new ClassToExcelRowConverterPropertyData
                    {
                        Property = property,
                        RowNumber = displayAttribute.RowNumber,
                        ColumnLetter = displayAttribute.ColumnLetter,
                        DecimalPlaces = displayAttribute.DecimalPlaces
                    };
                    columns.Add(newData);
                }
            }

            return columns.OrderBy(o => o.RowNumber).ThenBy(o => o.ColumnLetter).ToList();
        }


        protected void LogMessage(ClassToExcelMessage message)
        {
            if (_loggingCallback != null)
            {
                _loggingCallback(message);
            }
        }

        protected void LogMessage(ClassToExcelMessageType messageType, string message)
        {
            if (_loggingCallback != null)
            {
                _loggingCallback(new ClassToExcelMessage(messageType, message));
            }
        }

    }
}
