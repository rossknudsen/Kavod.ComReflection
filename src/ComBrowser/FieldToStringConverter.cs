using System;
using System.Globalization;
using System.Windows.Data;
using Kavod.ComReflection.Members;

namespace ComBrowser
{
    public class FieldToStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Convert Field to String
            if (targetType != typeof(string))
            {
                return null;
            }

            var field = value as Field;
            if (field == null)
            {
                return value?.ToString();
            }
            
            if (field.IsTypeMember)
            {
                return $"{field.Name} As {field.Type}";
            }
            if (field.IsConstant)
            {
                return $"Const {field.Name} As {field.Type} = {field.Value}";
            }
            if (field.IsEnumMember)
            {
                return $"{field.Name} = {field.Value}";
            }
            if (field.IsField)
            {
                return $"{field.Name} As {field.Type}";
            }
            throw new NotImplementedException();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Convert string back to Field.  Not Implemented.
            return null;
        }
    }
}