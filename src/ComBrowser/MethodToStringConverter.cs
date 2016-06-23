using System;
using System.Globalization;
using System.Text;
using System.Windows.Data;
using Kavod.ComReflection.Members;

namespace ComBrowser
{
    public class MethodToStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Convert Method to String
            if (targetType != typeof(string))
            {
                return null;
            }

            var method = value as Method;
            if (method == null)
            {
                return value?.ToString();
            }

            if (method.IsFunction)
            {
                return $"{method.Name}({ConvertParametersForSignature(method)}) As {method.ReturnType}";
            }
            if (method.IsSubroutine)
            {
                return $"{method.Name}({ConvertParametersForSignature(method)})";
            }
            if (method.IsProperty)
            {
                var propAccess = string.Empty;
                if (method.CanRead)
                {
                    propAccess += "Get";
                }
                if (method.CanWrite)
                {
                    if (propAccess.Length > 0)
                    {
                        propAccess += "/";
                    }
                    propAccess += method.ReturnType.IsPrimitive ? "Let" : "Set";
                }
                return $"Property {propAccess} {method.Name}({ConvertParametersForSignature(method)}) As {method.ReturnType}";
            }
            throw new NotImplementedException();
        }

        private static string ConvertParametersForSignature(Method method)
        {
            var builder = new StringBuilder();
            foreach (var p in method.Parameters)
            {
                if (builder.Length > 0)
                {
                    builder.Append(", ");
                }
                builder.Append(ConvertParameterForSignature(p));
            }
            return builder.ToString();
        }

        public static string ConvertParameterForSignature(Parameter parameter)
        {
            var result = "";
            if (parameter.IsOptional)
            {
                result += "Optional ";
            }
            if (parameter.IsOut)
            {
                result += "Out ";
            }
            result += $"{parameter.ParamName} As {parameter.ParamType}";
            if (parameter.HasDefaultValue)
            {
                result += $" = {parameter.DefaultValue}";
            }
            return result;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Convert from String back to Method.  Not implemented.
            return null;
        }
    }
}
