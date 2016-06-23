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

            var method = value as MethodInfo;
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

        private static string ConvertParametersForSignature(MethodInfo methodInfo)
        {
            var builder = new StringBuilder();
            foreach (var p in methodInfo.Parameters)
            {
                if (builder.Length > 0)
                {
                    builder.Append(", ");
                }
                builder.Append(ConvertParameterForSignature(p));
            }
            return builder.ToString();
        }

        public static string ConvertParameterForSignature(ParameterInfo parameterInfo)
        {
            var result = "";
            if (parameterInfo.IsOptional)
            {
                result += "Optional ";
            }
            if (parameterInfo.IsOut)
            {
                result += "Out ";
            }
            result += $"{parameterInfo.ParamName} As {parameterInfo.ParamType}";
            if (parameterInfo.HasDefaultValue)
            {
                result += $" = {parameterInfo.DefaultValue}";
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
