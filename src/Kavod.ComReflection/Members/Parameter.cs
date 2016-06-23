using Kavod.ComReflection.Types;

namespace Kavod.ComReflection.Members
{
    public class Parameter
    {
        public Parameter(string paramName, VbaType paramType)
        {
            ParamName = paramName;
            ParamType = paramType;
        }

        public string ParamName { get; internal set; }

        public VbaType ParamType { get; internal set; }

        public bool IsOptional { get; internal set; }

        public bool IsOut { get; internal set; }

        public object DefaultValue { get; internal set; }

        public bool HasDefaultValue { get; internal set; }
    }
}