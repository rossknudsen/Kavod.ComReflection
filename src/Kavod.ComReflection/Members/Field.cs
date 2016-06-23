using System.Runtime.InteropServices.ComTypes;
using Kavod.ComReflection.Types;

namespace Kavod.ComReflection.Members
{
    public class Field
    {
        public Field(VARDESC vardesc, ITypeInfo info, IVbaTypeRepository repo)
        {
            Name = ComHelper.GetMemberName(info, vardesc);
            Type = repo.GetVbaType(vardesc.elemdescVar.tdesc, info);
            IsConstant = ComHelper.IsConstant(vardesc);
            if (IsConstant)
            {
                Value = ComHelper.GetConstantValue(vardesc);
            }
            else
            {
                // TODO handle other cases here.  May need reference to parent type.
                IsField = true;
            }
        }

        public Field(string name, VbaType type, object value, bool isConstant)
        {
            Name = name;
            Type = type;
            IsConstant = isConstant;
            Value = value;
        }

        public VbaType Type { get; }

        public string Name { get; }

        public object Value { get; }

        public bool IsConstant { get; }

        public bool IsTypeMember { get; }

        public bool IsField { get; }

        public bool IsEnumMember { get; }
    }
}