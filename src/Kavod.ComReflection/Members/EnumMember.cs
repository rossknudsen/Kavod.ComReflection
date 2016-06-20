using Kavod.ComReflection.Types;

namespace Kavod.ComReflection.Members
{
    public class EnumMember : Constant
    {
        public EnumMember(string name, VbaType type, object value) : base(name, type, value)
        { }

        public override string ToSignatureString()
        {
            return $"{Name} = {Value}";
        }
    }
}