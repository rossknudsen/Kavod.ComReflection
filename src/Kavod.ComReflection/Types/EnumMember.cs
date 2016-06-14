using Kavod.ComReflection.Members;

namespace Kavod.ComReflection.Types
{
    public class EnumMember : Constant
    {
        public EnumMember(string name, Object type, object value) : base(name, type, value)
        { }

        public override string ToSignatureString()
        {
            return $"{Name} = {Value}";
        }
    }
}