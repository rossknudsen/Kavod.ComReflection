using Kavod.ComReflection.Types;

namespace Kavod.ComReflection.Members
{
    public class TypeMember : Field
    {
        public TypeMember(string name, VbaType type, bool isConstant) 
            : base(name, type, isConstant)
        { }

        public override string ToSignatureString()
        {
            return $"{Name} As {Type}";
        }
    }
}