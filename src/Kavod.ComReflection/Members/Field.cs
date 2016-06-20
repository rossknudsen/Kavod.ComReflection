using Kavod.ComReflection.Types;

namespace Kavod.ComReflection.Members
{
    public class Field
    {
        public Field(string name, VbaType type, bool isConstant)
        {
            Name = name;
            Type = type;
            IsConstant = isConstant;
        }

        public VbaType Type { get; }

        public string Name { get; }

        public bool IsConstant { get; }

        public virtual string ToSignatureString()
        {
            return $"Const {Name} As {Type}";
        }
    }
}