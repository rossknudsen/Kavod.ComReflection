using Kavod.ComReflection.Types;

namespace Kavod.ComReflection.Members
{
    public class Constant : Field
    {
        public Constant(string name, VbaType type, object value) : base(name, type, true)
        {
            Value = value;
        }

        public object Value { get; }

        public override string ToSignatureString()
        {
            return $"Const {Name} As {Type} = {Value}";
        }
    }
}