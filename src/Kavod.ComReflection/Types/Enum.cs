namespace Kavod.ComReflection.Types
{
    public class Enum : VbaType
    {
        public Enum(TypeInfoAndTypeAttr info) : base(info) { }

        public override string ToString() => Name;
    }
}