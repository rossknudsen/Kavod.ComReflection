namespace Kavod.ComReflection.Types
{
    public class Module : VbaType
    {
        internal Module(TypeInfoAndTypeAttr info) : base(info) { }

        public override string ToString() => Name;
    }
}
