namespace Kavod.ComReflection.Types
{
    public class Type : VbaType
    {
        public Type(TypeInfoAndTypeAttr info) : base(info) { }
        
        public override string ToString() => Name;
    }
}