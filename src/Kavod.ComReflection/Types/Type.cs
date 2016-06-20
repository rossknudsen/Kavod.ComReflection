namespace Kavod.ComReflection.Types
{
    public class Type : VbaType
    {
        public Type(string name) : base(name) { }
        
        public override string ToString() => Name;
    }
}