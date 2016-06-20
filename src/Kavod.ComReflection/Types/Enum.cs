namespace Kavod.ComReflection.Types
{
    public class Enum : VbaType
    {
        public Enum(string name) : base(name) { }

        public override string ToString() => Name;
    }
}