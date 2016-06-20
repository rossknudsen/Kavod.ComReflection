namespace Kavod.ComReflection.Types
{
    public class Module : VbaType
    {
        internal Module(string name) : base(name) { }

        public override string ToString() => Name;
    }
}
