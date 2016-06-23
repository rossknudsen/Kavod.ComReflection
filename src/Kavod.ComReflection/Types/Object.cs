namespace Kavod.ComReflection.Types
{
    public class Object : VbaType
    {
        public static readonly Object Instance = new Object(nameof(Object));

        public Object(string name) : base(name) { }
    }
}