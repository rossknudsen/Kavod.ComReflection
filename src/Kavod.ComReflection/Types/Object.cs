namespace Kavod.ComReflection.Types
{
    public class Object : VbaType
    {
        public static Object Instance = new Object(nameof(Object));

        public Object(string name) : base(name) { }
    }
}