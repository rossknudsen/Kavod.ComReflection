namespace Kavod.ComReflection.Types
{
    public class Void : VbaType
    {
        internal static Void Instance = new Void();

        public Void() : base(nameof(Void)) { }
    }
}
