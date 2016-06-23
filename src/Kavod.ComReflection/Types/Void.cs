namespace Kavod.ComReflection.Types
{
    public class Void : VbaType
    {
        public static readonly Void Instance = new Void();

        public Void() : base(nameof(Void)) { }
    }
}
