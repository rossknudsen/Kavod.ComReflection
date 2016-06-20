namespace Kavod.ComReflection.Types
{
    public sealed class Variant : VbaType
    {
        public static Variant Instance = new Variant();

        private Variant() : base(nameof(Variant)) { }

        public override string ToString() => nameof(Variant);
    }
}