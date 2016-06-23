namespace Kavod.ComReflection.Types
{
    public sealed class Variant : VbaType
    {
        public static readonly Variant Instance = new Variant();

        private Variant() : base(nameof(Variant)) { }
    }
}