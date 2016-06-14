namespace Kavod.ComReflection.Types
{
    public sealed class Variant : Object
    {
        public static Variant Instance = new Variant();

        private Variant() { }

        public override string ToString() => nameof(Variant);
    }
}