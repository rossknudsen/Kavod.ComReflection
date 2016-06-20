namespace Kavod.ComReflection.Types
{
    internal class UnsupportedVariant : VbaType
    {
        public static readonly UnsupportedVariant Instance = new UnsupportedVariant();

        private UnsupportedVariant() : base(nameof(UnsupportedVariant)) { }

        public override string ToString() => $"<{nameof(UnsupportedVariant)}>";
    }
}