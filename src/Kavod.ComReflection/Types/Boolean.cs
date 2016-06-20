namespace Kavod.ComReflection.Types
{
    public sealed class Boolean : VbaType
    {
        public static Boolean Instance = new Boolean();

        private Boolean() : base(nameof(Boolean))
        {
            IsPrimitive = true;
        }

        public override string ToString() => nameof(Boolean);
    }
}