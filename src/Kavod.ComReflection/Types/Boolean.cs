namespace Kavod.ComReflection.Types
{
    public sealed class Boolean : VbaType
    {
        public static readonly Boolean Instance = new Boolean();

        private Boolean() : base(nameof(Boolean))
        {
            IsPrimitive = true;
        }
    }
}