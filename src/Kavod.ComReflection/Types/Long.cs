namespace Kavod.ComReflection.Types
{
    public sealed class Long : VbaType
    {
        public static Long Instance = new Long();

        private Long() : base(nameof(Long))
        {
            IsPrimitive = true;
        }

        public override string ToString() => nameof(Long);
    }
}