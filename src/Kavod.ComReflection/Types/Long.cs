namespace Kavod.ComReflection.Types
{
    public sealed class Long : Object
    {
        public static Long Instance = new Long();

        private Long() : base(nameof(Long))
        {
            IsPrimitive = true;
        }

        public override string ToString() => nameof(Long);
    }
}