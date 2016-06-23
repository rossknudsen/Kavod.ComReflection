namespace Kavod.ComReflection.Types
{
    public sealed class Long : VbaType
    {
        public static readonly Long Instance = new Long();

        private Long() : base(nameof(Long))
        {
            IsPrimitive = true;
        }
    }
}