namespace Kavod.ComReflection.Types
{
    public sealed class Double : VbaType
    {
        public static readonly Double Instance = new Double();

        private Double() : base(nameof(Double))
        {
            IsPrimitive = true;
        }
    }
}