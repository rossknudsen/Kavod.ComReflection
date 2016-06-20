namespace Kavod.ComReflection.Types
{
    public sealed class Double : VbaType
    {
        public static Double Instance = new Double();

        private Double() : base(nameof(Double))
        {
            IsPrimitive = true;
        }

        public override string ToString() => nameof(Double);
    }
}