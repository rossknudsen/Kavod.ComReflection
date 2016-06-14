namespace Kavod.ComReflection.Types
{
    public sealed class Double : Object
    {
        public static Double Instance = new Double();

        private Double()
        {
            IsPrimitive = true;
        }

        public override string ToString() => nameof(Double);
    }
}