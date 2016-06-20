namespace Kavod.ComReflection.Types
{
    public sealed class Integer : VbaType
    {
        public static Integer Instance = new Integer();

        private Integer() : base(nameof(Integer))
        {
            IsPrimitive = true;
        }

        public override string ToString() => nameof(Integer);
    }
}