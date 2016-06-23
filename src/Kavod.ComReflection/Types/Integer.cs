namespace Kavod.ComReflection.Types
{
    public sealed class Integer : VbaType
    {
        public static readonly Integer Instance = new Integer();

        private Integer() : base(nameof(Integer))
        {
            IsPrimitive = true;
        }
    }
}