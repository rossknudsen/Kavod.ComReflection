namespace Kavod.ComReflection.Types
{
    public sealed class Byte : VbaType
    {
        public static readonly Byte Instance = new Byte();

        private Byte() : base(nameof(Byte))
        {
            IsPrimitive = true;
        }
    }
}