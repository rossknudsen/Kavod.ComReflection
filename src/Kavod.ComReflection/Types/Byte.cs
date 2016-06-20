namespace Kavod.ComReflection.Types
{
    public sealed class Byte : VbaType
    {
        public static Byte Instance = new Byte();

        private Byte() : base(nameof(Byte))
        {
            IsPrimitive = true;
        }

        public override string ToString() => nameof(Byte);
    }
}