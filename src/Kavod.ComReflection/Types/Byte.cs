namespace Kavod.ComReflection.Types
{
    public sealed class Byte : Object
    {
        public static Byte Instance = new Byte();

        private Byte()
        {
            IsPrimitive = true;
        }

        public override string ToString() => nameof(Byte);
    }
}