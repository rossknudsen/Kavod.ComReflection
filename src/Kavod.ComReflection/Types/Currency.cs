namespace Kavod.ComReflection.Types
{
    public sealed class Currency : VbaType
    {
        public static Currency Instance = new Currency();

        private Currency() : base(nameof(Currency))
        {
            IsPrimitive = true;
        }

        public override string ToString() => nameof(Currency);
    }
}