namespace Kavod.ComReflection.Types
{
    public sealed class Currency : Object
    {
        public static Currency Instance = new Currency();

        private Currency()
        {
            IsPrimitive = true;
        }

        public override string ToString() => nameof(Currency);
    }
}