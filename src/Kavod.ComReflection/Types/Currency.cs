namespace Kavod.ComReflection.Types
{
    public sealed class Currency : VbaType
    {
        public static readonly Currency Instance = new Currency();

        private Currency() : base(nameof(Currency))
        {
            IsPrimitive = true;
        }
    }
}