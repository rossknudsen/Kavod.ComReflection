namespace Kavod.ComReflection.Types
{
    public sealed class Single : VbaType
    {
        public static readonly Single Instance = new Single();

        private Single() : base(nameof(Single))
        {
            IsPrimitive = true;
        }
    }
}