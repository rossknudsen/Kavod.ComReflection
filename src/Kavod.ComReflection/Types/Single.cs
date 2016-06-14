namespace Kavod.ComReflection.Types
{
    public sealed class Single : Object
    {
        public static Single Instance = new Single();

        private Single()
        {
            IsPrimitive = true;
        }

        public override string ToString() => nameof(Single);
    }
}