namespace Kavod.ComReflection.Types
{
    public sealed class Unknown : Object
    {
        public static readonly Unknown Instance = new Unknown();

        private Unknown() { }

        public override string ToString() => nameof(Unknown);
    }
}