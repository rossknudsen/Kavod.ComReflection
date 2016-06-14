namespace Kavod.ComReflection.Types
{
    public sealed class HResult : Object
    {
        public static readonly HResult Instance = new HResult();

        private HResult() { }

        public override string ToString() => nameof(HResult);
    }
}