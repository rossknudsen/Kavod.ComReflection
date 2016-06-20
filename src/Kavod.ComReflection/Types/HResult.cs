namespace Kavod.ComReflection.Types
{
    public sealed class HResult : VbaType
    {
        public static readonly HResult Instance = new HResult();

        private HResult() : base(nameof(HResult)) { }

        public override string ToString() => nameof(HResult);
    }
}