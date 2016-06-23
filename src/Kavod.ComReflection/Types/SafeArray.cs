namespace Kavod.ComReflection.Types
{
    public sealed class SafeArray : VbaType
    {
        public static readonly SafeArray Instance = new SafeArray();

        private SafeArray() : base(nameof(SafeArray)) { }
    }
}