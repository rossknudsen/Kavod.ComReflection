namespace Kavod.ComReflection.Types
{
    public sealed class SafeArray : Object
    {
        public static readonly SafeArray Instance = new SafeArray();

        private SafeArray() { }

        public override string ToString() => nameof(SafeArray);
    }
}