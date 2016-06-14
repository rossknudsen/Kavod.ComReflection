namespace Kavod.ComReflection.Types
{
    public sealed class Boolean : Object
    {
        public static Boolean Instance = new Boolean();

        private Boolean()
        {
            IsPrimitive = true;
        }

        public override string ToString() => nameof(Boolean);
    }
}