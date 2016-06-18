namespace Kavod.ComReflection.Types
{
    public sealed class String : Object
    {
        public static String Instance = new String();

        private String() : base(nameof(String))
        {
            IsPrimitive = true;
        }

        public override string ToString() => nameof(String);
    }
}