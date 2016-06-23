namespace Kavod.ComReflection.Types
{
    public sealed class String : VbaType
    {
        public static readonly String Instance = new String();

        private String() : base(nameof(String))
        {
            IsPrimitive = true;
        }
    }
}