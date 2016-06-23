namespace Kavod.ComReflection.Types
{
    public sealed class Date : VbaType
    {
        public static readonly Date Instance = new Date();

        private Date() : base(nameof(Date))
        {
            IsPrimitive = true;
        }
    }
}