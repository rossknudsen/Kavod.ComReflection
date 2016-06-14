namespace Kavod.ComReflection.Types
{
    public sealed class Date : Object
    {
        public static Date Instance = new Date();

        private Date()
        {
            IsPrimitive = true;
        }

        public override string ToString() => nameof(Date);
    }
}