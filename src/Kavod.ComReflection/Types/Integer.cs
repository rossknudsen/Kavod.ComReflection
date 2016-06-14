namespace Kavod.ComReflection.Types
{
    public sealed class Integer : Object
    {
        public static Integer Instance = new Integer();

        private Integer()
        {
            IsPrimitive = true;
        }

        public override string ToString() => nameof(Integer);
    }
}