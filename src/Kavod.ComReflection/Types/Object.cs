namespace Kavod.ComReflection.Types
{
    public class Object
    {
        private static readonly Object Instance = new Object();
        public static Object GetInstance()
        {
            return Instance;
        } 

        protected Object() { }

        public virtual bool IsPrimitive { get; protected set; }

        public override string ToString() => nameof(Object);
    }
}
