namespace Kavod.ComReflection.Types
{
    public sealed class Any : Object
    {
        public static readonly Any Instance = new Any();

        private Any() : base(nameof(Any)) { }

        public override string ToString() => nameof(Any);
    }
}