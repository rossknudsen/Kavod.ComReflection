namespace Kavod.ComReflection.Types
{
    /// <summary>
    /// Any is a special type used in Declare statements and designates a parameter
    /// which should not receive any type checking.
    /// </summary>
    public sealed class Any : VbaType
    {
        public static readonly Any Instance = new Any();

        private Any() : base(nameof(Any)) { }
    }
}