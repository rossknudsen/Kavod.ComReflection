using System.Collections.Generic;

namespace Kavod.ComReflection.Types
{
    public sealed class Array : Object
    {
        // TODO it would be nice if this could be a generic class.
        private static readonly IDictionary<Object, Array> Instances = new Dictionary<Object, Array>();

        public static Array GetInstance(Object arrayType)
        {
            if (!Instances.ContainsKey(arrayType))
            {
                Instances.Add(arrayType, new Array(arrayType));
            }
            return Instances[arrayType];
        }

        private Array(Object arrayType)
        {
            ArrayType = arrayType;
        }

        public Object ArrayType { get; }

        public override string ToString() => nameof(Array);
    }
}