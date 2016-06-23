using System.Collections.Generic;

namespace Kavod.ComReflection.Types
{
    public sealed class Array : VbaType
    {
        // TODO it would be nice if this could be a generic class.
        private static readonly IDictionary<VbaType, Array> Instances = new Dictionary<VbaType, Array>();

        public static Array GetInstance(VbaType arrayType)
        {
            if (!Instances.ContainsKey(arrayType))
            {
                Instances.Add(arrayType, new Array(arrayType));
            }
            return Instances[arrayType];
        }

        private Array(VbaType arrayType) : base($"{nameof(Array)}(Of {arrayType.Name})")
        {
            ArrayType = arrayType;
        }

        public VbaType ArrayType { get; }
    }
}