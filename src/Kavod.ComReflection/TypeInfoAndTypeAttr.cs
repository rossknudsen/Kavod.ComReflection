using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Kavod.ComReflection
{
    internal class TypeInfoAndTypeAttr
    {
        internal TypeInfoAndTypeAttr(ComTypes.ITypeInfo typeInfo)
        {
            TypeInfo = typeInfo;
            TypeAttr = ComHelper.GetTypeAttr(typeInfo);
            Name = ComHelper.GetTypeName(typeInfo);
        }

        internal ComTypes.TYPEATTR TypeAttr { get; }

        internal ComTypes.ITypeInfo TypeInfo { get; }

        internal string Name { get; }
    }
}