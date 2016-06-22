using System.Runtime.InteropServices.ComTypes;
using Kavod.ComReflection.Types;

namespace Kavod.ComReflection
{
    public interface IVbaTypeRepository
    {
        VbaType GetVbaType(TYPEDESC tdesc, ITypeInfo context);
    }
}