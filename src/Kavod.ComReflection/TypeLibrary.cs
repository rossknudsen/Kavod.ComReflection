using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Runtime.InteropServices;
using Kavod.ComReflection.Types;
using Array = Kavod.ComReflection.Types.Array;
using Boolean = Kavod.ComReflection.Types.Boolean;
using Byte = Kavod.ComReflection.Types.Byte;
using ComTypes = System.Runtime.InteropServices.ComTypes;
using Double = Kavod.ComReflection.Types.Double;
using Object = Kavod.ComReflection.Types.Object;
using Single = Kavod.ComReflection.Types.Single;
using String = Kavod.ComReflection.Types.String;

namespace Kavod.ComReflection
{
    public class TypeLibrary : IVbaTypeRepository
    {
        private const string StdOleLibGuidString = "{00020430-0000-0000-C000-000000000046}";
        private static readonly Guid StdOleLibGuid = new Guid(StdOleLibGuidString);
        private static readonly VbaType[] _primitives = {
            Boolean.Instance,
            String.Instance,
            Long.Instance,
            Integer.Instance,
            UnsupportedVariant.Instance,
            Variant.Instance,
            Byte.Instance,
            Single.Instance,
            Double.Instance,
            Object.Instance,
            Currency.Instance,
            Date.Instance,
            Any.Instance,
            HResult.Instance,
            SafeArray.Instance
        };

        private readonly ComTypes.ITypeLib _typeLib;
        private readonly TypeLibraries _typeLibraries;
        private List<TypeInfoAndTypeAttr> _infoAndAttrs;
        private readonly List<VbaType> _userDefinedTypes = new List<VbaType>();

        internal TypeLibrary(string filePath, TypeLibraries typeLibraries) 
            : this(filePath, ComHelper.LoadTypeLibrary(filePath), typeLibraries)
        { }

        internal TypeLibrary(ComTypes.ITypeLib typeLib, TypeLibraries typeLibraries)
            : this("", typeLib, typeLibraries)
        { }

        private TypeLibrary(string filePath, ComTypes.ITypeLib typeLib, TypeLibraries typeLibraries)
        {
            FilePath = filePath;
            Name = ComHelper.GetTypeLibName(typeLib);
            _typeLib = typeLib;
            _typeLibraries = typeLibraries;

            var libAttr = ComHelper.GetTypeLibAttr(typeLib);
            Guid = libAttr.guid;
            Hidden = libAttr.wLibFlags.HasFlag(ComTypes.LIBFLAGS.LIBFLAG_FHIDDEN)
                     || libAttr.wLibFlags.HasFlag(ComTypes.LIBFLAGS.LIBFLAG_FRESTRICTED);
            Lcid = libAttr.lcid;
            MajorVersion = libAttr.wMajorVerNum;
            MinorVersion = libAttr.wMinorVerNum;

            CreateTypeInformation();
            BuildMembers();
        }

        private void CreateTypeInformation()
        {
            _infoAndAttrs = (from info in ComHelper.GetTypeInformation(_typeLib)
                             select new TypeInfoAndTypeAttr(info)).ToList();
            
            var aliases = new List<TypeInfoAndTypeAttr>();
            // build the basic list of types.
            foreach (var info in _infoAndAttrs)
            {
                switch (info.TypeAttr.typekind)
                {
                    case ComTypes.TYPEKIND.TKIND_ENUM:
                    case ComTypes.TYPEKIND.TKIND_RECORD:
                    case ComTypes.TYPEKIND.TKIND_MODULE:
                    case ComTypes.TYPEKIND.TKIND_INTERFACE:   // VTABLE only, not dual.
                    case ComTypes.TYPEKIND.TKIND_DISPATCH:    // true dispatch type or dual type.
                    case ComTypes.TYPEKIND.TKIND_COCLASS:
                        _userDefinedTypes.Add(new VbaType(info));
                        break;

                    case ComTypes.TYPEKIND.TKIND_ALIAS:
                        aliases.Add(info);
                        break;

                    case ComTypes.TYPEKIND.TKIND_UNION:
                    case ComTypes.TYPEKIND.TKIND_MAX:
                    default:
                        throw new NotImplementedException();
                }
            }
            foreach (var a in aliases)
            {
                AddAliasToType(a);  // possibly should create separate types.
            }
        }

        private void BuildMembers()
        {
            foreach (var type in UserDefinedTypes)
            {
                AddImplementedInterfaces(type);
                type.BuildMembers(this);
            }
        }

        private void AddAliasToType(TypeInfoAndTypeAttr info)
        {
            var type = GetVbaType(info.TypeAttr.tdescAlias, info.TypeInfo);
            type.AddAlias(info.Name);
        }

        private void AddImplementedInterfaces(VbaType type)
        {
            // add the implemented interfaces.
            var info = _infoAndAttrs.First(i => i.Name == type.Name);
            foreach (var inherited in ComHelper.GetInheritedTypeInfos(info))
            {
                var implemented = FindExistingType(inherited);
                if (implemented == null)
                {
                    LoadTypeLibrary(inherited);
                    implemented = FindExistingType(inherited);
                }
                type.AddImplementedType(implemented);
            }
        }

        public VbaType GetVbaType(ComTypes.TYPEDESC tdesc, ComTypes.ITypeInfo context)
        {
            var vt = (VarEnum)tdesc.vt;
            ComTypes.TYPEDESC tdesc2;
            switch (vt)
            {
                case VarEnum.VT_PTR:
                    tdesc2 = Marshal.PtrToStructure<ComTypes.TYPEDESC>(tdesc.lpValue);
                    return GetVbaType(tdesc2, context);

                case VarEnum.VT_USERDEFINED:
                    if (context == null)
                    {
                        throw new InvalidOperationException($"{nameof(context)} is null.  This is required for {VarEnum.VT_USERDEFINED}");
                    }
                    var refTypeInfo = ComHelper.GetRefTypeInfo(tdesc, context);
                    var loadedType = FindExistingType(refTypeInfo);
                    if (loadedType == null)
                    {
                        LoadTypeLibrary(refTypeInfo);
                        loadedType = FindExistingType(refTypeInfo);
                    }
                    return loadedType;

                case VarEnum.VT_UNKNOWN:
                    var stdOleLib = _typeLibraries.LoadLibrary(StdOleLibGuid);
                    return stdOleLib.UserDefinedTypes.First(t => t.Name == "IUnknown");

                case VarEnum.VT_CARRAY:
                    tdesc2 = Marshal.PtrToStructure<ComTypes.TYPEDESC>(tdesc.lpValue);
                    dynamic arrayType = GetVbaType(tdesc2, context);
                    return Array.GetInstance(arrayType);
                // lpValue is actually an ARRAYDESC structure containing dimension info.

                case VarEnum.VT_BOOL:
                    return Boolean.Instance;

                case VarEnum.VT_LPSTR:
                case VarEnum.VT_LPWSTR:
                case VarEnum.VT_BSTR:
                    return String.Instance;

                case VarEnum.VT_ERROR:
                case VarEnum.VT_INT:
                case VarEnum.VT_I4:
                    return Long.Instance;

                case VarEnum.VT_I2:
                    return Integer.Instance;

                case VarEnum.VT_UINT:
                case VarEnum.VT_UI4:
                case VarEnum.VT_UI2:
                    return UnsupportedVariant.Instance;

                case VarEnum.VT_UI8:
                case VarEnum.VT_I1:
                case VarEnum.VT_VARIANT:
                    return Variant.Instance;

                case VarEnum.VT_UI1:
                    return Byte.Instance;

                case VarEnum.VT_R4:
                    return Single.Instance;

                case VarEnum.VT_R8:
                    return Double.Instance;

                case VarEnum.VT_DISPATCH:
                case VarEnum.VT_BLOB:
                case VarEnum.VT_STREAM:
                case VarEnum.VT_STORAGE:
                case VarEnum.VT_STREAMED_OBJECT:
                case VarEnum.VT_STORED_OBJECT:
                case VarEnum.VT_BLOB_OBJECT:
                    return Object.Instance;

                case VarEnum.VT_CY:
                    return Currency.Instance;

                case VarEnum.VT_DATE:
                    return Date.Instance;

                case VarEnum.VT_VOID:
                    return Any.Instance;

                case VarEnum.VT_HRESULT:
                    return HResult.Instance;    // TODO this should be removed, errors are returned as exceptions

                case VarEnum.VT_SAFEARRAY:
                    return SafeArray.Instance;  // TODO should get more info for this type

                case VarEnum.VT_DECIMAL:
                // Currency?

                case VarEnum.VT_I8:
                // TODO this is a LongLong on 64 bit Office

                case VarEnum.VT_NULL:
                case VarEnum.VT_EMPTY:
                case VarEnum.VT_CLSID:
                case VarEnum.VT_RECORD:
                case VarEnum.VT_FILETIME:
                case VarEnum.VT_CF:
                case VarEnum.VT_VECTOR:
                case VarEnum.VT_ARRAY:
                case VarEnum.VT_BYREF:
                default:
                    throw new NotImplementedException();
            }
        }

        private VbaType FindExistingType(ComTypes.ITypeInfo info)
        {
            Contract.Requires<ArgumentNullException>(info != null);

            var typeName = ComHelper.GetTypeName(info);
            var query = PrimitiveTypes
                .Concat(UserDefinedTypes)
                .Concat(_typeLibraries.LoadedLibraries.SelectMany(l => l.UserDefinedTypes))
                .Where(t => t.MatchNameOrAlias(typeName));

            return query.FirstOrDefault();
        }

        private void LoadTypeLibrary(ComTypes.ITypeInfo info)
        {
            Contract.Requires<ArgumentNullException>(info != null);

            var lib = ComHelper.GetContainingTypeLib(info);
            _typeLibraries.LoadLibrary(lib);
        }

        public string Name { get; }

        public string FilePath { get; }

        public Guid Guid { get; }

        public bool Hidden { get; }

        public short MinorVersion { get; }

        public short MajorVersion { get; }

        public int Lcid { get; }

        public IEnumerable<VbaType> PrimitiveTypes => _primitives;

        public IEnumerable<VbaType> UserDefinedTypes => _userDefinedTypes;
    }
}
