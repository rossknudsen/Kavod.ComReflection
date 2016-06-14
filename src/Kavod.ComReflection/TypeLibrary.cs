using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Runtime.InteropServices;
using Kavod.ComReflection.Members;
using Kavod.ComReflection.Types;
using Array = Kavod.ComReflection.Types.Array;
using Boolean = Kavod.ComReflection.Types.Boolean;
using Byte = Kavod.ComReflection.Types.Byte;
using ComTypes = System.Runtime.InteropServices.ComTypes;
using Double = Kavod.ComReflection.Types.Double;
using Enum = Kavod.ComReflection.Types.Enum;
using Object = Kavod.ComReflection.Types.Object;
using Single = Kavod.ComReflection.Types.Single;
using String = Kavod.ComReflection.Types.String;
using Type = Kavod.ComReflection.Types.Type;

namespace Kavod.ComReflection
{
    public class TypeLibrary
    {
        private readonly ComTypes.ITypeLib _typeLib;
        private readonly TypeLibraries _typeLibraries;
        private List<TypeInfoAndTypeAttr> _infoAndAttrs;
        private List<UserDefinedType> _udts;
        private List<Enum> _enums;
        private List<Type> _types;

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

            var typeAttr = ComHelper.GetTypeLibAttr(typeLib);
            // TODO there is additional information in the TypeLibAttr class of interest.

            CreateTypeInformation();
            BuildUserDefinedTypeMembers();
        }

        private void CreateTypeInformation()
        {
            _infoAndAttrs = (from info in ComHelper.GetTypeInformation(_typeLib)
                             select new TypeInfoAndTypeAttr(info)).ToList();

            // build the basic list of types.
            _udts = new List<UserDefinedType>();
            _enums = new List<Enum>();
            _types = new List<Type>();
            foreach (var info in _infoAndAttrs)
            {
                switch (info.TypeAttr.typekind)
                {
                    case ComTypes.TYPEKIND.TKIND_ENUM:
                        var enumMembers = BuildEnumMembers(info.Name);
                        _enums.Add(new Enum(info.Name, enumMembers));
                        break;

                    case ComTypes.TYPEKIND.TKIND_RECORD:
                        var typeMembers = BuildTypeMembers(info.Name);
                        _types.Add(new Type(info.Name, typeMembers));
                        break;

                    case ComTypes.TYPEKIND.TKIND_MODULE:
                    case ComTypes.TYPEKIND.TKIND_INTERFACE:
                    case ComTypes.TYPEKIND.TKIND_DISPATCH:
                    case ComTypes.TYPEKIND.TKIND_COCLASS:
                    case ComTypes.TYPEKIND.TKIND_ALIAS:
                    case ComTypes.TYPEKIND.TKIND_UNION:
                    case ComTypes.TYPEKIND.TKIND_MAX:
                    default:
                        _udts.Add(new UserDefinedType(info.Name));
                        break;
                }
            }
        }

        private IEnumerable<TypeMember> BuildTypeMembers(string typeName)
        {
            var info = _infoAndAttrs.First(i => i.Name == typeName);
            foreach (var vardesc in ComHelper.GetTypeVariables(info))
            {
                var name = ComHelper.GetMemberName(info.TypeInfo, vardesc);
                var type = GetType(vardesc.elemdescVar.tdesc, info.TypeInfo);
                if (!ComHelper.IsConstant(vardesc))
                {
                    yield return new TypeMember(name, type, false);
                }
                else
                {
                    throw new NotImplementedException();
                }
            }
        }

        private IEnumerable<EnumMember> BuildEnumMembers(string enumName)
        {
            var info = _infoAndAttrs.First(i => i.Name == enumName);
            foreach (var vardesc in ComHelper.GetTypeVariables(info))
            {
                var name = ComHelper.GetMemberName(info.TypeInfo, vardesc);
                var type = GetType(vardesc.elemdescVar.tdesc, info.TypeInfo);
                if (ComHelper.IsConstant(vardesc))
                {
                    var constantValue = ComHelper.GetConstantValue(vardesc);
                    yield return new EnumMember(name, type, constantValue);
                }
                else
                {
                    throw new NotImplementedException();
                }
            }
        }

        private void BuildUserDefinedTypeMembers()
        {
            // add the implemented type references to each type.
            foreach (var currentUdt in _udts)
            {
                var info = _infoAndAttrs.First(i => i.Name == currentUdt.Name);

                // build the fields and constants for each type.
                foreach (var varDesc in ComHelper.GetTypeVariables(info))
                {
                    var name = ComHelper.GetMemberName(info.TypeInfo, varDesc);
                    var type = GetType(varDesc.elemdescVar.tdesc, info.TypeInfo);
                    if (ComHelper.IsConstant(varDesc))
                    {
                        var constantValue = ComHelper.GetConstantValue(varDesc);
                        currentUdt.AddField(new Constant(name, type, constantValue));
                    }
                    else
                    {
                        currentUdt.AddField(new Field(name, type, false));
                    }
                }

                // build the methods for each type.
                var methods = (from funcDesc in ComHelper.GetFuncDescs(info)
                    select BuildMethod(funcDesc, info.TypeInfo)).ToList();

                ConsolidateProperties(methods);
                currentUdt.AddMethods(methods);

                // add the implemented interfaces.
                foreach (var inherited in ComHelper.GetInheritedTypeInfos(info))
                {
                    var implementedName = ComHelper.GetTypeName(inherited);
                    Object implemented = _udts.FirstOrDefault(n => n.Name == implementedName);
                    if (implemented == null)
                    {
                        switch (implementedName)
                        {
                            case "IDispatch":
                                implemented = Object.GetInstance();
                                break;
                            case "IUnknown":
                                implemented = Unknown.Instance;
                                break;
                            default:
                                throw new Exception();
                        }
                    }
                    currentUdt.AddImplementedType(implemented);
                }
            }
        }

        private static void ConsolidateProperties(List<Method> methods)
        {
            var i = 0;
            while (i < methods.Count)
            {
                var currentMethod = methods[i] as Property;
                if (currentMethod != null)
                {
                    // check to see if there is another property of the same name.
                    var j = i + 1;
                    while (j < methods.Count)
                    {
                        var nextMethod = methods[j] as Property;
                        if (nextMethod != null
                            && nextMethod.Name == currentMethod.Name)
                        {
                            // Consolidate the two properties.
                            if (nextMethod.CanRead)
                            {
                                currentMethod.CanRead = true;
                            }
                            if (nextMethod.CanWrite)
                            {
                                currentMethod.CanWrite = true;
                            }
                            methods.Remove(nextMethod);
                        }
                        j++;
                    }
                }
                i++;
            }
        }

        private Method BuildMethod(ComTypes.FUNCDESC funcDesc, ComTypes.ITypeInfo typeInfo)
        {
            // collect parameters.
            var parameters = new List<Parameter>();
            var parameterNames = ComHelper.GetParameterNames(typeInfo, funcDesc).ToList();
            var elemDescs = ComHelper.GetElemDescs(funcDesc).ToList();
            for (var index = 0; index < parameterNames.Count; index++)
            {
                var elemDesc = elemDescs[index];
                var flags = elemDesc.desc.paramdesc.wParamFlags;
                var param = new Parameter(parameterNames[index], GetType(elemDesc.tdesc, typeInfo))
                {
                    IsOptional = flags.HasFlag(ComTypes.PARAMFLAG.PARAMFLAG_FOPT),
                    IsOut = flags.HasFlag(ComTypes.PARAMFLAG.PARAMFLAG_FOUT),
                    HasDefaultValue = flags.HasFlag(ComTypes.PARAMFLAG.PARAMFLAG_FHASDEFAULT)
                };
                if (param.HasDefaultValue)
                {
                    param.DefaultValue = ComHelper.GetDefaultValue(elemDesc.desc.paramdesc);
                }
                parameters.Add(param);
            }

            var name = ComHelper.GetMemberName(typeInfo, funcDesc);
            if (funcDesc.invkind.HasFlag(ComTypes.INVOKEKIND.INVOKE_PROPERTYGET))
            {
                var returnType = GetType(funcDesc.elemdescFunc.tdesc, typeInfo);
                var method = new Property(name, parameters, returnType, true, false);
                return method;
            }
            if (funcDesc.invkind.HasFlag(ComTypes.INVOKEKIND.INVOKE_FUNC)
                && (VarEnum)funcDesc.elemdescFunc.tdesc.vt != VarEnum.VT_VOID)
            {
                var returnType = GetType(funcDesc.elemdescFunc.tdesc, typeInfo);
                var method = new Function(name, parameters, returnType);
                return method;
            }
            if (funcDesc.invkind.HasFlag(ComTypes.INVOKEKIND.INVOKE_PROPERTYPUT)
                || funcDesc.invkind.HasFlag(ComTypes.INVOKEKIND.INVOKE_PROPERTYPUTREF))
            {
                var returnType = GetType(funcDesc.elemdescFunc.tdesc, typeInfo);
                var method = new Property(name, parameters, returnType, true, false);
                return method;
            }
            if (funcDesc.invkind.HasFlag(ComTypes.INVOKEKIND.INVOKE_FUNC)
                && (VarEnum)funcDesc.elemdescFunc.tdesc.vt == VarEnum.VT_VOID)
            {
                return new Sub(name, parameters);
            }
            throw new Exception();
        }

        private Object GetType(ComTypes.TYPEDESC tdesc, ComTypes.ITypeInfo context)
        {
            var vt = (VarEnum)tdesc.vt;
            ComTypes.TYPEDESC tdesc2;
            switch (vt)
            {
                case VarEnum.VT_PTR:
                    tdesc2 = Marshal.PtrToStructure<ComTypes.TYPEDESC>(tdesc.lpValue);
                    return GetType(tdesc2, context);

                case VarEnum.VT_USERDEFINED:
                    ComTypes.ITypeInfo refTypeInfo;
                    context.GetRefTypeInfo(tdesc.lpValue.ToInt32(), out refTypeInfo);
                    var typeName = ComHelper.GetTypeName(refTypeInfo);
                    var loadedType = FindUserDefinedType(typeName, refTypeInfo);
                    return loadedType;

                case VarEnum.VT_CARRAY:
                    tdesc2 = Marshal.PtrToStructure<ComTypes.TYPEDESC>(tdesc.lpValue);
                    dynamic arrayType = GetType(tdesc2, context);
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
                    return Object.GetInstance();

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

                case VarEnum.VT_UNKNOWN:
                    return Unknown.Instance;

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

        private Object FindUserDefinedType(string typeName, ComTypes.ITypeInfo refTypeInfo)
        {
            Contract.Ensures(Contract.Result<Object>() != null);

            var loadedType = AllTypes
                .Concat(_typeLibraries.LoadedLibraries.SelectMany(l => l.AllTypes))
                .FirstOrDefault(t => t.Name == typeName);
            if (loadedType == null)
            {
                ComTypes.ITypeLib referencedTypeLib;
                var index = -1;
                refTypeInfo.GetContainingTypeLib(out referencedTypeLib, out index);
                if (index == -1)
                {
                    throw new Exception("it wasn't me");
                }
                var library = _typeLibraries.LoadLibrary(referencedTypeLib);
                loadedType = library.AllTypes.Single(t => t.Name == typeName);
            }
            return loadedType;
        }

        public string Name { get; }

        public string FilePath { get; }

        public IEnumerable<UserDefinedType> VbaTypes => _udts;

        public IEnumerable<Enum> Enums => _enums;

        public IEnumerable<Type> Types => _types;

        public IEnumerable<Object> AllTypes => _udts
            .Concat<Object>(_enums)
            .Concat(_types);
    }
}
