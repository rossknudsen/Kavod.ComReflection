using System;
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
using Object = Kavod.ComReflection.Types.Object;
using Single = Kavod.ComReflection.Types.Single;
using String = Kavod.ComReflection.Types.String;

namespace Kavod.ComReflection
{
    public class TypeLibrary
    {
        private readonly ComTypes.ITypeLib _typeLib;
        private readonly TypeLibraries _typeLibraries;
        private List<TypeInfoAndTypeAttr> _infoAndAttrs;
        private List<UserDefinedType> _udts;

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
            BuildTypeMembers();
        }

        private void CreateTypeInformation()
        {
            _infoAndAttrs = (from info in ComHelper.GetTypeInformation(_typeLib)
                             select new TypeInfoAndTypeAttr(info)).ToList();

            // build the basic list of types.
            _udts = (from t in _infoAndAttrs
                     select new UserDefinedType(t.Name)).ToList();
        }

        private void BuildTypeMembers()
        {
            // add the implemented type references to each type.
            foreach (var info in _infoAndAttrs)
            {
                var currentUdt = _udts.First(udt => udt.Name == info.Name);

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

        private UserDefinedType FindUserDefinedType(string typeName, ComTypes.ITypeInfo refTypeInfo)
        {
            Contract.Ensures(Contract.Result<UserDefinedType>() != null);

            var loadedType = VbaTypes
                .Concat(_typeLibraries.LoadedLibraries.SelectMany(l => l.VbaTypes))
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
                loadedType = library.VbaTypes.Single(t => t.Name == typeName);
            }
            return loadedType;
        }

        public string Name { get; }

        public string FilePath { get; }

        public IEnumerable<UserDefinedType> VbaTypes => _udts;
    }
}
