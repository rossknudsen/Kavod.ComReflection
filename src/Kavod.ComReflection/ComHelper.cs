using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Kavod.ComReflection
{
    internal class ComHelper
    {
        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        static extern IntPtr LoadTypeLib(string fileName, out ComTypes.ITypeLib typeLib);

        internal static ComTypes.ITypeLib LoadTypeLibrary(string file)
        {
            ComTypes.ITypeLib typeLib = null;
            var ptr = IntPtr.Zero;
            if (!File.Exists(file))
            {
                throw new FileNotFoundException(file);
            }
            try
            {
                var hr = LoadTypeLib(file, out typeLib);
                if (hr != IntPtr.Zero)
                {
                    // A ComException may be thrown here anyway.
                    throw new Exception("COM Error...");
                }
                return typeLib;
            }
            finally
            {
                if (typeLib != null && ptr != IntPtr.Zero)
                {
                    typeLib.ReleaseTLibAttr(ptr);
                }
            }
        }

        internal static ComTypes.ITypeLib GetContainingTypeLib(ComTypes.ITypeInfo info)
        {
            ComTypes.ITypeLib referencedTypeLib;
            var index = -1;
            info.GetContainingTypeLib(out referencedTypeLib, out index);
            if (index == -1)
            {
                throw new Exception("it wasn't me");
            }
            return referencedTypeLib;
        }

        internal static string GetTypeLibName(ComTypes.ITypeLib refTypeInfo)
        {
            string refTypeName;
            string strHelpFile;
            int dwHelpContext;
            string strDocString;
            refTypeInfo.GetDocumentation(-1, out refTypeName, out strDocString, out dwHelpContext, out strHelpFile);
            return refTypeName;
        }

        internal static ComTypes.ITypeInfo GetRefTypeInfo(ComTypes.TYPEDESC tdesc, ComTypes.ITypeInfo context)
        {
            ComTypes.ITypeInfo refTypeInfo;
            context.GetRefTypeInfo(tdesc.lpValue.ToInt32(), out refTypeInfo);
            return refTypeInfo;
        }

        internal static string GetTypeName(ComTypes.ITypeInfo refTypeInfo)
        {
            string refTypeName;
            string strHelpFile;
            int dwHelpContext;
            string strDocString;
            refTypeInfo.GetDocumentation(-1, out refTypeName, out strDocString, out dwHelpContext, out strHelpFile);
            return refTypeName;
        }

        internal static ComTypes.TYPELIBATTR GetTypeLibAttr(ComTypes.ITypeLib typeLib)
        {
            IntPtr ptr;
            typeLib.GetLibAttr(out ptr);
            var typeLibAttr = Marshal.PtrToStructure<ComTypes.TYPELIBATTR>(ptr);
            typeLib.ReleaseTLibAttr(ptr);
            return typeLibAttr;
        }

        internal static IEnumerable<ComTypes.ITypeInfo> GetInheritedTypeInfos(TypeInfoAndTypeAttr infoAndAttr)
        {
            return GetInheritedTypeInfos(infoAndAttr.TypeInfo, infoAndAttr.TypeAttr);
        }

        internal static IEnumerable<ComTypes.ITypeInfo> GetInheritedTypeInfos(ComTypes.ITypeInfo typeInfo, ComTypes.TYPEATTR typeAttr)
        {
            for (var iImplType = 0; iImplType < typeAttr.cImplTypes; iImplType++)
            {
                int href;
                typeInfo.GetRefTypeOfImplType(iImplType, out href);
                ComTypes.ITypeInfo implTypeInfo;
                typeInfo.GetRefTypeInfo(href, out implTypeInfo);
                yield return implTypeInfo;
            }
        }

        internal static IEnumerable<ComTypes.ITypeInfo> GetTypeInformation(ComTypes.ITypeLib typeLib)
        {
            var comTypeCount = typeLib.GetTypeInfoCount();
            for (var typeIndex = 0; typeIndex < comTypeCount; typeIndex++)
            {
                ComTypes.ITypeInfo comTypeInfo;
                typeLib.GetTypeInfo(typeIndex, out comTypeInfo);
                yield return comTypeInfo;
            }
        }

        internal static string GetMemberName(ComTypes.ITypeInfo typeInfo, ComTypes.FUNCDESC funcDesc)
        {
            // first element in the sequence is the name of the method.
            return GetNames(typeInfo, funcDesc.memid).First();
        }

        internal static string GetMemberName(ComTypes.ITypeInfo typeInfo, ComTypes.VARDESC varDesc)
        {
            // first element in the sequence is the name of the method.
            return GetNames(typeInfo, varDesc.memid).First();
        }

        internal static IEnumerable<string> GetParameterNames(ComTypes.ITypeInfo typeInfo, ComTypes.FUNCDESC funcDesc)
        {
            // skip the first in the sequence because it is the name of the method.
            return GetNames(typeInfo, funcDesc.memid).Skip(1);
        }

        private static IEnumerable<string> GetNames(ComTypes.ITypeInfo typeInfo, int memberId)
        {
            const int maxNames = 255;
            var names = new string[maxNames];
            int pcNames;
            typeInfo.GetNames(memberId, names, maxNames, out pcNames);
            for (var i = 0; i < pcNames; i++)
            {
                yield return names[i];
            }
        }

        internal static IEnumerable<ComTypes.FUNCDESC> GetFuncDescs(TypeInfoAndTypeAttr infoAndAttr)
        {
            return GetFuncDescs(infoAndAttr.TypeInfo, infoAndAttr.TypeAttr);
        }

        private static IEnumerable<ComTypes.FUNCDESC> GetFuncDescs(ComTypes.ITypeInfo typeInfo, ComTypes.TYPEATTR typeAttr)
        {
            for (var iFunc = 0; iFunc < typeAttr.cFuncs; iFunc++)
            {
                IntPtr pFuncDesc;
                typeInfo.GetFuncDesc(iFunc, out pFuncDesc);
                var funcDesc = Marshal.PtrToStructure<ComTypes.FUNCDESC>(pFuncDesc);
                yield return funcDesc;
                typeInfo.ReleaseFuncDesc(pFuncDesc);
            }
        }

        internal static ComTypes.TYPEATTR GetTypeAttr(ComTypes.ITypeInfo typeInfo)
        {
            IntPtr pTypeAttr;
            typeInfo.GetTypeAttr(out pTypeAttr);
            var typeAttr = Marshal.PtrToStructure<ComTypes.TYPEATTR>(pTypeAttr);
            typeInfo.ReleaseTypeAttr(pTypeAttr);
            return typeAttr;
        }

        internal static IEnumerable<ComTypes.VARDESC> GetTypeVariables(TypeInfoAndTypeAttr info)
        {
            return GetTypeVariables(info.TypeInfo, info.TypeAttr);
        }

        internal static IEnumerable<ComTypes.VARDESC> GetTypeVariables(ComTypes.ITypeInfo typeInfo, ComTypes.TYPEATTR typeAttr)
        {
            for (var iVar = 0; iVar < typeAttr.cVars; iVar++)
            {
                IntPtr pVarDesc;
                typeInfo.GetVarDesc(iVar, out pVarDesc);
                var varDesc = Marshal.PtrToStructure<ComTypes.VARDESC>(pVarDesc);
                yield return varDesc;
                typeInfo.ReleaseVarDesc(pVarDesc);
            }
        }

        internal static IEnumerable<ComTypes.ELEMDESC> GetElemDescs(ComTypes.FUNCDESC funcDesc)
        {
            for (var iParam = 0; iParam < funcDesc.cParams; iParam++)
            {
                yield return Marshal.PtrToStructure<ComTypes.ELEMDESC>(
                    new IntPtr(funcDesc.lprgelemdescParam.ToInt64() +
                               Marshal.SizeOf(typeof(ComTypes.ELEMDESC)) * iParam));
            }
        }

        internal static bool IsConstant(ComTypes.VARDESC varDesc)
        {
            return varDesc.varkind == ComTypes.VARKIND.VAR_CONST;
        }

        internal static object GetConstantValue(ComTypes.VARDESC varDesc)
        {
            if (varDesc.varkind == ComTypes.VARKIND.VAR_CONST)
            {
                try
                {
                    var constantValue = varDesc.desc.lpvarValue;
                    return Marshal.GetObjectForNativeVariant(constantValue);
                }
                catch (Exception e)
                {
                    return null;
                }
            }
            return null;
        }

        internal static object GetDefaultValue(ComTypes.PARAMDESC paramDesc)
        {
            if (paramDesc.wParamFlags.HasFlag(ComTypes.PARAMFLAG.PARAMFLAG_FHASDEFAULT))
            {
                try
                {
                    // http://stackoverflow.com/a/20610486
                    var defaultValue = new IntPtr((long)paramDesc.lpVarValue + 8);
                    return Marshal.GetObjectForNativeVariant(defaultValue);
                }
                catch (Exception e)
                {
                    return null;
                }
            }
            return null;
        }
    }
}