using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Kavod.ComReflection.Types;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using Void = Kavod.ComReflection.Types.Void;

namespace Kavod.ComReflection.Members
{
    public class MethodInfo
    {
        private readonly List<ParameterInfo> _parameters = new List<ParameterInfo>();

        internal MethodInfo(FUNCDESC funcDesc, ITypeInfo info, IVbaTypeRepository repo)
        {
            Contract.Requires<ArgumentNullException>(info != null);
            Contract.Requires<ArgumentNullException>(repo != null);

            BuildMethod(funcDesc, info, repo);
            RewriteHResult();
        }

        public string Name { get; private set; }

        public IEnumerable<ParameterInfo> Parameters => _parameters;

        public bool Hidden { get; private set; }

        public bool CanRead { get; internal set; }

        public bool CanWrite { get; internal set; }

        public VbaType ReturnType { get; private set; }

        public bool IsProperty { get; private set; }

        public bool IsFunction { get; private set; }

        public bool IsSubroutine { get; private set; }

        private void BuildMethod(FUNCDESC funcDesc, ITypeInfo info, IVbaTypeRepository repo)
        {
            // collect parameters.
            var parameterNames = ComHelper.GetParameterNames(info, funcDesc).ToList();
            var elemDescs = ComHelper.GetElemDescs(funcDesc).ToList();
            for (var index = 0; index < parameterNames.Count; index++)
            {
                var elemDesc = elemDescs[index];
                var param = new ParameterInfo(parameterNames[index], elemDesc, info, repo);
                _parameters.Add(param);
            }

            Name = ComHelper.GetMemberName(info, funcDesc);
            ReturnType = repo.GetVbaType(funcDesc.elemdescFunc.tdesc, info);
            Hidden = ((System.Runtime.InteropServices.ComTypes.FUNCFLAGS)funcDesc.wFuncFlags).HasFlag(System.Runtime.InteropServices.ComTypes.FUNCFLAGS.FUNCFLAG_FHIDDEN);

            if (funcDesc.invkind.HasFlag(System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_PROPERTYGET))
            {
                IsProperty = true;
                CanRead = true;
            }
            else if (funcDesc.invkind.HasFlag(System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_FUNC)
                && (VarEnum)funcDesc.elemdescFunc.tdesc.vt != VarEnum.VT_VOID)
            {
                IsFunction = true;
            }
            else if (funcDesc.invkind.HasFlag(System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_PROPERTYPUT)
                || funcDesc.invkind.HasFlag(System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_PROPERTYPUTREF))
            {
                IsProperty = true;
                CanWrite = true;
            }
            else if (funcDesc.invkind.HasFlag(System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_FUNC)
                && (VarEnum)funcDesc.elemdescFunc.tdesc.vt == VarEnum.VT_VOID)
            {
                IsSubroutine = true;
            }
            else
            {
                throw new NotImplementedException();
            }
        }

        private void RewriteHResult()
        {
            if (ReturnType is HResult)
            {
                var outParams = Parameters.Where(p => p.IsOut).ToList();
                switch (outParams.Count)
                {
                    case 0:
                        IsFunction = false;
                        IsSubroutine = true;
                        IsProperty = false;
                        ReturnType = Void.Instance;
                        break;

                    case 1:
                        var outParam = outParams[0];
                        _parameters.Remove(outParam);
                        ReturnType = outParam.ParamType;
                        break;
                }
            }
        }
    }
}