using System;
using System.Diagnostics.Contracts;
using System.Runtime.InteropServices.ComTypes;
using Kavod.ComReflection.Types;

namespace Kavod.ComReflection.Members
{
    public class ParameterInfo
    {
        public ParameterInfo(string paramName, ELEMDESC elemDesc, ITypeInfo info, IVbaTypeRepository repo)
        {
            Contract.Requires<ArgumentNullException>(!string.IsNullOrEmpty(paramName));
            Contract.Requires<ArgumentNullException>(info != null);
            Contract.Requires<ArgumentNullException>(repo != null);

            ParamType = repo.GetVbaType(elemDesc.tdesc, info);
            ParamName = paramName;

            var flags = elemDesc.desc.paramdesc.wParamFlags;
            IsOptional = flags.HasFlag(PARAMFLAG.PARAMFLAG_FOPT);
            IsOut = flags.HasFlag(PARAMFLAG.PARAMFLAG_FOUT);
            HasDefaultValue = flags.HasFlag(PARAMFLAG.PARAMFLAG_FHASDEFAULT);

            if (HasDefaultValue)
            {
                DefaultValue = ComHelper.GetDefaultValue(elemDesc.desc.paramdesc);
            }
        }

        public string ParamName { get; }

        public VbaType ParamType { get; }

        public bool IsOptional { get; }

        public bool IsOut { get; }

        public object DefaultValue { get; }

        public bool HasDefaultValue { get; }
    }
}