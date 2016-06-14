using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using Object = Kavod.ComReflection.Types.Object;

namespace Kavod.ComReflection.Members
{
    public class Function : Method
    {
        public Function(string name, IEnumerable<Parameter> parameters, Object returnType) : base(name, parameters)
        {
            Contract.Requires<ArgumentNullException>(!string.IsNullOrEmpty(name));
            Contract.Requires<ArgumentNullException>(parameters != null);

            ReturnType = returnType;
        }

        public Object ReturnType { get; }

        public override string ToSignatureString() => $"{Name}({ConvertParametersForSignature()}) As {ReturnType}";
    }
}