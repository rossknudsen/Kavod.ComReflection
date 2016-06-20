using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using Kavod.ComReflection.Types;

namespace Kavod.ComReflection.Members
{
    public class Function : Method
    {
        public Function(string name, IEnumerable<Parameter> parameters, VbaType returnType) : base(name, parameters)
        {
            Contract.Requires<ArgumentNullException>(!string.IsNullOrEmpty(name));
            Contract.Requires<ArgumentNullException>(parameters != null);

            ReturnType = returnType;
        }

        public VbaType ReturnType { get; }

        public override string ToSignatureString() => $"{Name}({ConvertParametersForSignature()}) As {ReturnType}";
    }
}