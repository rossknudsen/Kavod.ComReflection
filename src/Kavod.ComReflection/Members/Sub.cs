using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using Kavod.ComReflection.Types;

namespace Kavod.ComReflection.Members
{
    public class Sub : Method
    {
        public Sub(string name, IEnumerable<Parameter> parameters, VbaType returnType) 
            : base(name, parameters, returnType)
        {
            Contract.Requires<ArgumentNullException>(!string.IsNullOrEmpty(name));
            Contract.Requires<ArgumentNullException>(parameters != null);
            Contract.Requires<ArgumentNullException>(returnType != null);
        }

        public override string ToSignatureString() => $"{Name}({ConvertParametersForSignature()})";
    }
}