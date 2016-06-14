using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;

namespace Kavod.ComReflection.Members
{
    public class Sub : Method
    {
        public Sub(string name, IEnumerable<Parameter> parameters) : base(name, parameters)
        {
            Contract.Requires<ArgumentNullException>(!string.IsNullOrEmpty(name));
            Contract.Requires<ArgumentNullException>(parameters != null);
        }

        public override string ToSignatureString() => $"{Name}({ConvertParametersForSignature()})";
    }
}