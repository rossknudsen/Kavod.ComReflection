using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using Kavod.ComReflection.Types;

namespace Kavod.ComReflection.Members
{
    public class Property : Method
    {
        public Property(string name, IEnumerable<Parameter> parameters, VbaType returnType, bool canRead, bool canWrite) 
            : base(name, parameters)
        {
            Contract.Requires<ArgumentNullException>(returnType != null);
            Contract.Requires<ArgumentNullException>(!string.IsNullOrEmpty(name));
            Contract.Requires<ArgumentNullException>(parameters != null);
            Contract.Ensures(ReturnType != null);

            ReturnType = returnType;
            CanRead = canRead;
            CanWrite = canWrite;
        }

        public bool CanRead { get; internal set; }

        public bool CanWrite { get; internal set; }

        public VbaType ReturnType { get; }

        public override string ToSignatureString()
        {
            var propAccess = string.Empty;
            if (CanRead)
            {
                propAccess += "Get";
            }
            if (CanWrite)
            {
                if (propAccess.Length > 0)
                {
                    propAccess += "/";
                }
                propAccess += ReturnType.IsPrimitive ? "Let" : "Set";
            }
            return $"Property {propAccess} {Name}({ConvertParametersForSignature()}) As {ReturnType}";
        }
    }
}