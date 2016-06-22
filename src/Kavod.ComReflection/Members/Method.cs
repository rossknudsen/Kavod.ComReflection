using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Text;
using Kavod.ComReflection.Types;

namespace Kavod.ComReflection.Members
{
    public abstract class Method
    {
        protected Method(string name, IEnumerable<Parameter> parameters, VbaType returnType)
        {
            Contract.Requires<ArgumentNullException>(!string.IsNullOrEmpty(name));
            Contract.Requires<ArgumentNullException>(parameters != null);
            Contract.Requires<ArgumentNullException>(returnType != null);

            Name = name;
            Parameters = parameters;
            ReturnType = returnType;
        }

        public string Name { get; }

        public IEnumerable<Parameter> Parameters { get; }

        public bool Hidden { get; internal set; }

        public VbaType ReturnType { get; }

        public abstract string ToSignatureString();

        protected string ConvertParametersForSignature()
        {
            var builder = new StringBuilder();
            foreach (var p in Parameters)
            {
                if (builder.Length > 0)
                {
                    builder.Append(", ");
                }
                builder.Append(p.ToSignatureString());
            }
            return builder.ToString();
        }
    }
}