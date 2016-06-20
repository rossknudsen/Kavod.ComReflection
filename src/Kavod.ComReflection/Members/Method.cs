using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Text;

namespace Kavod.ComReflection.Members
{
    public abstract class Method
    {
        protected Method(string name, IEnumerable<Parameter> parameters)
        {
            Contract.Requires<ArgumentNullException>(!string.IsNullOrEmpty(name));
            Contract.Requires<ArgumentNullException>(parameters != null);

            Name = name;
            Parameters = parameters;
        }

        public string Name { get; }

        public IEnumerable<Parameter> Parameters { get; }

        public bool Hidden { get; internal set; }

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