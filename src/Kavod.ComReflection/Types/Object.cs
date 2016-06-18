using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;

namespace Kavod.ComReflection.Types
{
    public class Object
    {
        private static readonly Object Instance = new Object(nameof(Object));

        public static Object GetInstance()
        {
            return Instance;
        } 

        private HashSet<string> _aliases = new HashSet<string>();

        protected Object(string name)
        {
            Name = name;
        }

        public string Name { get; protected set; }

        public virtual bool IsPrimitive { get; protected set; }

        public IEnumerable<string> Aliases => _aliases;

        public void AddAlias(string alias)
        {
            Contract.Requires<ArgumentNullException>(!string.IsNullOrEmpty(alias));

            _aliases.Add(alias);
        }

        public bool MatchNameOrAlias(string nameOrAlias)
        {
            return Name == nameOrAlias || _aliases.Contains(nameOrAlias);
        }

        public override string ToString() => nameof(Object);
    }
}
