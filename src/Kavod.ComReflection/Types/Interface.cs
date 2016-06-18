using System.Collections.Generic;
using Kavod.ComReflection.Members;

namespace Kavod.ComReflection.Types
{
    public class Interface : Object
    {
        private readonly List<Method> _methods = new List<Method>();

        public Interface(string name) : base(name) { }

        public IEnumerable<Method> Methods => _methods;

        internal void AddMethod(Method method)
        {
            _methods.Add(method);
        }
    }
}
