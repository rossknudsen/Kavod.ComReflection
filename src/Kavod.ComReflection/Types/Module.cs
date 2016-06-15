using System.Collections.Generic;
using Kavod.ComReflection.Members;

namespace Kavod.ComReflection.Types
{
    public class Module : Object
    {
        private readonly List<Field> _fields = new List<Field>();
        private readonly List<Method> _methods = new List<Method>();

        internal Module(string name)
        {
            Name = name;
        }

        public IEnumerable<Field> Fields => _fields;

        public IEnumerable<Method> Methods => _methods;

        public override string ToString() => Name;

        internal void AddField(Field field)
        {
            _fields.Add(field);
        }

        internal void AddMethod(Method method)
        {
            _methods.Add(method);
        }
    }
}
