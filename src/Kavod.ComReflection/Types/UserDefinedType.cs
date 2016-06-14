using System.Collections.Generic;
using Kavod.ComReflection.Members;

namespace Kavod.ComReflection.Types
{
    public class UserDefinedType : Object
    {
        private readonly List<Field> _fields = new List<Field>();
        private readonly List<Method> _methods = new List<Method>();
        private readonly List<Object> _implementedTypes = new List<Object>();

        internal UserDefinedType(string name)
        {
            Name = name;
        }

        public string Name { get; }

        public IEnumerable<Field> Fields => _fields;

        public IEnumerable<Method> Methods => _methods;

        internal void AddImplementedType(Object implType)
        {
            _implementedTypes.Add(implType);
        }

        internal void AddImplementedTypes(IEnumerable<Object> implTypes)
        {
            _implementedTypes.AddRange(implTypes);
        }

        internal void AddField(Field field)
        {
            _fields.Add(field);
        }

        internal void AddFields(IEnumerable<Field> fields)
        {
            _fields.AddRange(fields);
        }

        public void AddMethods(IEnumerable<Method> methods)
        {
            _methods.AddRange(methods);
        }
    }
}
