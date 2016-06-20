using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using Kavod.ComReflection.Members;

namespace Kavod.ComReflection.Types
{
    public class VbaType
    {
        protected readonly List<Field> _fields = new List<Field>();
        protected readonly List<Method> _methods = new List<Method>();
        protected readonly List<VbaType> _implementedTypes = new List<VbaType>();
        protected readonly List<EnumMember> _enumMembers = new List<EnumMember>();
        protected readonly List<TypeMember> _members = new List<TypeMember>();

        protected VbaType(string name)
        {
            Name = name;
        }

        public string Name { get; protected set; }

        public IEnumerable<Field> Fields => _fields;

        public IEnumerable<Method> Methods => _methods;

        public IEnumerable<EnumMember> EnumMembers => _enumMembers;

        public IEnumerable<TypeMember> TypeMembers => _members;

        public virtual bool IsPrimitive { get; protected set; }

        public bool Hidden { get; internal set; }

        internal void AddImplementedType(VbaType implType)
        {
            _implementedTypes.Add(implType);
        }

        internal void AddImplementedTypes(IEnumerable<VbaType> implTypes)
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
        
        internal void AddMethod(Method method)
        {
            _methods.Add(method);
        }

        internal void AddMethods(IEnumerable<Method> methods)
        {
            _methods.AddRange(methods);
        }

        public void AddEnumMember(EnumMember member)
        {
            _enumMembers.Add(member);
        }

        public void AddTypeMember(TypeMember member)
        {
            _members.Add(member);
        }

        public override string ToString() => Name;
    }
}
