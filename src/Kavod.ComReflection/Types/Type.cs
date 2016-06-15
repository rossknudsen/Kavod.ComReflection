using System.Collections.Generic;
using Kavod.ComReflection.Members;

namespace Kavod.ComReflection.Types
{
    public class Type : Object
    {
        private readonly List<TypeMember> _members = new List<TypeMember>();

        public Type(string name)
        {
            Name = name;
        }

        public IEnumerable<TypeMember> TypeMembers => _members;

        public override string ToString() => Name;

        public void AddMember(TypeMember member)
        {
            _members.Add(member);
        }
    }

    public class TypeMember : Field
    {
        public TypeMember(string name, Object type, bool isConstant) 
            : base(name, type, isConstant)
        { }

        public override string ToSignatureString()
        {
            return $"{Name} As {Type}";
        }
    }
}