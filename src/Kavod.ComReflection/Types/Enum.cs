using System.Collections.Generic;

namespace Kavod.ComReflection.Types
{
    public class Enum : Object
    {
        private readonly List<EnumMember> _members = new List<EnumMember>();

        public Enum(string name) : base(name) { }

        public IEnumerable<EnumMember> Members => _members;

        public override string ToString() => Name;

        public void AddMember(EnumMember member)
        {
            _members.Add(member);
        }
    }
}