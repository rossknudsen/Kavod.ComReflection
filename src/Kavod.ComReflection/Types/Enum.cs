using System.Collections.Generic;

namespace Kavod.ComReflection.Types
{
    public class Enum : Object
    {
        private readonly List<EnumMember> _members;

        public Enum(string name, IEnumerable<EnumMember> members)
        {
            Name = name;
            _members = new List<EnumMember>(members);
        }

        public IEnumerable<EnumMember> Members => _members;

        public override string ToString() => Name;
    }
}