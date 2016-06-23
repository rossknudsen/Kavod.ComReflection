using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using Kavod.ComReflection.Members;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Kavod.ComReflection.Types
{
    public class VbaType
    {
        protected readonly TypeInfoAndTypeAttr _info;
        protected readonly List<Field> _fields = new List<Field>();
        protected readonly List<Method> _methods = new List<Method>();
        protected readonly List<VbaType> _implementedTypes = new List<VbaType>();
        protected HashSet<string> _aliases = new HashSet<string>();

        public VbaType(TypeInfoAndTypeAttr info) : this(info.Name)
        {
            _info = info;
            Hidden = _info.TypeAttr.wTypeFlags.HasFlag(ComTypes.TYPEFLAGS.TYPEFLAG_FHIDDEN);
            switch (_info.TypeAttr.typekind)
            {
                case ComTypes.TYPEKIND.TKIND_ENUM:
                    IsEnum = true;
                    break;

                case ComTypes.TYPEKIND.TKIND_RECORD:
                    IsType = true;
                    break;

                case ComTypes.TYPEKIND.TKIND_MODULE:
                    IsModule = true;
                    break;

                case ComTypes.TYPEKIND.TKIND_INTERFACE:   // VTABLE only, not dual.
                    IsInterface = true;
                    break;

                case ComTypes.TYPEKIND.TKIND_DISPATCH:    // true dispatch type or dual type.
                    IsDispatch = true;
                    break;

                case ComTypes.TYPEKIND.TKIND_COCLASS:
                    IsCoClass = true;
                    break;

                case ComTypes.TYPEKIND.TKIND_ALIAS:
                case ComTypes.TYPEKIND.TKIND_UNION:
                case ComTypes.TYPEKIND.TKIND_MAX:
                default:
                    throw new NotImplementedException();
            }
        }

        protected VbaType(string name)
        {
            Name = name;
        }

        internal void BuildMembers(IVbaTypeRepository repo)
        {
            BuildMethods(repo);
            BuildFields(repo);
        }

        private void BuildFields(IVbaTypeRepository repo)
        {
            foreach (var vardesc in ComHelper.GetTypeVariables(_info))
            {
                AddField(new Field(vardesc, _info.TypeInfo, repo));
            }
        }

        private void BuildMethods(IVbaTypeRepository repo)
        {
            var methods = (from funcDesc in ComHelper.GetFuncDescs(_info)
                           select new Method(funcDesc, _info.TypeInfo, repo)).ToList();

            ConsolidateProperties(methods);

            _methods.AddRange(methods);
        }

        private static void ConsolidateProperties(List<Method> methods)
        {
            var i = 0;
            while (i < methods.Count)
            {
                var currentMethod = methods[i];
                if (currentMethod.IsProperty)
                {
                    // check to see if there is another property of the same name.
                    var j = i + 1;
                    while (j < methods.Count)
                    {
                        var nextMethod = methods[j];
                        if (nextMethod.IsProperty
                            && nextMethod.Name == currentMethod.Name)
                        {
                            // Consolidate the two properties.
                            if (nextMethod.CanRead)
                            {
                                currentMethod.CanRead = true;
                            }
                            if (nextMethod.CanWrite)
                            {
                                currentMethod.CanWrite = true;
                            }
                            methods.Remove(nextMethod);
                        }
                        j++;
                    }
                }
                i++;
            }
        }

        public string Name { get; protected set; }

        public IEnumerable<Field> Fields => _fields;

        public IEnumerable<Method> Methods => _methods;

        public IEnumerable<string> Aliases => _aliases;

        public bool IsEnum { get; }

        public bool IsType { get; }

        public bool IsModule { get; }

        public bool IsInterface { get; }

        public bool IsDispatch { get; }

        public bool IsCoClass { get; }

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

        public void AddAlias(string alias)
        {
            Contract.Requires<ArgumentNullException>(!string.IsNullOrEmpty(alias));

            _aliases.Add(alias);
        }

        public bool MatchNameOrAlias(string nameOrAlias)
        {
            return Name == nameOrAlias || _aliases.Contains(nameOrAlias);
        }

        public override string ToString() => Name;
    }
}
