using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Runtime.InteropServices;
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
        protected readonly List<EnumMember> _enumMembers = new List<EnumMember>();
        protected readonly List<TypeMember> _members = new List<TypeMember>();
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
                var name = ComHelper.GetMemberName(_info.TypeInfo, vardesc);
                var type = repo.GetVbaType(vardesc.elemdescVar.tdesc, _info.TypeInfo);
                var isConstant = ComHelper.IsConstant(vardesc);
                AddField(new Field(name, type, isConstant));
            }
        }

        private void BuildMethods(IVbaTypeRepository repo)
        {
            // build the methods for each type.
            var methods = (from funcDesc in ComHelper.GetFuncDescs(_info)
                           select BuildMethod(funcDesc, repo)).ToList();

            ConsolidateProperties(methods);

            _methods.AddRange(methods);
        }

        private Method BuildMethod(ComTypes.FUNCDESC funcDesc, IVbaTypeRepository repo)
        {
            // collect parameters.
            var parameters = new List<Parameter>();
            var parameterNames = ComHelper.GetParameterNames(_info.TypeInfo, funcDesc).ToList();
            var elemDescs = ComHelper.GetElemDescs(funcDesc).ToList();
            for (var index = 0; index < parameterNames.Count; index++)
            {
                var elemDesc = elemDescs[index];
                var flags = elemDesc.desc.paramdesc.wParamFlags;
                var param = new Parameter(parameterNames[index], repo.GetVbaType(elemDesc.tdesc, _info.TypeInfo))
                {
                    IsOptional = flags.HasFlag(ComTypes.PARAMFLAG.PARAMFLAG_FOPT),
                    IsOut = flags.HasFlag(ComTypes.PARAMFLAG.PARAMFLAG_FOUT),
                    HasDefaultValue = flags.HasFlag(ComTypes.PARAMFLAG.PARAMFLAG_FHASDEFAULT)
                };
                if (param.HasDefaultValue)
                {
                    param.DefaultValue = ComHelper.GetDefaultValue(elemDesc.desc.paramdesc);
                }
                parameters.Add(param);
            }

            var name = ComHelper.GetMemberName(_info.TypeInfo, funcDesc);
            var hidden = ((ComTypes.FUNCFLAGS)funcDesc.wFuncFlags).HasFlag(ComTypes.FUNCFLAGS.FUNCFLAG_FHIDDEN);
            // TODO there are some other FUNCFLAGS that may be of interest.
            if (funcDesc.invkind.HasFlag(ComTypes.INVOKEKIND.INVOKE_PROPERTYGET))
            {
                var returnType = repo.GetVbaType(funcDesc.elemdescFunc.tdesc, _info.TypeInfo);
                var method = new Property(name, parameters, returnType, true, false)
                {
                    Hidden = hidden
                };
                return method;
            }
            if (funcDesc.invkind.HasFlag(ComTypes.INVOKEKIND.INVOKE_FUNC)
                && (VarEnum)funcDesc.elemdescFunc.tdesc.vt != VarEnum.VT_VOID)
            {
                var returnType = repo.GetVbaType(funcDesc.elemdescFunc.tdesc, _info.TypeInfo);
                var method = new Function(name, parameters, returnType)
                {
                    Hidden = hidden
                };
                return method;
            }
            if (funcDesc.invkind.HasFlag(ComTypes.INVOKEKIND.INVOKE_PROPERTYPUT)
                || funcDesc.invkind.HasFlag(ComTypes.INVOKEKIND.INVOKE_PROPERTYPUTREF))
            {
                var returnType = repo.GetVbaType(funcDesc.elemdescFunc.tdesc, _info.TypeInfo);
                var method = new Property(name, parameters, returnType, true, false)
                {
                    Hidden = hidden
                };
                return method;
            }
            if (funcDesc.invkind.HasFlag(ComTypes.INVOKEKIND.INVOKE_FUNC)
                && (VarEnum)funcDesc.elemdescFunc.tdesc.vt == VarEnum.VT_VOID)
            {
                var method = new Sub(name, parameters)
                {
                    Hidden = hidden
                };
                return method;
            }
            throw new Exception();
        }

        private static void ConsolidateProperties(List<Method> methods)
        {
            var i = 0;
            while (i < methods.Count)
            {
                var currentMethod = methods[i] as Property;
                if (currentMethod != null)
                {
                    // check to see if there is another property of the same name.
                    var j = i + 1;
                    while (j < methods.Count)
                    {
                        var nextMethod = methods[j] as Property;
                        if (nextMethod != null
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

        public IEnumerable<EnumMember> EnumMembers => _enumMembers;

        public IEnumerable<TypeMember> TypeMembers => _members;

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

        public void AddEnumMember(EnumMember member)
        {
            _enumMembers.Add(member);
        }

        public void AddTypeMember(TypeMember member)
        {
            _members.Add(member);
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
