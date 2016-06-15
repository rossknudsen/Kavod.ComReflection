using System.Collections.Generic;
using GalaSoft.MvvmLight;
using Kavod.ComReflection;
using Kavod.ComReflection.Types;

namespace ComBrowser.ViewModel
{
    public class LibraryOrTypeNodeViewModel : ViewModelBase
    {
        public LibraryOrTypeNodeViewModel(TypeLibrary library)
        {
            TypeLibrary = library;
            Name = library.Name;

            foreach (var t in library.VbaTypes)
            {
                ChildNodes.Add(new LibraryOrTypeNodeViewModel(t));
                TypesOrMembers.Add(new MemberViewModel(t));
            }
            foreach (var e in library.Enums)
            {
                ChildNodes.Add(new LibraryOrTypeNodeViewModel(e));
                TypesOrMembers.Add(new MemberViewModel(e));
            }
            foreach (var t in library.Types)
            {
                ChildNodes.Add(new LibraryOrTypeNodeViewModel(t));
                TypesOrMembers.Add(new MemberViewModel(t));
            }
            foreach (var m in library.Modules)
            {
                ChildNodes.Add(new LibraryOrTypeNodeViewModel(m));
                TypesOrMembers.Add(new MemberViewModel(m));
            }
        }

        private LibraryOrTypeNodeViewModel(UserDefinedType userDefinedType)
        {
            Name = userDefinedType.Name;

            foreach (var f in userDefinedType.Fields)
            {
                TypesOrMembers.Add(new MemberViewModel(f));
            }
            foreach (var m in userDefinedType.Methods)
            {
                TypesOrMembers.Add(new MemberViewModel(m));
            }
        }

        private LibraryOrTypeNodeViewModel(Enum enumType)
        {
            Name = enumType.Name;
            foreach (var m in enumType.Members)
            {
                TypesOrMembers.Add(new MemberViewModel(m));
            }
        }

        private LibraryOrTypeNodeViewModel(Type type)
        {
            Name = type.Name;
            foreach (var m in type.TypeMembers)
            {
                TypesOrMembers.Add(new MemberViewModel(m));
            }
        }

        private LibraryOrTypeNodeViewModel(Module module)
        {
            Name = module.Name;
            foreach (var f in module.Fields)
            {
                TypesOrMembers.Add(new MemberViewModel(f));
            }
            foreach (var m in module.Methods)
            {
                TypesOrMembers.Add(new MemberViewModel(m));
            }
        }

        public string Name { get; private set; }

        public IList<LibraryOrTypeNodeViewModel> ChildNodes { get; } = new List<LibraryOrTypeNodeViewModel>();

        public IList<MemberViewModel> TypesOrMembers { get; } = new List<MemberViewModel>();

        public TypeLibrary TypeLibrary { get; }
    }
}
