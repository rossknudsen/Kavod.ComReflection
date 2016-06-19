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
            IconUriSource = @"\Resources\library.png";

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
            foreach (var i in library.Interfaces)
            {
                ChildNodes.Add(new LibraryOrTypeNodeViewModel(i));
                TypesOrMembers.Add(new MemberViewModel(i));
            }
            foreach (var i in library.DispatchInterfaces)
            {
                ChildNodes.Add(new LibraryOrTypeNodeViewModel(i));
                TypesOrMembers.Add(new MemberViewModel(i));
            }
        }

        private LibraryOrTypeNodeViewModel(UserDefinedType userDefinedType)
        {
            Name = userDefinedType.Name;
            IconUriSource = @"\Resources\udt.png";

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
            IconUriSource = @"\Resources\enum.png";
            foreach (var m in enumType.Members)
            {
                TypesOrMembers.Add(new MemberViewModel(m));
            }
        }

        private LibraryOrTypeNodeViewModel(Type type)
        {
            Name = type.Name;
            IconUriSource = @"\Resources\type.png";
            foreach (var m in type.TypeMembers)
            {
                TypesOrMembers.Add(new MemberViewModel(m));
            }
        }

        private LibraryOrTypeNodeViewModel(Module module)
        {
            Name = module.Name;
            IconUriSource = @"\Resources\module.png";
            foreach (var f in module.Fields)
            {
                TypesOrMembers.Add(new MemberViewModel(f));
            }
            foreach (var m in module.Methods)
            {
                TypesOrMembers.Add(new MemberViewModel(m));
            }
        }

        private LibraryOrTypeNodeViewModel(Interface @interface)
        {
            Name = @interface.Name;
            IconUriSource = @"\Resources\interface.png";
            foreach (var m in @interface.Methods)
            {
                TypesOrMembers.Add(new MemberViewModel(m));
            }
        }

        private LibraryOrTypeNodeViewModel(Dispatch dispatch)
        {
            Name = dispatch.Name;
        }

        public string Name { get; private set; }

        public IList<LibraryOrTypeNodeViewModel> ChildNodes { get; } = new List<LibraryOrTypeNodeViewModel>();

        public IList<MemberViewModel> TypesOrMembers { get; } = new List<MemberViewModel>();

        public TypeLibrary TypeLibrary { get; }

        public string IconUriSource { get; private set; }
    }
}
