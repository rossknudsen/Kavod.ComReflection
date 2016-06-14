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

        public string Name { get; private set; }

        public IList<LibraryOrTypeNodeViewModel> ChildNodes { get; } = new List<LibraryOrTypeNodeViewModel>();

        public IList<MemberViewModel> TypesOrMembers { get; } = new List<MemberViewModel>();

        public TypeLibrary TypeLibrary { get; }
    }
}
