using System.Collections.Generic;
using GalaSoft.MvvmLight;
using Kavod.ComReflection;
using Kavod.ComReflection.Types;

namespace ComBrowser.ViewModel
{
    public class LibraryOrTypeNodeViewModel : ViewModelBase
    {
        private static readonly Dictionary<VbaType, LibraryOrTypeNodeViewModel> _vmRegistry 
            = new Dictionary<VbaType, LibraryOrTypeNodeViewModel>();

        public LibraryOrTypeNodeViewModel(TypeLibrary library)
        {
            TypeLibrary = library;
            Name = library.Name;
            IconUriSource = @"\Resources\library.png";
            if (library.Hidden)
            {
                AccessUriSource = @"\Resources\lock.png";
            }

            foreach (var t in library.UserDefinedTypes)
            {
                GetViewModelForVbaType(t);
            }
        }

        private LibraryOrTypeNodeViewModel(VbaType type)
        {
            Name = type.Name;
            foreach (var f in type.Fields)
            {
                TypesOrMembers.Add(new MemberViewModel(f));
            }
            foreach (var m in type.Methods)
            {
                TypesOrMembers.Add(new MemberViewModel(m));
            }
            if (type.Hidden)
            {
                AccessUriSource = @"\Resources\lock.png";
            }
            foreach (var i in type.ImplementedInterfaces)
            {
                ChildNodes.Add(new LibraryOrTypeNodeViewModel(i));
            }
            SetIconUriSource(type);
        }

        private void GetViewModelForVbaType(VbaType type)
        {
            LibraryOrTypeNodeViewModel vm;
            if (!_vmRegistry.TryGetValue(type, out vm))
            {
                vm = new LibraryOrTypeNodeViewModel(type);
                _vmRegistry[type] = vm;
            }
            ChildNodes.Add(vm);
        }

        private void SetIconUriSource(VbaType type)
        {
            if (type.IsEnum)
            {
                IconUriSource = @"\Resources\enum.png";
            }
            else if (type.IsType)
            {
                IconUriSource = @"\Resources\type.png";
            }
            else if (type.IsModule)
            {
                IconUriSource = @"\Resources\module.png";
            }
            else if (type.IsInterface)
            {
                IconUriSource = @"\Resources\interface.png";
            }
            else if (type.IsDispatch)
            {
                IconUriSource = @"\Resources\udt.png";
            }
            else if (type.IsCoClass)
            {
                IconUriSource = @"\Resources\udt.png";
            }
        }

        public string Name { get; private set; }

        public IList<LibraryOrTypeNodeViewModel> ChildNodes { get; } = new List<LibraryOrTypeNodeViewModel>();

        public IList<MemberViewModel> TypesOrMembers { get; } = new List<MemberViewModel>();

        public TypeLibrary TypeLibrary { get; }

        public string IconUriSource { get; private set; }

        public string AccessUriSource { get; private set; }
    }
}
