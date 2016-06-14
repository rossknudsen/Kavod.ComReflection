using GalaSoft.MvvmLight;
using Kavod.ComReflection;

namespace ComBrowser.ViewModel
{
    public class LibraryRegistrationViewModel : ViewModelBase
    {
        public LibraryRegistrationViewModel(LibraryRegistration registration)
        {
            LibraryRegistration = registration;
        }

        public string FilePath => LibraryRegistration.FilePath;

        public string Name => LibraryRegistration.Name;

        public LibraryRegistration LibraryRegistration { get; }
    }
}