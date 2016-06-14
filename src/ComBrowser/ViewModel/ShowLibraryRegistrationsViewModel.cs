using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.CommandWpf;
using GalaSoft.MvvmLight.Messaging;
using Kavod.ComReflection;

namespace ComBrowser.ViewModel
{
    public class ShowLibraryRegistrationsViewModel : ViewModelBase
    {
        private LibraryRegistrationViewModel _selectedLibrary;

        public ShowLibraryRegistrationsViewModel()
        {
            var registeredLibraries = LibraryRegistration.GetRegisteredTypeLibraryEntries()
                .Select(registeredLibrary => new LibraryRegistrationViewModel(registeredLibrary))
                .OrderBy(l => l.Name);
            RegisteredLibraries = new ObservableCollection<LibraryRegistrationViewModel>(registeredLibraries);
        }

        public ObservableCollection<LibraryRegistrationViewModel> RegisteredLibraries { get; set; }

        public LibraryRegistrationViewModel SelectedLibraryRegistration
        {
            get { return _selectedLibrary; }
            set
            {
                if (_selectedLibrary == value)
                {
                    return;
                }
                _selectedLibrary = value;
                RaisePropertyChanged(() => SelectedLibraryRegistration);
            }
        }

        public ICommand AcceptCommand => new RelayCommand(Accepted);

        public ICommand CancelCommand => new RelayCommand(Cancelled);

        private void Accepted()
        {
            Messenger.Default.Send(
                new NotificationMessage<LibraryRegistrationViewModel>(SelectedLibraryRegistration, 
                Messages.LoadTypeLibrary)
            );
            Messenger.Default.Send(
                new NotificationMessage(Messages.CloseOpenRegTypeLibViewModel)
            );
        }

        private void Cancelled()
        {
            Messenger.Default.Send(
                new NotificationMessage(Messages.CloseOpenRegTypeLibViewModel)
            );
        }
    }
}
