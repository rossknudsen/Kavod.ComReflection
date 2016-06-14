using System.Collections.ObjectModel;
using System.Linq;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using Kavod.ComReflection;
using System.Windows;
using System.Windows.Input;
using GalaSoft.MvvmLight.Messaging;

namespace ComBrowser.ViewModel
{
    public class MainViewModel : ViewModelBase
    {
        private readonly TypeLibraries _libraries;
        private string _welcomeTitle = "COM Type Library Viewer";

        public MainViewModel(TypeLibraries libraries)
        {
            _libraries = libraries;
            LoadedTypeLibraries = new ObservableCollection<LibraryOrTypeNodeViewModel>();

            PreLoadSomeLibraries();

            foreach (var library in _libraries.LoadedLibraries)
            {
                AddTypeLibraryViewModel(library);
            }

            Messenger.Default.Register<NotificationMessage<LibraryRegistrationViewModel>>(this, OpenLibrary);
        }

        private void PreLoadSomeLibraries()
        {
            var vbaLibraries = LibraryRegistration.GetRegisteredTypeLibraryEntries().Where(l => l.FilePath.Contains("VBA"));
            foreach (var l in vbaLibraries)
            {
                _libraries.LoadLibrary(l);
            }
        }

        public ObservableCollection<LibraryOrTypeNodeViewModel> LoadedTypeLibraries { get; }

        public string WelcomeTitle
        {
            get { return _welcomeTitle; }
            set { Set(ref _welcomeTitle, value); }
        }

        public ICommand QuitCommand { get; } = new RelayCommand(() => Application.Current.MainWindow.Close());

        public ICommand OpenRegTypeLibCommand { get; } = new RelayCommand(() =>
        {
            Messenger.Default.Send(new NotificationMessage(Messages.ShowRegisteredTypeLibrarySelectionDialog));
        });

        private void OpenLibrary(NotificationMessage<LibraryRegistrationViewModel> message)
        {
            if (message.Notification == Messages.LoadTypeLibrary)
            {
                var library = _libraries.LoadLibrary(message.Content.LibraryRegistration);
                AddTypeLibraryViewModel(library);
            }
        }

        private void AddTypeLibraryViewModel(TypeLibrary library)
        {
            if (library.FilePath == null)
            {
                return;
            }
            var viewModel = (from vm in LoadedTypeLibraries
                             where vm.TypeLibrary.FilePath != null
                                   && vm.TypeLibrary.FilePath == library.FilePath
                             select vm).SingleOrDefault();
            if (viewModel == null)
            {
                viewModel = new LibraryOrTypeNodeViewModel(library);
                LoadedTypeLibraries.Add(viewModel);
            }
        }
    }
}