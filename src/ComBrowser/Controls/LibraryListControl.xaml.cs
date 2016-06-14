using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using ComBrowser.ViewModel;

namespace ComBrowser.Controls
{
    /// <summary>
    /// Interaction logic for LibraryListControl.xaml
    /// </summary>
    public partial class LibraryListControl : UserControl
    {
        public LibraryListControl()
        {
            InitializeComponent();
        }

        public ObservableCollection<LibraryRegistrationViewModel> RegisteredLibraries
        {
            get { return (ObservableCollection<LibraryRegistrationViewModel>) GetValue(RegisteredLibrariesProperty); }
            set { SetValue(RegisteredLibrariesProperty, value); }
        }

        public static readonly DependencyProperty RegisteredLibrariesProperty = DependencyProperty.Register(
            nameof(RegisteredLibraries),
            typeof(ObservableCollection<LibraryRegistrationViewModel>),
            typeof(LibraryListControl),
            new UIPropertyMetadata(OnRegisteredTypeLibrariesChanged));

        private static void OnRegisteredTypeLibrariesChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d == null)
            {
                return;
            }
            ((LibraryListControl) d).RegisteredLibraries = (ObservableCollection<LibraryRegistrationViewModel>) e.NewValue;
        }

        public LibraryRegistrationViewModel SelectedLibraryRegistration
        {
            get { return (LibraryRegistrationViewModel)GetValue(SelectedLibraryRegistrationProperty); }
            set { SetValue(SelectedLibraryRegistrationProperty, value); }
        }

        public static readonly DependencyProperty SelectedLibraryRegistrationProperty = DependencyProperty.Register(
            nameof(SelectedLibraryRegistration),
            typeof(LibraryRegistrationViewModel),
            typeof(LibraryListControl), 
            new PropertyMetadata(OnSelectedTypeLibraryChanged));

        private static void OnSelectedTypeLibraryChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d == null)
            {
                return;
            }
            ((LibraryListControl) d).SelectedLibraryRegistration = (LibraryRegistrationViewModel) e.NewValue;
        }
    }
}
