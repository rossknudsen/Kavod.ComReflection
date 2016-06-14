using System.Windows;
using ComBrowser.ViewModel;
using GalaSoft.MvvmLight.Messaging;

namespace ComBrowser.Views
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Views.ShowLibraryRegistrationsDialog _dialog;

        public MainWindow()
        {
            InitializeComponent();
            Closing += (s, e) => ViewModelLocator.Cleanup();
            Messenger.Default.Register<NotificationMessage>(this, OpenRegTypeLibWindow);
            Messenger.Default.Register<NotificationMessage>(this, CloseRegTypeLibWindow);
        }

        // This code could be moved to a UI level window coordinator of some kind.
        private void OpenRegTypeLibWindow(NotificationMessage message)
        {
            if (message.Notification != Messages.ShowRegisteredTypeLibrarySelectionDialog)
            {
                return;
            }
            _dialog = new Views.ShowLibraryRegistrationsDialog();
            var result = _dialog.ShowDialog();
        }

        private void CloseRegTypeLibWindow(NotificationMessage message)
        {
            if (message.Notification == Messages.CloseOpenRegTypeLibViewModel)
            {
                _dialog?.Close();
            }
        }
    }
}