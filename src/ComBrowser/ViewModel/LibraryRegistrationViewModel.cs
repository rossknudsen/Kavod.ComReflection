using System.IO;
using System.Windows.Media;
using GalaSoft.MvvmLight;
using Kavod.ComReflection;

namespace ComBrowser.ViewModel
{
    public class LibraryRegistrationViewModel : ViewModelBase
    {
        private static readonly Color DefaultFontColor = Colors.Black;
        private static readonly Color FileMissingFontColor = Colors.Red;

        public LibraryRegistrationViewModel(LibraryRegistration registration)
        {
            LibraryRegistration = registration;
            FontColor = File.Exists(registration.FilePath) ? DefaultFontColor : FileMissingFontColor;
        }

        public string FilePath => LibraryRegistration.FilePath;

        public string Name => LibraryRegistration.Name;

        public LibraryRegistration LibraryRegistration { get; }

        public Color FontColor { get; private set; }
    }
}