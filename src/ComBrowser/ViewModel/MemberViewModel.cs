using System.Globalization;
using System.Windows.Media;
using GalaSoft.MvvmLight;
using Kavod.ComReflection.Members;
using Kavod.ComReflection.Types;

namespace ComBrowser.ViewModel
{
    public class MemberViewModel : ViewModelBase
    {
        private static MethodToStringConverter _converter = new MethodToStringConverter();
        private static Color DefaultFontColor = Colors.Black;
        private static Color HiddenFontColor = Colors.DarkGray;

        public MemberViewModel(Method method)
        {
            if (method.Hidden)
            {
                FontColor = HiddenFontColor;
            }
            Name = (string) _converter.Convert(method, typeof(string), null, CultureInfo.CurrentCulture);
        }

        public MemberViewModel(VbaType type)
        {
            if (type.Hidden)
            {
                FontColor = HiddenFontColor;
            }
            Name = type.Name;
        }

        public MemberViewModel(Field field)
        {
            Name = field.ToSignatureString();
        }

        public string Name { get; set; }

        public Color FontColor { get; private set; } = DefaultFontColor;
    }
}