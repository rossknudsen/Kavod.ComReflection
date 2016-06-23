using System.Globalization;
using System.Windows.Media;
using GalaSoft.MvvmLight;
using Kavod.ComReflection.Members;
using Kavod.ComReflection.Types;

namespace ComBrowser.ViewModel
{
    public class MemberViewModel : ViewModelBase
    {
        private static MethodToStringConverter _methodConverter = new MethodToStringConverter();
        private static FieldToStringConverter _fieldConverter = new FieldToStringConverter();
        private static Color DefaultFontColor = Colors.Black;
        private static Color HiddenFontColor = Colors.DarkGray;

        public MemberViewModel(Method method)
        {
            if (method.Hidden)
            {
                FontColor = HiddenFontColor;
            }
            Name = (string) _methodConverter.Convert(method, typeof(string), null, CultureInfo.CurrentCulture);
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
            Name = (string) _fieldConverter.Convert(field, typeof(string), null, CultureInfo.CurrentCulture);
        }

        public string Name { get; set; }

        public Color FontColor { get; private set; } = DefaultFontColor;
    }
}