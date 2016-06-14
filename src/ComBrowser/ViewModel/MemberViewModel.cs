using GalaSoft.MvvmLight;
using Kavod.ComReflection.Members;
using Kavod.ComReflection.Types;

namespace ComBrowser.ViewModel
{
    public class MemberViewModel : ViewModelBase
    {
        public MemberViewModel(Method method)
        {
            Name = method.ToSignatureString();
        }

        public MemberViewModel(Object type)
        {
            Name = type.Name;
        }

        public MemberViewModel(Field field)
        {
            Name = field.ToSignatureString();
        }

        public string Name { get; set; }
    }
}