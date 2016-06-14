using GalaSoft.MvvmLight;
using Kavod.ComReflection;
using Kavod.ComReflection.Members;

namespace ComBrowser.ViewModel
{
    public class MemberViewModel : ViewModelBase
    {
        public MemberViewModel(Method method)
        {
            Name = method.ToSignatureString();
        }

        public MemberViewModel(UserDefinedType type)
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