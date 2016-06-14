using System;
using System.Windows;
using System.Windows.Controls;
using ComBrowser.ViewModel;

namespace ComBrowser
{
    /// <summary>
    /// Interaction logic for LibraryTypeListControl.xaml
    /// </summary>
    public partial class LibraryTypeListControl : UserControl
    {
        public LibraryTypeListControl()
        {
            InitializeComponent();
        }

        public LibraryOrTypeNodeViewModel SelectedItem
        {
            get { return (LibraryOrTypeNodeViewModel) GetValue(SelectedItemProperty); }
            set { SetValue(SelectedItemProperty, value); }
        }
        
        public static readonly DependencyProperty SelectedItemProperty = DependencyProperty.Register(
            "SelectedItem",
            typeof(LibraryOrTypeNodeViewModel),
            typeof(LibraryTypeListControl));

        private void TreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            SelectedItem = (LibraryOrTypeNodeViewModel) e.NewValue;

            var handler = SelectedItemChanged;
            if (handler != null)
            {
                handler(sender, e);
            }
        }

        public EventHandler<RoutedPropertyChangedEventArgs<object>> SelectedItemChanged;
    }
}
