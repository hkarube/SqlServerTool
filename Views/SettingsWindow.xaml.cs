using SqlServerTool.ViewModels;
using System.Windows;

namespace SqlServerTool.Views
{
    public partial class SettingsWindow : Window
    {
        private readonly SettingsViewModel _vm;

        public SettingsWindow(SettingsViewModel viewModel)
        {
            InitializeComponent();
            _vm = viewModel;
            DataContext = _vm;

            PasswordBox.Password = _vm.Password;

            _vm.RequestClose += () => Close();
        }

        private void PasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
        {
            _vm.Password = PasswordBox.Password;
        }
    }
}
