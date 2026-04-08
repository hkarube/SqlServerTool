using SqlServerTool.ViewModels;
using System.Windows;

namespace SqlServerTool.Views
{
    public partial class ConnectionDialog : Window
    {
        public ConnectionViewModel ViewModel { get; }

        public ConnectionDialog(ConnectionViewModel viewModel)
        {
            InitializeComponent();
            ViewModel = viewModel;
            DataContext = ViewModel;

            // 接続成功でダイアログを閉じる
            ViewModel.RequestClose += () =>
            {
                DialogResult = true;
                Close();
            };

            // 保存済みパスワードをPasswordBoxに反映
            PasswordBox.Password = ViewModel.Password;
        }

        private void PasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
        {
            ViewModel.Password = PasswordBox.Password;
        }

        private void SqlAuthRadio_Checked(object sender, RoutedEventArgs e)
            => ViewModel.UseWindowsAuth = false;

        private void WinAuthRadio_Checked(object sender, RoutedEventArgs e)
            => ViewModel.UseWindowsAuth = true;

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
