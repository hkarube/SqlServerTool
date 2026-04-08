using SqlServerTool.ViewModels;
using System.Windows;

namespace SqlServerTool.Views
{
    public partial class BackupDialog : Window
    {
        public BackupDialog(BackupViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
            viewModel.RequestClose += () => Close();
        }
    }
}
