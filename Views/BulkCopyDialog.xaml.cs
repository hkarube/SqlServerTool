using SqlServerTool.ViewModels;
using System.Windows;

namespace SqlServerTool.Views
{
    public partial class BulkCopyDialog : Window
    {
        public BulkCopyDialog(BulkCopyViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }
    }
}
