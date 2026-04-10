using SqlServerTool.ViewModels;
using System.Windows;

namespace SqlServerTool.Views
{
    public partial class ExportDefinitionDialog : Window
    {
        public ExportDefinitionDialog(ExportDefinitionViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
            => Close();
    }
}
