using SqlServerTool.ViewModels;
using System.Windows;

namespace SqlServerTool.Views
{
    public partial class ScriptDialog : Window
    {
        public ScriptDialog(ScriptViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
