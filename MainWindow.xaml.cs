using SqlServerTool.Models;
using SqlServerTool.Services;
using SqlServerTool.ViewModels;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace SqlServerTool
{
    public partial class MainWindow : Window
    {
        private MainViewModel Vm => (MainViewModel)DataContext;

        public MainWindow(DatabaseService dbService)
        {
            InitializeComponent();
            DataContext = new MainViewModel(dbService);
        }

        private void ObjectGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is DataGrid grid && grid.SelectedItem is ObjectInfo item)
                Vm.OpenDetailCommand.Execute(item);
        }

        private void ObjectGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is DataGrid grid)
                Vm.SelectedObjectInfos = grid.SelectedItems.Cast<ObjectInfo>().ToList();
        }

        private void ExportDefinition_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem mi
             && mi.Parent is ContextMenu cm
             && cm.PlacementTarget is DataGrid grid)
            {
                var selected = grid.SelectedItems.Cast<ObjectInfo>().ToList();
                Vm.ExportTableDefinitionCommand.Execute(selected);
            }
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
            => System.Windows.Application.Current.Shutdown();

        private void About_Click(object sender, RoutedEventArgs e)
            => System.Windows.MessageBox.Show("SQL Server Tool\nVersion 1.0", "バージョン情報",
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
    }
}
