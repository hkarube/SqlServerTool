using SqlServerTool.Services;
using SqlServerTool.ViewModels;
using SqlServerTool.Views;
using System.Windows;

namespace SqlServerTool
{
    public partial class App : System.Windows.Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            var dbService = new DatabaseService();
            var vm        = new ConnectionViewModel(dbService);
            var dialog    = new ConnectionDialog(vm);
            dialog.ShowDialog();

            if (!vm.IsConnected)
            {
                Shutdown();
                return;
            }

            ShutdownMode = ShutdownMode.OnMainWindowClose;
            var main = new MainWindow(dbService);
            MainWindow = main;
            main.Show();
        }
    }
}
