using SqlServerTool.Models;
using SqlServerTool.Services;
using System.Windows;
using System.Windows.Input;
using WpfMsg  = System.Windows.MessageBox;
using WpfMsgB = System.Windows.MessageBoxButton;
using WpfMsgI = System.Windows.MessageBoxImage;
using WpfMsgR = System.Windows.MessageBoxResult;

namespace SqlServerTool.Views
{
    public partial class SqlHistoryWindow : Window
    {
        private readonly DatabaseService? _dbService;

        public SqlHistoryWindow(DatabaseService? dbService = null)
        {
            InitializeComponent();
            _dbService = dbService;
            HistoryGrid.ItemsSource = SqlHistoryService.Instance.Entries;
        }

        private void HistoryGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
            => OpenSelected();

        private void ReapplyButton_Click(object sender, RoutedEventArgs e)
            => OpenSelected();

        private void OpenSelected()
        {
            if (HistoryGrid.SelectedItem is not SqlHistoryEntry entry) return;

            var preview = new SqlPreviewDialog(entry.OperationType, entry.ObjectName, entry.Sql)
            {
                Owner = this
            };
            if (preview.ShowDialog() != true) return;

            if (_dbService == null)
            {
                WpfMsg.Show("再適用にはDB接続が必要です。", "情報", WpfMsgB.OK, WpfMsgI.Information);
                return;
            }

            try
            {
                using var conn = _dbService.GetConnection();
                conn.Open();
                new Microsoft.Data.SqlClient.SqlCommand(entry.Sql, conn)
                    { CommandTimeout = 120 }
                    .ExecuteNonQuery();

                SqlHistoryService.Instance.Add(new SqlHistoryEntry
                {
                    ExecutedAt    = DateTime.Now,
                    OperationType = entry.OperationType + " (再適用)",
                    ObjectName    = entry.ObjectName,
                    Sql           = entry.Sql
                });
                WpfMsg.Show("再適用しました。", "完了", WpfMsgB.OK, WpfMsgI.Information);
            }
            catch (Exception ex)
            {
                WpfMsg.Show($"実行エラー: {ex.Message}", "エラー", WpfMsgB.OK, WpfMsgI.Error);
            }
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            var r = WpfMsg.Show("履歴をすべてクリアしますか？", "確認",
                WpfMsgB.YesNo, WpfMsgI.Question);
            if (r == WpfMsgR.Yes) SqlHistoryService.Instance.Clear();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
            => Close();
    }
}
