using System.Windows;

namespace SqlServerTool.Views
{
    public partial class SqlPreviewDialog : Window
    {
        private readonly string _sql;

        public SqlPreviewDialog(string operationType, string objectName, string sql, bool viewOnly = false)
        {
            InitializeComponent();
            _sql = sql;
            Title          = viewOnly
                ? $"SQL ログ参照 - {objectName}"
                : $"SQL プレビュー - {operationType}";
            InfoLabel.Text = $"操作: {operationType}  /  対象: {objectName}";
            SqlEditor.Text = sql;

            if (viewOnly)
            {
                ExecuteButton.Visibility = Visibility.Collapsed;
                CancelButton.Content     = "閉じる";
            }
        }

        private void ExecuteButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void CopyButton_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < 3; i++)
            {
                try { System.Windows.Clipboard.SetDataObject(_sql, true); return; }
                catch { System.Threading.Thread.Sleep(100); }
            }
        }
    }
}
