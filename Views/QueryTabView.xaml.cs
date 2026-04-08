using SqlServerTool.ViewModels;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WpfUserControl = System.Windows.Controls.UserControl;
using WpfMouseEvent  = System.Windows.Input.MouseEventArgs;
using WpfDragEvent   = System.Windows.DragEventArgs;
using WpfPoint       = System.Windows.Point;
using WpfDragDrop    = System.Windows.DragDropEffects;
using WpfDataFormats  = System.Windows.DataFormats;
using WpfIDataObject  = System.Windows.IDataObject;

namespace SqlServerTool.Views
{
    public partial class QueryTabView : WpfUserControl
    {
        public QueryTabView()
        {
            InitializeComponent();

            // DataContext が変わったとき AvalonEdit にテキストをセット
            DataContextChanged += (s, e) =>
            {
                if (e.NewValue is QueryTab vm)
                {
                    SqlEditor.Text = vm.SqlText;
                    vm.PropertyChanged += (_, pe) =>
                    {
                        if (pe.PropertyName == nameof(QueryTab.SqlText)
                         && SqlEditor.Text != vm.SqlText)
                            SqlEditor.Text = vm.SqlText;
                    };
                }
            };

            // AvalonEdit の変更を ViewModel に反映
            SqlEditor.TextChanged += (s, e) =>
            {
                if (DataContext is QueryTab vm)
                    vm.SqlText = SqlEditor.Text;
            };

            // F5 は PreviewKeyDown で確実に捕捉（AvalonEdit 内でもきく）
            SqlEditor.PreviewKeyDown += (s, e) =>
            {
                if (e.Key == Key.F5 && DataContext is QueryTab vm)
                {
                    vm.ExecuteSqlCommand.Execute(null);
                    e.Handled = true;
                }
            };

            // ツリーアイテムの D&D 開始
            ObjectTree.PreviewMouseMove += ObjectTree_PreviewMouseMove;

            // エディタへのドロップ（文字列＋ファイル両対応）
            SqlEditor.Drop     += SqlEditor_Drop;
            SqlEditor.DragOver += SqlEditor_DragOver;

            // UserControl 全体へのファイルドロップ（エディタ外に落とした場合）
            Drop     += UserControl_Drop;
            DragOver += UserControl_DragOver;
        }

        // ─── ツリー D&D 開始 ─────────────────────────────────────────────────

        private WpfPoint _dragStartPoint;
        private bool     _isDragging;

        private void ObjectTree_PreviewMouseMove(object sender, WpfMouseEvent e)
        {
            if (e.LeftButton != MouseButtonState.Pressed) { _isDragging = false; return; }

            if (!_isDragging)
            {
                _dragStartPoint = e.GetPosition(null);
                _isDragging = true;
                return;
            }

            var pos  = e.GetPosition(null);
            var diff = pos - _dragStartPoint;
            if (Math.Abs(diff.X) < SystemParameters.MinimumHorizontalDragDistance &&
                Math.Abs(diff.Y) < SystemParameters.MinimumVerticalDragDistance)
                return;

            var item = GetTreeItemUnderCursor(e.OriginalSource as DependencyObject);
            if (item?.DataContext is QueryTreeNode node)
            {
                _isDragging = false;
                DragDrop.DoDragDrop(item, node.Name, WpfDragDrop.Copy);
            }
        }

        private static TreeViewItem? GetTreeItemUnderCursor(DependencyObject? source)
        {
            while (source != null && source is not TreeViewItem)
                source = System.Windows.Media.VisualTreeHelper.GetParent(source);
            return source as TreeViewItem;
        }

        // ─── エディタへのドロップ ─────────────────────────────────────────────

        private void SqlEditor_DragOver(object sender, WpfDragEvent e)
        {
            if (e.Data.GetDataPresent(WpfDataFormats.StringFormat) || HasSqlFile(e.Data))
                e.Effects = WpfDragDrop.Copy;
            else
                e.Effects = WpfDragDrop.None;
            e.Handled = true;
        }

        private void SqlEditor_Drop(object sender, WpfDragEvent e)
        {
            // ファイルドロップ（.sql）
            if (HasSqlFile(e.Data))
            {
                LoadSqlFile(e.Data);
                e.Handled = true;
                return;
            }

            // 文字列ドロップ（ツリーノード）
            if (!e.Data.GetDataPresent(WpfDataFormats.StringFormat)) return;
            var text = (string)e.Data.GetData(WpfDataFormats.StringFormat);
            var pos  = SqlEditor.GetPositionFromPoint(e.GetPosition(SqlEditor));
            if (pos != null)
            {
                var offset = SqlEditor.Document.GetOffset(pos.Value.Line, pos.Value.Column);
                SqlEditor.Document.Insert(offset, text);
            }
            else
            {
                SqlEditor.Document.Insert(SqlEditor.CaretOffset, text);
            }
            e.Handled = true;
        }

        // ─── UserControl 全体へのファイルドロップ ────────────────────────────

        private void UserControl_DragOver(object sender, WpfDragEvent e)
        {
            if (HasSqlFile(e.Data))
            {
                e.Effects = WpfDragDrop.Copy;
                e.Handled = true;
            }
        }

        private void UserControl_Drop(object sender, WpfDragEvent e)
        {
            if (!HasSqlFile(e.Data)) return;
            LoadSqlFile(e.Data);
            e.Handled = true;
        }

        // ─── ヘルパー ─────────────────────────────────────────────────────────

        private static bool HasSqlFile(WpfIDataObject data)
        {
            if (!data.GetDataPresent(WpfDataFormats.FileDrop)) return false;
            var files = data.GetData(WpfDataFormats.FileDrop) as string[];
            return files?.Any(f => f.EndsWith(".sql", StringComparison.OrdinalIgnoreCase)) == true;
        }

        private void LoadSqlFile(WpfIDataObject data)
        {
            var files   = (string[])data.GetData(WpfDataFormats.FileDrop);
            var sqlFile = files.FirstOrDefault(f =>
                f.EndsWith(".sql", StringComparison.OrdinalIgnoreCase));
            if (sqlFile == null) return;
            SqlEditor.Text = File.ReadAllText(sqlFile, Encoding.UTF8);
        }
    }
}
