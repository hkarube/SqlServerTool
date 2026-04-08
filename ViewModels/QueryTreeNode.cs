using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.ObjectModel;

namespace SqlServerTool.ViewModels
{
    public partial class QueryTreeNode : ObservableObject
    {
        public string Name     { get; }
        public string NodeType { get; }  // "category" / "TABLE" / "VIEW" / "PROCEDURE" / "FUNCTION" / "column" / "loading" / "error"

        private readonly Func<IEnumerable<QueryTreeNode>>? _loader;
        private bool _loaded;

        public ObservableCollection<QueryTreeNode> Children { get; } = new();

        [ObservableProperty] private bool isExpanded;

        public QueryTreeNode(string name, string nodeType,
                             Func<IEnumerable<QueryTreeNode>>? loader = null)
        {
            Name     = name;
            NodeType = nodeType;
            _loader  = loader;

            // ローダーがある場合はダミー子を挿入して展開矢印を表示する
            if (loader != null)
                Children.Add(new QueryTreeNode("読み込み中...", "loading"));
        }

        partial void OnIsExpandedChanged(bool value)
        {
            if (!value || _loaded || _loader == null) return;
            _loaded = true;
            Children.Clear();
            try
            {
                foreach (var child in _loader())
                    Children.Add(child);
            }
            catch
            {
                Children.Add(new QueryTreeNode("取得エラー", "error"));
            }
        }
    }
}
