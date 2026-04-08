using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows;
using WpfKey = System.Windows.Input.Key;
using WpfKeyArgs = System.Windows.Input.KeyEventArgs;

namespace SqlServerTool.Views
{
    public partial class RenameDialog : Window
    {
        private readonly string _currentName;
        private readonly ISet<string> _existingNames;

        /// <summary>ユーザーが入力した新しい名前</summary>
        public string NewName { get; private set; } = string.Empty;

        public RenameDialog(string currentName, ISet<string> existingNames)
        {
            InitializeComponent();
            _currentName   = currentName;
            _existingNames = existingNames;

            OldNameBox.Text = currentName;
            NewNameBox.Text = currentName;
            NewNameBox.SelectAll();
            NewNameBox.Focus();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
            => TryAccept();

        private void NewNameBox_KeyDown(object sender, WpfKeyArgs e)
        {
            if (e.Key == WpfKey.Enter) TryAccept();
        }

        private void TryAccept()
        {
            var name = NewNameBox.Text.Trim();
            var error = Validate(name);
            if (error != null)
            {
                ErrorText.Text       = error;
                ErrorText.Visibility = Visibility.Visible;
                NewNameBox.Focus();
                return;
            }

            NewName      = name;
            DialogResult = true;
        }

        private string? Validate(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "名前を入力してください。";

            if (name.Length > 128)
                return "名前は128文字以内にしてください。";

            // SQL Server 識別子として無効な文字を簡易チェック
            if (!Regex.IsMatch(name, @"^[A-Za-z_#@][A-Za-z0-9_#@$]*$"))
                return "使用できない文字が含まれています。\n（先頭: 英字 / _ / # / @、以降: 英数字 / _ / # / @ / $）";

            if (string.Equals(name, _currentName, System.StringComparison.OrdinalIgnoreCase))
                return "現在の名前と同じです。";

            if (_existingNames.Contains(name))
                return $"「{name}」はすでに存在します。別の名前を入力してください。";

            return null;
        }
    }
}
