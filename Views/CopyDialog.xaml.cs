using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows;
using WpfKey = System.Windows.Input.Key;
using WpfKeyArgs = System.Windows.Input.KeyEventArgs;

namespace SqlServerTool.Views
{
    public partial class CopyDialog : Window
    {
        private readonly ISet<string> _existingNames;

        /// <summary>ユーザーが入力したコピー先テーブル名</summary>
        public string DestName { get; private set; } = string.Empty;

        public CopyDialog(string sourceName, ISet<string> existingNames)
        {
            InitializeComponent();
            _existingNames = existingNames;

            SourceNameBox.Text = sourceName;
            DestNameBox.Text   = sourceName;
            DestNameBox.SelectAll();
            DestNameBox.Focus();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
            => TryAccept();

        private void DestNameBox_KeyDown(object sender, WpfKeyArgs e)
        {
            if (e.Key == WpfKey.Enter) TryAccept();
        }

        private void TryAccept()
        {
            var name  = DestNameBox.Text.Trim();
            var error = Validate(name);
            if (error != null)
            {
                ErrorText.Text       = error;
                ErrorText.Visibility = Visibility.Visible;
                DestNameBox.Focus();
                return;
            }

            DestName     = name;
            DialogResult = true;
        }

        private string? Validate(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "名前を入力してください。";

            if (name.Length > 128)
                return "名前は128文字以内にしてください。";

            if (!Regex.IsMatch(name, @"^[\p{L}_#@][\p{L}\p{N}_#@$]*$"))
                return "使用できない文字が含まれています。\n（先頭: 文字 / _ / # / @、以降: 文字・数字 / _ / # / @ / $）";

            if (_existingNames.Contains(name))
                return $"「{name}」はすでに存在します。別の名前を入力してください。";

            return null;
        }
    }
}
