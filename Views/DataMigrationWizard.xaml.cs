using SqlServerTool.ViewModels;
using System.Windows;
using WpfControls = System.Windows.Controls;

namespace SqlServerTool.Views
{
    public partial class DataMigrationWizard : Window
    {
        private readonly DataMigrationViewModel _vm;

        public DataMigrationWizard(DataMigrationViewModel viewModel)
        {
            InitializeComponent();
            _vm = viewModel;
            DataContext = _vm;

            // PasswordBox はデータバインディング非対応のため code-behind で処理
            DestPasswordBox.PasswordChanged  += (_, _) => _vm.DestPassword  = DestPasswordBox.Password;
            NewDbPasswordBox.PasswordChanged += (_, _) => _vm.NewDbPassword = NewDbPasswordBox.Password;

            // CurrentStep の変化に合わせてステップインジケーターを更新
            _vm.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(_vm.CurrentStep))
                    UpdateStepIndicator();
            };
        }

        // ─── ステップインジケーター更新 ──────────────────────────────────────

        private void UpdateStepIndicator()
        {
            int step = _vm.CurrentStep;

            // Step1
            StepHeader1.Background = step == 0
                ? System.Windows.Media.Brushes.SteelBlue
                : System.Windows.Media.Brushes.MediumSeaGreen;

            // Step2
            StepHeader2.Background = step == 1
                ? System.Windows.Media.Brushes.SteelBlue
                : (step > 1 ? System.Windows.Media.Brushes.MediumSeaGreen
                             : System.Windows.Media.Brushes.LightGray);
            StepHeader2Text.Foreground = step >= 1
                ? System.Windows.Media.Brushes.White
                : System.Windows.Media.Brushes.Gray;

            // Step3
            StepHeader3.Background = step == 2
                ? System.Windows.Media.Brushes.SteelBlue
                : System.Windows.Media.Brushes.LightGray;
            StepHeader3Text.Foreground = step == 2
                ? System.Windows.Media.Brushes.White
                : System.Windows.Media.Brushes.Gray;

            // 「次へ」ボタンのラベル
            NextButton.Content    = step == 2 ? "閉じる" : "次へ ▶";
            NextButton.Visibility = Visibility.Visible;
            PrevButton.Visibility = step == 0 ? Visibility.Collapsed : Visibility.Visible;
        }

        // ─── 出力方法ラジオボタン ─────────────────────────────────────────────

        private void OutputMethod_Checked(object sender, RoutedEventArgs e)
        {
            if (sender is WpfControls.RadioButton rb)
                _vm.OutputMethod = rb.Tag?.ToString() ?? "DB";
        }

        // ─── 宛先 DB モードラジオボタン ──────────────────────────────────────

        private void DestDbMode_Checked(object sender, RoutedEventArgs e)
        {
            if (sender is WpfControls.RadioButton rb)
                _vm.DestDbMode = rb.Tag?.ToString() ?? "Existing";
        }

        // ─── 認証ラジオボタン ─────────────────────────────────────────────────

        private void DestAuth_Checked(object sender, RoutedEventArgs e)
        {
            if (sender is WpfControls.RadioButton rb)
                _vm.DestWindowsAuth = rb.Tag?.ToString() == "Windows";
        }

        // ─── 次へ / 閉じる ───────────────────────────────────────────────────

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            if (_vm.CurrentStep == 2)
                Close();
            else
                _vm.NextStepCommand.Execute(null);
        }

        // ─── キャンセル ───────────────────────────────────────────────────────

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            if (_vm.IsRunning) return;
            Close();
        }
    }
}
