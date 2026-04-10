using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;
using WpfColor   = System.Windows.Media.Color;
using WpfBrushes = System.Windows.Media.Brushes;

namespace SqlServerTool.Helpers
{
    /// <summary>
    /// セル値が null または DBNull の場合に黄色ブラシを返す。それ以外は Transparent。
    /// DataGrid の ElementStyle/Background バインドで使用する。
    /// </summary>
    public class NullBackgroundConverter : IValueConverter
    {
        private static readonly SolidColorBrush YellowBrush
            = new SolidColorBrush(WpfColor.FromRgb(0xFF, 0xF0, 0x80));

        public object Convert(object value, Type targetType,
            object parameter, CultureInfo culture)
            => (value is null || value == System.DBNull.Value)
                ? YellowBrush
                : WpfBrushes.Transparent;

        public object ConvertBack(object value, Type targetType,
            object parameter, CultureInfo culture)
            => throw new NotImplementedException();
    }

    /// <summary>
    /// セル値が null または DBNull の場合に "(NULL)" 文字列を返す。
    /// DataGrid の Binding Converter で使用する。
    /// </summary>
    public class DbNullDisplayConverter : IValueConverter
    {
        public object Convert(object value, Type targetType,
            object parameter, CultureInfo culture)
            => (value is null || value == System.DBNull.Value)
                ? "(NULL)"
                : value.ToString()!;

        public object ConvertBack(object value, Type targetType,
            object parameter, CultureInfo culture)
            => throw new NotImplementedException();
    }
}
