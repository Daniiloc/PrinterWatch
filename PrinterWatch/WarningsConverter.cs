using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace PrinterWatch
{
    public class WarningsConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if ((int)(long)value < MainWindow.criticalValue) return new SolidColorBrush(Colors.OrangeRed);
            else if ((int)(long)value < MainWindow.criticalValue + 5) return new SolidColorBrush(Colors.Yellow);
            return new SolidColorBrush(Colors.White);
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new Exception("The method or operation is not implemented.");
        }
    }     
}
