using System.Text.RegularExpressions;
using System.Windows;

namespace PrinterWatch
{
    /// <summary>
    /// Логика взаимодействия для AddPrinterWindow.xaml
    /// </summary>
    public partial class AddPrinterWindow : Window
    {
        public AddPrinterWindow()
        {
            InitializeComponent();
        }

        private void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            Regex rg = new Regex(@"\A((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\Z");
            if (rg.IsMatch(IpTextBox.Text))
            {
                DialogResult = true;
                Close();
            }
            else
            {
                MessageBox.Show("Неправильный формат IP адреса");
            }
        }
    }
}
