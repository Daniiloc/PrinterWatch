using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

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
            Regex rg = new Regex(@"\A\d{1,3}.\d{1,3}.\d{1,3}.\d{1,3}\Z");
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
