using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.Data;
using System.Data.SQLite;
using OfficeOpenXml;

namespace PrinterWatch
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private Dictionary<string, int> colorIndexes;
        public static string[] colorNames;
        public static  int criticalValue = 2;
        public MainWindow()
        {
            colorNames = new string[4] { "Black", "Yellow", "Red", "Blue"};
            InitializeComponent();
            colorIndexes = new Dictionary<string, int>();
            RealRefresh();
        }
        private void RealRefresh()
        {
            string conneciontStringForSqlite = new SQLiteConnectionStringBuilder()
            {
                DataSource = "SNMP_devices.db",
            }.ToString();

            using (var connection = new SQLiteConnection(conneciontStringForSqlite))
            {
                connection.Open();
                var command = new SQLiteCommand()
                {
                    CommandText = "SELECT * from Printers",
                    Connection = connection
                };

                SQLiteDataAdapter adapter = new SQLiteDataAdapter(command.CommandText, connection);
                DataSet ds = new DataSet();

                adapter.Fill(ds);
                colorIndexes[colorNames[0]] = ds.Tables[0].Columns.IndexOf(colorNames[0]);
                colorIndexes[colorNames[1]] = ds.Tables[0].Columns.IndexOf(colorNames[1]);
                colorIndexes[colorNames[2]] = ds.Tables[0].Columns.IndexOf(colorNames[2]);
                colorIndexes[colorNames[3]] = ds.Tables[0].Columns.IndexOf(colorNames[3]);
                TestGrid.ItemsSource = ds.Tables[0].DefaultView;
            }
        }
        private void Refresh()
        {
            RealRefresh();
            TestGrid.Columns[TestGrid.Columns.Count - 1].Visibility = Visibility.Collapsed;
        }
        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            SNMP.UpdateDataBaseInfo();
            Refresh();
        }
        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            TestGrid.Columns[TestGrid.Columns.Count - 1].Visibility = Visibility.Collapsed;
        }
        private void CountWarnings_Click(object sender, RoutedEventArgs e)
        {
            Dictionary<string, int> criticalLevel = new Dictionary<string, int>()
            {
                {colorNames[0], 0},
                {colorNames[1], 0},
                {colorNames[2], 0},
                {colorNames[3], 0}
            };
            foreach (DataRowView a in TestGrid.Items)
            {
                var currentRow = a.Row.ItemArray;
                if (currentRow[colorIndexes[colorNames[0]]] is not DBNull) if ((int)(long)currentRow[colorIndexes[colorNames[0]]] <= criticalValue) criticalLevel[colorNames[0]]++;
                if (currentRow[colorIndexes[colorNames[1]]] is not DBNull) if ((int)(long)currentRow[colorIndexes[colorNames[1]]] <= criticalValue) criticalLevel[colorNames[1]]++;
                if (currentRow[colorIndexes[colorNames[2]]] is not DBNull) if ((int)(long)currentRow[colorIndexes[colorNames[2]]] <= criticalValue) criticalLevel[colorNames[2]]++;
                if (currentRow[colorIndexes[colorNames[3]]] is not DBNull) if ((int)(long)currentRow[colorIndexes[colorNames[3]]] <= criticalValue) criticalLevel[colorNames[3]]++;
            }

            MessageBox.Show(String.Format("{0} {1}\n{2} {3}\n{4} {5}\n{6} {7}",
                colorNames[0], criticalLevel[colorNames[0]], 
                colorNames[1], criticalLevel[colorNames[1]], 
                colorNames[2], criticalLevel[colorNames[2]], 
                colorNames[3], criticalLevel[colorNames[3]]));
        }
        private void AddPrinterButton_Click(object sender, RoutedEventArgs e)
        {
            AddPrinterWindow addWindow = new AddPrinterWindow();

            addWindow.Owner = this;
            addWindow.ShowDialog();
            if (addWindow.DialogResult == true)
            {
                SNMP.AddToDatabase(addWindow.IpTextBox.Text, addWindow.LocationTextBox.Text);
            }
            Refresh();
        }
        private void DeletePrinterButton_Click(object sender, RoutedEventArgs e)
        {
            if (TestGrid.SelectedItem is not null)
            {
                string ip = (TestGrid.SelectedCells[0].Column.GetCellContent(TestGrid.SelectedItem) as TextBlock).Text;
                string model = (TestGrid.SelectedCells[1].Column.GetCellContent(TestGrid.SelectedItem) as TextBlock).Text;
                string location = (TestGrid.SelectedCells[2].Column.GetCellContent(TestGrid.SelectedItem) as TextBlock).Text;

                if (MessageBox.Show(String.Format("Delete Printer:\nIP: {0}\nModel: {1}\nLocation: {2}",
                    ip, model, location),
                    "Confirm Delete",
                    MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    string conneciontStringForSqlite = new SQLiteConnectionStringBuilder()
                    {
                        DataSource = "SNMP_devices.db",
                    }.ToString();

                    using (var connection = new SQLiteConnection(conneciontStringForSqlite))
                    {
                        connection.Open();
                        SQLiteCommand command = null;
                        command = new SQLiteCommand()
                        {
                            CommandText = String.Format("DELETE FROM Printers WHERE IP = '{0}'", ip),
                            Connection = connection
                        };

                        command.ExecuteNonQuery();
                    }
                    Refresh();
                }
            }
        }
        private void ExportExcelButton_Click(object sender, RoutedEventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Printer Report");
            var cells = sheet.Cells;
            
            for(int i = 0; i < TestGrid.Columns.Count; i++)
            {
                cells[1, i + 1].Value = TestGrid.Columns[i].Header;
            }

            for (int i = 0; i < TestGrid.Items.Count; i++)
            {
                for (int k = 0; k < TestGrid.Columns.Count; k++)
                {
                    if (colorIndexes.ContainsValue(k))
                    {
                        var temp = ((DataRowView)TestGrid.Items.GetItemAt(i)).Row.ItemArray[k].ToString();
                        if(temp != "" && Int32.Parse(temp) <= criticalValue)
                        { 
                            cells[i + 2, k + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.MediumGray;
                            cells[i + 2, k + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromName(TestGrid.Columns[k].Header.ToString()));
                        }
                    }   
                    cells[i + 2, k + 1].Value = ((DataRowView)TestGrid.Items.GetItemAt(i)).Row.ItemArray[k];
                }
            }

            File.WriteAllBytes(String.Format("Printers {0}.{1}.{2}-{3}.xlsx", 
                    System.DateTime.Now.Day.ToString(),
                    System.DateTime.Now.Month.ToString(), 
                    System.DateTime.Now.Hour.ToString(),
                    System.DateTime.Now.Minute.ToString()), package.GetAsByteArray());
        }
    }
}
