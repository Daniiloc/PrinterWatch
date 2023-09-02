using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.IO;
using System.Data;
using System.Data.SQLite;
using OfficeOpenXml;
using System.Windows.Input;
using Microsoft.Win32;
using Microsoft.Win32.TaskScheduler;
using System.Threading.Tasks;

namespace PrinterWatch
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private readonly Dictionary<string, int> colorIndexes;
        private readonly string path;
        public static string[] colorNames;
        public static int criticalValue = 2;
        public MainWindow()
        {
            WindowsTask();
            InitializeComponent();
            var docFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments, Environment.SpecialFolderOption.Create);
            Directory.CreateDirectory(docFolder + @"\PrinterWatch");
            path = Path.Combine(docFolder, @"PrinterWatch\PrinterWatch.db");
            colorNames = new string[4] { "Black", "Yellow", "Red", "Blue" };
            colorIndexes = new Dictionary<string, int>();
            RealRefresh();
            SNMP.CheckSheets();
        }
        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            TestGrid.Columns[TestGrid.Columns.Count - 1].Visibility = Visibility.Collapsed;
        }
        private void WindowsTask()
        {
            bool myTask = false;
            using (TaskService ts = new TaskService())
            {
                foreach (var task in ts.AllTasks)
                {
                    if (task.Name == "PrinterSheets")
                    {
                        myTask = true;
                        break;
                    }
                }
                if (!myTask)
                {
                    TaskDefinition td = ts.NewTask();
                    td.RegistrationInfo.Description = "Получение количества напечатанных страниц принтерами в базе данных PrinterWatch";
                    td.Triggers.Add(new DailyTrigger { DaysInterval = 1, StartBoundary = DateTime.Now.AddSeconds(10) });
                    td.Actions.Add(new ExecAction(Directory.GetCurrentDirectory() + @"\PrinterWatch.exe"));
                    ts.RootFolder.RegisterTaskDefinition(@"PrinterSheets", td);
                }
            }
        }
        private void RealRefresh()
        {
            string conneciontStringForSqlite = new SQLiteConnectionStringBuilder()
            {
                DataSource = path
            }.ToString();

            using (var connection = new SQLiteConnection(conneciontStringForSqlite))
            {
                SQLiteDataAdapter adapter;
                DataSet ds = new DataSet();
                SQLiteCommand command = new SQLiteCommand();
                connection.Open();
                command = new SQLiteCommand()
                {
                    CommandText = "SELECT * from Printers",
                    Connection = connection
                };
                try
                {
                    adapter = new SQLiteDataAdapter(command);
                    adapter.Fill(ds);
                }
                catch
                {
                    command.CommandText =
                        @"CREATE TABLE Printers (
                        IP TEXT PRIMARY KEY ASC UNIQUE NOT NULL,
                        Model TEXT(10) NOT NULL,
                        SerialNumber TEXT NOT NULL,
                        Location TEXT NOT NULL,
                        Status TEXT(10) NOT NULL DEFAULT unknown,
                        Black INTEGER(100) DEFAULT(0) NOT NULL,
                        Blue INTEGER(100) DEFAULT NULL,
                        Red INTEGER(100) DEFAULT NULL,
                        Yellow INTEGER(100) DEFAULT NULL,
                        MinColor AS(min(ifnull(Black, 100), ifnull(Red, 100), ifnull(Yellow, 100), ifnull(Blue, 100))));
                        INSERT INTO Printers VALUES('1.1.1.1', 'TestModel', 'TestSerialNumber', 'TestLocation', 'unknown', '100', '100','100','100');
                        CREATE TABLE SheetsByDay (
                        Day TEXT NOT NULL
                        DEFAULT ( (CURRENT_DATE) ),
                        Ip TEXT REFERENCES Printers (IP) ON DELETE NO ACTION MATCH [FULL],
                        Sheets INTEGER,
                        PRIMARY KEY (Day, Ip));";
                    command.ExecuteNonQuery();
                    command.CommandText = "SELECT * from Printers";
                    adapter = new SQLiteDataAdapter(command);
                    adapter.Fill(ds);
                }

                colorIndexes[colorNames[0]] = ds.Tables[0].Columns.IndexOf(colorNames[0]);
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
        private void CountCartridges_Click(object sender, RoutedEventArgs e)
        {
            Dictionary<string, int> criticalLevel = new Dictionary<string, int>()
            {
                {colorNames[0], 0},
                {colorNames[1], 0},
                {colorNames[2], 0},
                {colorNames[3], 0}
            };
            Dictionary<string, string> colorNamesRussian = new Dictionary<string, string>()
            {
                {colorNames[0], "Черный" },
                {colorNames[1], "Желтый" },
                {colorNames[2], "Красный" },
                {colorNames[3], "Синий" }
            };
            foreach (DataRowView a in TestGrid.Items)
            {
                var currentRow = a.Row.ItemArray;
                if (currentRow[colorIndexes[colorNames[0]]] is not DBNull) if ((int)(long)currentRow[colorIndexes[colorNames[0]]] <= criticalValue) criticalLevel[colorNames[0]]++;
                if (currentRow[colorIndexes[colorNames[1]]] is not DBNull) if ((int)(long)currentRow[colorIndexes[colorNames[1]]] <= criticalValue) criticalLevel[colorNames[1]]++;
                if (currentRow[colorIndexes[colorNames[2]]] is not DBNull) if ((int)(long)currentRow[colorIndexes[colorNames[2]]] <= criticalValue) criticalLevel[colorNames[2]]++;
                if (currentRow[colorIndexes[colorNames[3]]] is not DBNull) if ((int)(long)currentRow[colorIndexes[colorNames[3]]] <= criticalValue) criticalLevel[colorNames[3]]++;
            }

            MessageBox.Show(String.Format("{0}: {1}\n{2}: {3}\n{4}: {5}\n{6}: {7}",
                colorNamesRussian[colorNames[0]], criticalLevel[colorNames[0]],
                colorNamesRussian[colorNames[1]], criticalLevel[colorNames[1]],
                colorNamesRussian[colorNames[2]], criticalLevel[colorNames[2]],
                colorNamesRussian[colorNames[3]], criticalLevel[colorNames[3]]),
                "Необходимые картриджи",
                MessageBoxButton.OK,
                MessageBoxImage.Information,
                MessageBoxResult.OK);
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
                string serialNumber = (TestGrid.SelectedCells[2].Column.GetCellContent(TestGrid.SelectedItem) as TextBlock).Text;
                string location = (TestGrid.SelectedCells[3].Column.GetCellContent(TestGrid.SelectedItem) as TextBlock).Text;

                if (MessageBox.Show(String.Format(
                    "Delete Printer\n" +
                    "IP: {0}\n" +
                    "Модель: {1}\n" +
                    "Серийный номер: {2}\n" +
                    "Местонахождение: {3}",
                    ip, model, serialNumber, location),
                    "Confirm Delete",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    string conneciontStringForSqlite = new SQLiteConnectionStringBuilder()
                    {
                        DataSource = path
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
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.FileName = String.Format(
                "Printers {0}.{1}.{2}-{3}.xlsx",
                System.DateTime.Now.Day.ToString(),
                System.DateTime.Now.Month.ToString(),
                System.DateTime.Now.Hour.ToString(),
                System.DateTime.Now.Minute.ToString());
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel documents (.xlsx)|*.xlsx";

            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                string filename = dlg.FileName;

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var package = new ExcelPackage();
                var sheet = package.Workbook.Worksheets.Add("Printer Report");
                var cells = sheet.Cells;

                for (int i = 0; i < TestGrid.Columns.Count - 1; i++)
                {
                    cells[1, i + 1].Value = TestGrid.Columns[i].Header;
                }

                for (int i = 0; i < TestGrid.Items.Count; i++)
                {
                    for (int k = 0; k < TestGrid.Columns.Count - 1; k++)
                    {
                        if (colorIndexes.ContainsValue(k))
                        {
                            var temp = ((DataRowView)TestGrid.Items.GetItemAt(i)).Row.ItemArray[k].ToString();
                            if (temp != "" && Int32.Parse(temp) <= criticalValue)
                            {
                                cells[i + 2, k + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.MediumGray;
                                cells[i + 2, k + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromName(TestGrid.Columns[k].Header.ToString()));
                            }
                        }
                        cells[i + 2, k + 1].Value = ((DataRowView)TestGrid.Items.GetItemAt(i)).Row.ItemArray[k];
                    }
                    cells[sheet.Dimension.Address].AutoFitColumns();
                }
                File.WriteAllBytes(filename, package.GetAsByteArray());
            }
        }
        private async void GetSheets_Click(object sender, RoutedEventArgs e)
        {
            DataSet ds = new DataSet();
            Dictionary<string, int> sheets365 = new Dictionary<string, int>();
            Dictionary<string, int> sheets182 = new Dictionary<string, int>();
            Dictionary<string, int> sheets30 = new Dictionary<string, int>();
            List<string> sheetsDefault = new List<string>();

            string conneciontStringForSqlite = new SQLiteConnectionStringBuilder()
            {
                DataSource = path
            }.ToString();

            using (SQLiteConnection connection = new SQLiteConnection(conneciontStringForSqlite))
            {
                SQLiteCommand command = new SQLiteCommand()
                {
                    CommandText = @"SELECT ip from Printers",
                    Connection = connection
                };
                SQLiteDataAdapter adapter = new SQLiteDataAdapter(command);
                adapter.Fill(ds);

                foreach (DataRow row in ds.Tables[0].Rows)
                {
                    sheets365[row.ItemArray[0].ToString()] = 0;
                    sheets182[row.ItemArray[0].ToString()] = 0;
                    sheets30[row.ItemArray[0].ToString()] = 0;
                }
            }
            await System.Threading.Tasks.Task.Run(() =>
            Parallel.ForEach<string>(sheets365.Keys, (key) =>
            {
                SQLiteConnection connection = new SQLiteConnection(conneciontStringForSqlite);

                SQLiteCommand tempCommand = new SQLiteCommand()
                {
                    CommandText = String.Format(
                            "SELECT day, sheets FROM SheetsByday WHERE ip = '{0}' ORDER BY day DESC",
                            key),
                    Connection = connection
                };
                SQLiteDataAdapter tempAdapter = new SQLiteDataAdapter(tempCommand);
                DataSet tempDs = new DataSet();
                tempAdapter.Fill(tempDs);

                var rows = tempDs.Tables[0].Rows;
                TimeSpan ts30Low = new TimeSpan(15, 0, 0, 0);
                TimeSpan ts30Up = new TimeSpan(90, 0, 0, 0);
                TimeSpan ts182Low = new TimeSpan(90, 0, 0, 0);
                TimeSpan ts182Up = new TimeSpan(272, 0, 0, 0);
                TimeSpan ts365 = new TimeSpan(272, 0, 0, 0);
                DateTime previousDate = DateTime.Now.Date;

                for (int i = 0; i < rows.Count; i++)
                {
                    var currentDate = DateTime.Parse(rows[i].ItemArray[0].ToString()).Date;
                    if (DateTime.Now.Date - currentDate >= ts30Low && 
                        DateTime.Now.Date - currentDate <= ts30Up && 
                        sheets30[key] == 0)
                    {
                        if (DateTime.Now.Date.AddMonths(-1) - previousDate < currentDate - DateTime.Now.Date.AddMonths(-1))
                        {
                            sheets30[key] = Int32.Parse(rows[0].ItemArray[1].ToString()) - Int32.Parse(rows[i].ItemArray[1].ToString());
                        }
                        else
                        {
                            sheets30[key] = Int32.Parse(rows[0].ItemArray[1].ToString()) - Int32.Parse(rows[i - 1].ItemArray[1].ToString());
                        }
                    }
                    if (DateTime.Now.Date - currentDate >= ts182Low &&
                        DateTime.Now.Date - currentDate <= ts182Up && 
                        sheets182[key] == 0)
                    {
                        if (DateTime.Now.Date.AddMonths(-6) - previousDate < currentDate - DateTime.Now.Date.AddMonths(-6))
                        {
                            sheets182[key] = Int32.Parse(rows[0].ItemArray[1].ToString()) - Int32.Parse(rows[i].ItemArray[1].ToString());
                        }
                        else
                        {
                            sheets182[key] = Int32.Parse(rows[0].ItemArray[1].ToString()) - Int32.Parse(rows[i - 1].ItemArray[1].ToString());
                        }
                    }
                    if (DateTime.Now.Date - currentDate >= ts365 && sheets365[key] == 0)
                    {
                        if (DateTime.Now.Date.AddYears(-1) - previousDate < currentDate - DateTime.Now.Date.AddYears(-1))
                        {
                            sheets365[key] = Int32.Parse(rows[0].ItemArray[1].ToString()) - Int32.Parse(rows[i].ItemArray[1].ToString());
                        }
                        else
                        {
                            sheets365[key] = Int32.Parse(rows[0].ItemArray[1].ToString()) - Int32.Parse(rows[i - 1].ItemArray[1].ToString());
                        }
                    }
                    previousDate = currentDate;
                }
            }));
            System.Threading.Tasks.Task.WaitAll();

            string result = "";
            foreach (var key in sheets365.Keys)
            {
                string tempResult = String.Format("IP: {0}\n", key);
                if (sheets30[key] > 0) tempResult += String.Format("За 30 дней: {0}\n", sheets30[key]);
                if (sheets182[key] > 0) tempResult += String.Format("За 182 дня: {0}\n", sheets182[key]);
                if (sheets365[key] > 0) tempResult += String.Format("За 365 дней: {0}\n", sheets365[key]);
                if (tempResult == String.Format("IP: {0}\n", key)) tempResult += "Недостаточно данных\n";
                result += tempResult + "\n";
            }
            MessageBox.Show(result, "Отчет по страницам");
        }
        private void TestGrid_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            DataGrid dg = sender as DataGrid;
            if (dg != null)
            {
                DataGridRow dgr = (DataGridRow)(dg.ItemContainerGenerator.ContainerFromIndex(dg.SelectedIndex));
                if (e.Key == Key.Delete && !dgr.IsEditing)
                {
                    DeletePrinterButton_Click(sender, e);
                }
            }
        }
    }
}
