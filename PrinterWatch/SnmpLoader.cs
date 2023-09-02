using System;
using System.Net;
using System.Data;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Windows;
using Lextm.SharpSnmpLib;
using Lextm.SharpSnmpLib.Messaging;
using System.Threading.Tasks;
using System.IO;

namespace PrinterWatch
{
    public static class SNMP
    {
        private static readonly string path = 
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments, Environment.SpecialFolderOption.Create),
                @"PrinterWatch\PrinterWatch.db");
        private static readonly Dictionary<string, string> oid = new Dictionary<string, string>()
        {
            {"status", "1.3.6.1.2.1.25.3.5.1.1.1"},
            {"model", "1.3.6.1.2.1.25.3.2.1.3.1"},
            {"printedSheets", "1.3.6.1.2.1.43.10.2.1.4.1.1"},
            {"tonerBlackCurrentLevel", "1.3.6.1.2.1.43.11.1.1.9.1.1"},
            {"tonerBlueCurrentLevel", "1.3.6.1.2.1.43.11.1.1.9.1.2"},
            {"tonerRedCurrentLevel", "1.3.6.1.2.1.43.11.1.1.9.1.3"},
            {"tonerYellowCurrentLevel", "1.3.6.1.2.1.43.11.1.1.9.1.4"},
            {"serialNumber", ".1.3.6.1.2.1.43.5.1.1.17.1" }
        };
        private static readonly Dictionary<int, string> statusDiscribtion = new Dictionary<int, string>()
        {
            {1, "иной"},
            {2, "неизвестно"},
            {3, "бездействует"},
            {4, "печатает"},
            {5, "разогревается"}
        };
        private static readonly Dictionary<string, int?> colors = new Dictionary<string, int?>() 
        {
            {MainWindow.colorNames[0], 0},
            {MainWindow.colorNames[1], null},
            {MainWindow.colorNames[2], null},
            {MainWindow.colorNames[3], null}
        };
        public static void AddToDatabase(string printerIp, string location)
        {
            IList<Variable> result = new List<Variable>();

            try
            {
                result = Messenger.Get(
                    VersionCode.V2,
                    new IPEndPoint(IPAddress.Parse(printerIp), 161),
                    new OctetString("public"),
                    new List<Variable> {
                        new Variable(new ObjectIdentifier(oid["model"])),
                        new Variable(new ObjectIdentifier(oid["serialNumber"])),
                        new Variable(new ObjectIdentifier(oid["status"])),
                        new Variable(new ObjectIdentifier(oid["tonerBlackCurrentLevel"])),
                        new Variable(new ObjectIdentifier(oid["tonerBlueCurrentLevel"])),
                        new Variable(new ObjectIdentifier(oid["tonerRedCurrentLevel"])),
                        new Variable(new ObjectIdentifier(oid["tonerYellowCurrentLevel"]))
                    },
                    5000);

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
                        CommandText = String.Format(
                            "INSERT INTO Printers VALUES ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}')",
                            printerIp, result[0].Data.ToString(), result[1].Data.ToString(), location, statusDiscribtion[Int32.Parse(result[2].Data.ToString())],
                            result[3].Data.ToString(), result[4].Data.ToString(), result[5].Data.ToString(), result[6].Data.ToString()
                        ),
                        Connection = connection
                    };
                    command.ExecuteNonQuery();
                }
            }
            catch
            {
                MessageBox.Show(String.Format("IP: {0}\nНе отвечает", printerIp));
            }
        }
        public static async void UpdateDataBaseInfo()
        {
            string conneciontStringForSqlite = new SQLiteConnectionStringBuilder(){ 
                DataSource = path
            }.ToString();

            using (var connection = new SQLiteConnection(conneciontStringForSqlite))
            {
                connection.Open();
                SQLiteCommand command = null;
                command = new SQLiteCommand()
                {
                    CommandText = "SELECT ip FROM Printers",
                    Connection = connection
                };
                SQLiteDataAdapter adapter = new SQLiteDataAdapter(command.CommandText, connection);
                DataSet printerIps = new DataSet();
                List<string> ips = new List<string>();

                adapter.Fill(printerIps);

                foreach(DataRow ip in printerIps.Tables[0].Rows)
                {
                    ips.Add(ip.ItemArray[0].ToString());
                }

                await Task.Run(() =>
                Parallel.ForEach<string>(ips, (ip) =>
                {
                    IList<Variable> result = new List<Variable>();
                    try
                    {
                        result = Messenger.Get(
                            VersionCode.V2,
                            new IPEndPoint(IPAddress.Parse(ip), 161),
                            new OctetString("public"),
                            new List<Variable>
                            {
                                new Variable(new ObjectIdentifier(oid["model"])),
                                new Variable(new ObjectIdentifier(oid["serialNumber"])),
                                new Variable(new ObjectIdentifier(oid["status"])),
                                new Variable(new ObjectIdentifier(oid["tonerBlackCurrentLevel"])),
                                new Variable(new ObjectIdentifier(oid["tonerBlueCurrentLevel"])),
                                new Variable(new ObjectIdentifier(oid["tonerRedCurrentLevel"])),
                                new Variable(new ObjectIdentifier(oid["tonerYellowCurrentLevel"]))
                            },
                            5000);
                        SQLiteCommand command = new SQLiteCommand()
                        {
                            CommandText = String.Format(
                            "UPDATE Printers SET "+
                            "Model = '{0}', " +
                            "SerialNumber = '{1}', " +
                            "Status = '{2}', " +
                            "Black = '{3}', " +
                            "Blue = '{4}', " +
                            "Red = '{5}', " +
                            "Yellow = '{6}' " +
                            "WHERE IP = '{7}'",
                            result[0].Data.ToString(), result[1].Data.ToString(),
                            statusDiscribtion[Int32.Parse(result[2].Data.ToString())], result[3].Data.ToString(),
                            result[4].Data.ToString(), result[5].Data.ToString(),
                            result[6].Data.ToString(), ip),
                            Connection = connection
                        };
                        command.ExecuteNonQuery();
                    }
                    catch (Lextm.SharpSnmpLib.Messaging.TimeoutException)
                    {
                        MessageBox.Show(String.Format("IP: {0}\nНе отвечает", ip), "Ошибка обновления", 
                            MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK, MessageBoxOptions.DefaultDesktopOnly);
                    }
                    catch
                    {
                        MessageBox.Show(String.Format("IP: {0}\nОшибка в результате запроса", ip), "Ошибка обновления", 
                            MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK, MessageBoxOptions.DefaultDesktopOnly);
                    }
                }));
                Task.WaitAll();
            }
        }
        public static async void CheckSheets()
        {
            object locker = new();
            string log = "";
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
                    CommandText = "SELECT ip FROM Printers",
                    Connection = connection
                };

                SQLiteDataAdapter adapter = new SQLiteDataAdapter(command.CommandText, connection);
                DataSet printerIps = new DataSet();
                List<string> ips = new List<string>();

                adapter.Fill(printerIps);

                foreach (DataRow ip in printerIps.Tables[0].Rows)
                {
                    ips.Add(ip.ItemArray[0].ToString());
                }

                await Task.Run(() =>
                Parallel.ForEach<string>(ips, (ip) =>
                {
                    IList<Variable> result = new List<Variable>();
                    try
                    {
                        result = Messenger.Get(
                            VersionCode.V2,
                            new IPEndPoint(IPAddress.Parse(ip), 161),
                            new OctetString("public"),
                            new List<Variable>
                            {
                                new Variable(new ObjectIdentifier("1.3.6.1.2.1.43.10.2.1.4.1.1")),
                            },
                            5000);
                        SQLiteCommand newCommand = new SQLiteCommand()
                        {
                            CommandText = String.Format(
                            @"INSERT INTO SheetsByDay (ip, sheets) VALUES ('{0}', '{1}')", ip, result[0].Data.ToString()),
                            Connection = connection
                        };
                        newCommand.ExecuteNonQuery();
                    }
                    catch (Lextm.SharpSnmpLib.Messaging.TimeoutException)
                    {
                        lock (locker)
                        {
                            log += String.Format("Не отвечает IP: {0}\n", ip);
                        }
                    }
                    catch
                    {
                        lock (locker)
                        {
                            log += String.Format("Ошибка запроса IP: {0}\n", ip);
                        }
                    }
                }));
                Task.WaitAll();
                using(FileStream fs = new FileStream(Path.GetDirectoryName(path) + @"\log.txt", FileMode.Create))
                {
                    using (StreamWriter sw = new StreamWriter(fs))
                    {
                        sw.WriteLine(log);
                    }
                }
            }
        }
    }
}
