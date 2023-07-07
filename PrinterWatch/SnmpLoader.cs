using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Lextm.SharpSnmpLib;
using Lextm.SharpSnmpLib.Messaging;
using Lextm.SharpSnmpLib.Security;

namespace PrinterWatch
{
    public class SNMP
    {
        public static void AddToDatabase(string printer_ip, string location)
        {
            IList<Variable> result = new List<Variable>();
            Dictionary<string, string> oid = new Dictionary<string, string>()
            {
                {"status", "1.3.6.1.2.1.25.3.5.1.1.1"},
                {"model", "1.3.6.1.2.1.25.3.2.1.3.1"},
                {"printedSheets", "1.3.6.1.2.1.43.10.2.1.4.1.1"},
                {"tonerBlackMaxLevel", "1.3.6.1.2.1.43.11.1.1.8.1.1"},
                {"tonerBlueMaxLevel", "1.3.6.1.2.1.43.11.1.1.8.1.2"},
                {"tonerRedMaxLevel", "1.3.6.1.2.1.43.11.1.1.8.1.3"},
                {"tonerYellowMaxLevel", "1.3.6.1.2.1.43.11.1.1.8.1.4"},
                {"tonerBlackCurrentLevel", "1.3.6.1.2.1.43.11.1.1.9.1.1"},
                {"tonerBlueCurrentLevel", "1.3.6.1.2.1.43.11.1.1.9.1.2"},
                {"tonerRedCurrentLevel", "1.3.6.1.2.1.43.11.1.1.9.1.3"},
                {"tonerYellowCurrentLevel", "1.3.6.1.2.1.43.11.1.1.9.1.4"}
            };
            Dictionary<string, int?> colors = new Dictionary<string, int?>() {
                {MainWindow.colorNames[0], 0},
                {MainWindow.colorNames[1], null},
                {MainWindow.colorNames[2], null},
                {MainWindow.colorNames[3], null}
            };

            /*try // Сам принтер не возвращает необходимые OID, мб с эмулятором HP LJ принтера будет лучше
            {
                result = Messenger.Get(VersionCode.V2,
                               new IPEndPoint(IPAddress.Parse(printer_ip), 161),
                               new OctetString("public"),
                               new List<Variable> {
                               new Variable(new ObjectIdentifier(oid["model"])),
                               new Variable(new ObjectIdentifier(oid["printedSheets"])),
                               new Variable(new ObjectIdentifier(oid["tonerBlueMaxLevel"])),
                               new Variable(new ObjectIdentifier(oid["tonerBlueCurrentLevel"])),
                               },
                               60000);
           */

        }

        public static void UpdateDataBaseInfo()
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
                    CommandText = "SELECT ip FROM Printers",
                    Connection = connection
                };
                SQLiteDataAdapter adapter = new SQLiteDataAdapter(command.CommandText, connection);
                DataSet printerIps = new DataSet();

                adapter.Fill(printerIps);

                foreach (DataRow a in printerIps.Tables[0].Rows)
                {
                    // Получить результат через SNMP запрос и обновить поля расходников в БД через UPDATE запрос
                }
            }
        }

    }
}
