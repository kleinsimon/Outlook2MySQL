using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using MySql.Data.MySqlClient;
using Microsoft.Office.Interop.Outlook;

namespace Contacts2SQL
{
    public partial class MainForm : Form
    {
        DataTable EntryTable = new DataTable();

        MySqlConnectionStringBuilder sqlconn = new MySqlConnectionStringBuilder();

        public Microsoft.Office.Interop.Outlook.Application objOutlook = new Microsoft.Office.Interop.Outlook.Application();
        public Microsoft.Office.Interop.Outlook._NameSpace objNS;
        bool AutoMode = false;

        public string[] Columns =
        {
            "Name",
            "Vorname",
            "Titel",
            "Firma",
            "Abteilung",
            "Straße",
            "PLZ",
            "Ort",
            "Telefon",
            "Fax",
            "E-Mail",
        };

        public MainForm()
        {
            if (Environment.GetCommandLineArgs().Contains("-auto"))
            {
                AutoMode = true;
            }

            InitializeComponent();
            if (AutoMode) {
                this.Hide();
                this.SuspendLayout();
            }

            foreach (string tag in Columns)
            {
                EntryTable.Columns.Add(tag);
            }

            dataGridView1.DataSource = EntryTable;

            objNS = objOutlook.Session;

            if (AutoMode)
            {
                getContactData();
                updateTable();
                this.Dispose();
                this.Close();
                Environment.Exit(0);
            }
        }


        int countCharNum(string source, char needle)
        {
            int count = 0;
            foreach (char c in source)
                if (c == needle) count++;
            return count;
        }

        void MakeCol(string col, DataTable table)
        {
            if (!table.Columns.Contains(col))
                table.Columns.Add(col);
        }

        void dumpSQL()
        {

        }

        private void buttonConnect_Click(object sender, EventArgs e)
        {
            updateTable();
        }

        public void updateTable(bool noChange = false)
        {
            Process plink = null;

            if (Properties.Settings.Default.SSHtunnel)
            {
                try
                {
                    sqlconn.Server = @"localhost";
                    sqlconn.Port = uint.Parse(Properties.Settings.Default.SSHPort);

                    ProcessStartInfo psi = new ProcessStartInfo(Properties.Settings.Default.plinkPath);

                    psi.Arguments = "-ssh -l " + Properties.Settings.Default.SSHUser + " -L " + Properties.Settings.Default.port + ":localhost:" +
                                    Properties.Settings.Default.SSHPort + " -pw " + Properties.Settings.Default.SSHPass + " " + Properties.Settings.Default.server;


                    psi.RedirectStandardInput = true;
                    psi.RedirectStandardOutput = true;
                    psi.WindowStyle = ProcessWindowStyle.Hidden;
                    psi.UseShellExecute = false;
                    psi.CreateNoWindow = true;

                    plink = Process.Start(psi);
                    
                }
                catch
                {
                    MessageBox.Show("SSH-Tunnel konnte nicht aufgebaut werden");
                    return;
                }
            }
            else
            {
                sqlconn.Server = Properties.Settings.Default.server;
                sqlconn.Port = Properties.Settings.Default.port;
            }

            sqlconn.UserID = Properties.Settings.Default.user;
            sqlconn.Password = Properties.Settings.Default.password;
            sqlconn.Database = Properties.Settings.Default.database;

            string table = Properties.Settings.Default.table;
            EntryTable.TableName = table;

            using (MySqlConnection connection = new MySqlConnection(sqlconn.ConnectionString))
            {
                try
                {
                    connection.Open();
                    if (!connection.Ping())
                    {
                        throw new System.Exception("No Connection to Server");
                    }

                    if (!noChange)
                    {
                        MySqlCommand del = new MySqlCommand(@"TRUNCATE TABLE " + sqlconn.Database + "." + table, connection);
                        del.ExecuteNonQuery();

                        MySqlDataAdapter adapter = new MySqlDataAdapter(@"SELECT * FROM " + sqlconn.Database + "." + table, connection);
                        MySqlCommandBuilder builder = new MySqlCommandBuilder(adapter);
                        adapter.ContinueUpdateOnError = true;

                        adapter.Update(EntryTable);
                    }
                    if (!AutoMode) MessageBox.Show("Erfolgreich");
                }
                catch (System.Exception ex)
                {
                    string tmp = "";
                    if (plink != null)
                        tmp = plink.StandardOutput.ReadToEnd();
                    MessageBox.Show(ex.Message.ToString() + tmp);
                }
            }

            try
            {
                plink.StandardInput.WriteLine("exit");

                plink.Kill();
                plink.Dispose();
            }
            catch { }

            //EntryTable.Clear();
        }

        private void buttonSettings_Click(object sender, EventArgs e)
        {
            settings setWin = new settings(this);
            setWin.Show();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.Save();
        }

        private MAPIFolder getFolderFromPath(string Path, _NameSpace ns, MAPIFolder parent = null)
        {
            if (Path.StartsWith(@"\\"))
                Path = Path.Substring(2);

            if (Path.EndsWith(@"\"))
                Path = Path.Remove(Path.Length - 1);




            if (Path.Contains(@"\"))
            {
                MAPIFolder parentFolder;

                string[] subFolders = Path.Split(new char[] { '\\' }, 2);

                if (parent == null)
                    parentFolder = ns.Folders[subFolders[0]];
                else
                    parentFolder = parent.Folders[subFolders[0]];

                return getFolderFromPath(subFolders[1], ns, parentFolder);
            }
            else
            {
                if (parent == null)
                {
                    return ns.Folders[Path];

                }
                else
                {
                    return parent.Folders[Path];
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            getContactData();
            buttonConnect.Enabled = true;
        }

        private void getContactData()
        {
            MAPIFolder f;
            Items itms;

            try
            {
                f = getFolderFromPath(Properties.Settings.Default.outlookPath, objNS);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Fehler: Ordner konnte nicht geöffnet werden..." + Environment.NewLine + ex.Message);
                return;
            }

            if (Properties.Settings.Default.locationFilter != "")
            {
                try
                {
                    itms = f.Items.Restrict(Properties.Settings.Default.locationFilter);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Fehler: Filter konnte nicht angewendet werden..." + Environment.NewLine + ex.Message);
                    return;
                }
            }
            else
            {
                itms = f.Items;
            }

            this.Cursor = Cursors.WaitCursor;
            EntryTable.BeginLoadData();

            try
            {
                foreach (var tmp in itms)
                {
                    if (!(tmp is ContactItem)) continue;

                    ContactItem cIt = (ContactItem)tmp;
                    DataRow newRow = EntryTable.NewRow();

                    newRow["Name"] = makeValidString(cIt.LastName);
                    newRow["Vorname"] = makeValidString(cIt.FirstName);
                    newRow["Titel"] = makeValidString(cIt.User2);
                    newRow["Firma"] = makeValidString(cIt.CompanyName);
                    newRow["Abteilung"] = makeValidString(cIt.Department);
                    newRow["Strasse"] = makeValidString(cIt.BusinessAddressStreet);
                    newRow["PLZ"] = makeValidString(cIt.BusinessAddressPostalCode);
                    newRow["Ort"] = makeValidString(cIt.BusinessAddressCity);
                    newRow["Telefon"] = makeValidString(cIt.BusinessTelephoneNumber);
                    newRow["Fax"] = makeValidString(cIt.BusinessFaxNumber);
                    newRow["E-Mail"] = makeValidString(cIt.Email1Address);

                    EntryTable.Rows.Add(newRow);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            EntryTable.EndLoadData();
            this.Cursor = Cursors.Default;
        }

        string makeValidString(string s)
        {
            if (s == null || s==string.Empty)
            {
                return " ";
            }
            return s;
        }

        void OL_AdvancedSearchComplete(Search SearchObject)
        {
            MessageBox.Show("bla");
        }

    }

}
