using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Usporedba
{
    class MDB
    {
        string charsToIgnore = ".!`[]\"";
        public string filePath = "";
        public string folderPath = "";
        int maxLen = 64;

        public MDB() {
            FolderBrowserDialog openFolder = new FolderBrowserDialog();
            openFolder.Description = "Select the document folder";
            string folderName = "";
            
            if (openFolder.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                folderName = openFolder.SelectedPath;
                folderPath = folderName;
            }
            OpenFileDialog open = new OpenFileDialog();
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "mdb files (*.mdb)|*.mdb|All files (*.*)|*.*";
            string fileName = "";

            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileName = openFileDialog.FileName;
                filePath = fileName;
            }
           
        }

        public string getMDBpath(string reference)
        {
            string ret = null;
            string upit;
            string path = filePath;
            string connectString = @"Provider = Microsoft.Jet.OLEDB.4.0; Data Source =" + path;
            OleDbConnection cn = new OleDbConnection(connectString);
            cn.Open();
            OleDbDataReader reader;

            try
            {
                upit = "select mdb_path from path_reference where reference='" + reference + "';";
                OleDbCommand cmd = new OleDbCommand(upit, cn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    //row = new object[2];
                    //reader.GetValues(row);
                    ret = reader[0].ToString();
                }
            }
            catch { }
            cn.Close();
            return ret;
        }

        public void openMdb(OleDbConnection cn,string res)
        {
            
            try
            {
                cn.Open();
            }
            catch (InvalidOperationException e) { }
            string upit = "";
            upit = "insert into data(reference) " +
                    "values('" + res + "')";
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = cn;
            cmd.CommandText = upit;
            cmd.ExecuteNonQuery();
            cn.Close();
        }
        
        public DataTable getMDBdata(string reference)
        {
            string upit;
            DataTable lista = new DataTable();
            string path = getMDBpath(reference);
            //MessageBox.Show(path);
            string connectString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + folderPath + Path.DirectorySeparatorChar + path;
            OleDbConnection cn = new OleDbConnection(connectString);
            cn.Open();
            OleDbDataReader reader;
            try
            {
                upit = "select * from data;";
                OleDbCommand cmd = new OleDbCommand(upit, cn);
                reader = cmd.ExecuteReader();
                lista.Load(reader);
            }
            catch { }
            cn.Close();
            return lista;
        }

        public void CreateTable(string name, string reference)
        {
            string upit;
            string path = getMDBpath(reference);
            string connectString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + folderPath + Path.DirectorySeparatorChar + path;
            OleDbConnection cn = new OleDbConnection(connectString);
            cn.Open();
            List<string> tableNames = GetTableNames(cn);
            bool exists = false;
            foreach (string n in tableNames)
            {
                if (n.Contains("mapping"))
                {
                    exists = true;
                    break;

                }
            }
            if (!exists)
            {
                upit = "CREATE TABLE " + name +
                    "([id] AUTOINCREMENT NOT NULL PRIMARY KEY," +
                    "[sheet] VARCHAR(50)," +
                    "[mdb] VARCHAR(50)," +
                    "[description] VARCHAR(100)); ";
                OleDbCommand cmd = new OleDbCommand(upit, cn);
                cmd.ExecuteNonQuery();
                //AddStandardMapping(cn);
            }

            cn.Close();
        }
        //maknuti navodnike
        public void AddColumnsMDB(OleDbConnection cn, string colName, string type)
        {

            try
            {
                cn.Open();
            }
            catch (InvalidOperationException e) { }
            string upit = "";       
            upit = "alter table data add " + colName + "" + " " + type;
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = cn;
            cmd.CommandText = upit;
            cmd.ExecuteNonQuery(); 
        }

        public void DeletePartNumber(OleDbConnection cn, string reference,string path)
        {
            string dir = Path.GetDirectoryName(path) + Path.DirectorySeparatorChar;
            TextWriter tw = null;
            if (!File.Exists(dir + "Not existing PartNumbers.log"))
            {
                File.Create(dir + "Not existing PartNumbers.log").Dispose();
                tw = new StreamWriter(dir + "Not existing PartNumbers.log");

            }
            else if (File.Exists(dir + "Not existing PartNumbers.log"))
            {
                tw = new StreamWriter(dir + "Not existing PartNumbers.log", true);
            }
            try
            {
                cn.Open();
            }
            catch (InvalidOperationException e) { }
            string upit = "";
            upit = "delete * from data where reference='" + reference + "'";
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = cn;
            cmd.CommandText = upit;
            cmd.ExecuteNonQuery();
            tw.WriteLine(reference);
            tw.Close();
        }

        public void UpdateMDB(OleDbConnection cn, string colName, string value, string reference)
        {

            try
            {
                cn.Open();
            }
            catch (InvalidOperationException e) { }
            string upit = "";

            if (value == "")
            {
                value = 0.ToString();
            }
            upit = "update data set " + colName+"='"+value+"' where reference='"+reference+"'";
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = cn;
            cmd.CommandText = upit;
            cmd.ExecuteNonQuery();
        }

        public void AddStandardMapping(OleDbConnection con)
        {
            try
            {
                con.Open();
            }
            catch (InvalidOperationException e) { /*vec otvorena konekcija*/ }
            List<string> tableNames = GetTableNames(con);
            string upit;
            upit = "INSERT INTO mapping(sheet, mdb) VALUES ";
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = con;
            cmd.CommandText = upit + "('Designation', 'reference');";
            cmd.ExecuteNonQuery();
            cmd.CommandText = upit + "('Internal ID', 'internal_id');";
            cmd.ExecuteNonQuery();
            cmd.CommandText = upit + "('EAN', 'eeaann');";
            cmd.ExecuteNonQuery();
            cmd.CommandText = upit + "('Short description', 'short_description');";
            cmd.ExecuteNonQuery();
            cmd.CommandText = upit + "('Long description', 'long_description');";
            cmd.ExecuteNonQuery();
            cmd.CommandText = upit + "('URL', 'uurrll');";
            cmd.ExecuteNonQuery();
            cmd.CommandText = upit + "('Group ID', 'group_id');";
            cmd.ExecuteNonQuery();
        }

        public List<string> GetTableNames(OleDbConnection con)
        {
            List<string> tableNames = new List<string>();
            try
            {
                con.Open();
            }
            catch (InvalidOperationException e) { /*vec otvorena konekcija*/ }
            var tables = con.GetSchema("Tables");
            tableNames = tables.AsEnumerable().Select(drow => drow.Field<string>(2)).ToList();
            return tableNames;
        }
        

        public void AddTPName(OleDbConnection cn, ListBox lbMDB, string path)
        {
            string nameTP = "";
            string dir = Path.GetDirectoryName(path) + Path.DirectorySeparatorChar;
            TextWriter tw = null;
            if (!File.Exists(dir + "mapping.log"))
            {
                File.Create(dir + "mapping.log").Dispose();
                tw = new StreamWriter(dir + "mapping.log");

            }
            else if (File.Exists(dir + "mapping.log"))
            {
                tw = new StreamWriter(dir + "mapping.log", true);
            }

            try
            {
                cn.Open();
            }
            catch (InvalidOperationException e) { }
            List<string> listaMdb = new List<string>();
            OleDbDataReader reader;
            try
            {
                string upit = "select sheet from mapping;";
                OleDbCommand cmd = new OleDbCommand(upit, cn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    object[] data = new object[1];
                    reader.GetValues(data);
                    listaMdb.Add((string)data[0]);
                }
            }
            catch { }
            string odabraniCol = lbMDB.SelectedItem as string;

            if (odabraniCol == null)
            {
                MessageBox.Show("You have to choose column!");
            }
            else if (odabraniCol != null)
            {
                nameTP = "tp_" + odabraniCol.ToString();
            }
            if (!listaMdb.Contains(nameTP))
            {
                if (odabraniCol != null)
                {
                    string upit = "insert into mapping(sheet,mdb) " +
                    "values('" + "tp_" + odabraniCol + "','" + odabraniCol.ToString() + "')";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = cn;
                    cmd.CommandText = upit;
                    cmd.ExecuteNonQuery();
                }
            }
            tw.WriteLine("MDB: " + odabraniCol + " => " + nameTP);
            tw.Close();
            lbMDB.Items.Remove(odabraniCol);

        }

        public void AddExtraMapping(string sheet, string mdb, string description, OleDbConnection cn)
        {
            try
            {
                cn.Open();
            }
            catch (InvalidOperationException e) { }
            string desc;
            if (description != "")
            {
                desc = "'" + description + "'";
            } else
            {
                desc = "NULL";
            }
            string upit = "insert into mapping(sheet,mdb,description) " +
                "values('" + sheet + "','" + mdb + "'," + desc + ")";
            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = cn;
            cmd.CommandText = upit;
            cmd.ExecuteNonQuery();
        }

        public string NameForMDB(string oldName)
        {
            StringBuilder newName = new StringBuilder();
            string ret = "";
            oldName = oldName.Trim();
            int words = oldName.Count(a => a == ' ') + 1;
            for (int i = 0; i < oldName.Length; i++)
            {
                if (char.IsUpper(oldName[i]))
                {
                    newName.Append(char.ToLower(oldName[i]));
                    if (words == 1) newName.Append(char.ToLower(oldName[i]));
                }
                else if (oldName[i] == '\u03B1')
                {
                    newName.Append("alpha");
                }
                else if (oldName[i] == '\u03B2')
                {
                    newName.Append("beta");
                }
                else if (oldName[i] == '\u03B3')
                {
                    newName.Append("gama");
                }
                else if (oldName[i] == '\u03B4')
                {
                    newName.Append("delta");
                }
                else if (oldName[i] == '\u03B5')
                {
                    newName.Append("epsilon");
                }
                else if (oldName[i] == '\u03B7')
                {
                    newName.Append("eta");
                }
                else if (oldName[i] == '\u03B8')
                {
                    newName.Append("theta");
                }
                else if (oldName[i] == '\u03BB')
                {
                    newName.Append("lambda");
                }
                else if (oldName[i] == '\u03BC')
                {
                    newName.Append("mu");
                }
                else if (oldName[i] == '\u03C0')
                {
                    newName.Append("pi");
                }
                else if (oldName[i] == '\u03C1')
                {
                    newName.Append("rho");
                }
                else if (oldName[i] == '\u03C3')
                {
                    newName.Append("sigma");
                }
                else if (oldName[i] == '\u03C9')
                {
                    newName.Append("omega");
                }
                else if (charsToIgnore.Contains(oldName[i]))
                {
                    continue;
                }
                else if (!char.IsNumber(oldName[i]) && !char.IsLetter(oldName[i]))
                {
                    newName.Append("_");
                }
                else
                {
                    newName.Append(oldName[i]);
                }
            }
            if (newName.Length > maxLen)
            {
                ret = Regex.Replace(newName.ToString(), "([a-z])([a-z]*)", "$1");
            } else
            {
                ret = newName.ToString();
            }


            return ret;
        }


    
    }
    }



