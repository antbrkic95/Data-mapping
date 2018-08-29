using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Usporedba
{
    public partial class Form1 : Form
    {
        //lista part numbera
        List<string> listaExcel;
        //zastavice koje je partNo rijesio
        List<bool> bio;
        //zastavice za grupiranje PN
        List<bool> done;
        MDB m = null;
        Excel exl = null;
        OleDbConnection cn = null;
        int index;
        //Ime i opis kolone iz excela
        List<string> imenaKolona = null;
        List<string> opisKolona = null;
        List<string> posebnaImena = null;
        List<bool> usedColumn = null;
        List<bool> emptyColumn = null;
        List<List<string>> dat = null;
        //grupiranje partNumbera
        List<List<string>> groupingData = null;
        
        List<string> listzeroOne = null;
        List<string> partnumbers = null;
        List<string> excelDataZeros = null;
        List<string> excelDataOne = null;
        List<string> paths = null;
        List<string> listofMatches = null;
        //
        List<string> xmlData = null;
        List<string> dataMdb = null;
        DataTable listaMdb = null;
        //Rezultati neuspjesnih mapiranja (samo > 0%)
        Dictionary<string, Dictionary<string, double>> results;

        public Form1()
        {
            InitializeComponent();
            //this.TopMost = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            m = new MDB();
            exl = new Excel();
            SheetPicker _sp = new SheetPicker(exl);
            _sp.ShowDialog();
            //m.openFileDialog();
            listaExcel = exl.ReadColumn("A");
            bio = Enumerable.Repeat(false, listaExcel.Count).ToList();
            posebnaImena = getSpecialNames(exl);
            opisKolona = getColumnDescription(exl);
            imenaKolona = getColumnNames(exl);
        }

        private void btnDodaj_Click(object sender, EventArgs e)
        {
            if (listBox2MDB.SelectedItem == null || listBoxExcel.SelectedItem == null) return;
            string text = string.Format("Are you sure you want to map {0} with {1}",
                                listBox2MDB.SelectedItem.ToString(),
                                listBoxExcel.SelectedItem.ToString());
            var answer = MessageBox.Show(text, "", MessageBoxButtons.YesNo);
            if (answer == DialogResult.No) return;
            if (cn != null) //Ako je konekcija otvorena, dodaj mapiranje i obrisi elemente iz listboxa
            {
                int descIndex = 0, position = 0;
                for (int i = 0; i < listBoxExcel.Items.Count; i++)
                {
                    if (i == listBoxExcel.SelectedIndex) break;
                    if (listBoxExcel.SelectedItem.ToString() == listBoxExcel.Items[i].ToString())
                    {
                        position++;
                    }
                }
                for (int i = 0; i < imenaKolona.Count; i++)
                {
                    if (usedColumn[i] || emptyColumn[i]) continue;
                    if (imenaKolona[i] == listBoxExcel.SelectedItem.ToString())
                    {
                        if (position > 0)
                        {
                            position--;
                        }
                        else
                        {
                            descIndex = i;
                            break;
                        }
                    }
                }
                usedColumn[descIndex] = true;
                m.AddExtraMapping(listBoxExcel.SelectedItem.ToString(),
                                    listBox2MDB.SelectedItem.ToString(),
                                    opisKolona[descIndex],
                                    cn);
                LoadDataToMdb(listBox2MDB.SelectedItem.ToString(), descIndex);
                string dir = Path.GetDirectoryName(MDBpathLabel.Text) + Path.DirectorySeparatorChar;
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
                tw.WriteLine("MDB: " + listBox2MDB.SelectedItem.ToString() + " => " + "Excel: " + listBoxExcel.SelectedItem.ToString());
                tw.Close();
                listBox2MDB.Items.Remove(listBox2MDB.SelectedItem);
                listBoxExcel.Items.Remove(listBoxExcel.SelectedItem);
            }
        }



        private void btnUsporedi_Click(object sender, EventArgs e)
        {
            listBox2MDB.Items.Clear();
            listBoxExcel.Items.Clear();
            OutputTextbox.Clear();
            textBox1.Clear();
            textBox2.Clear();
            string endcol = exl.getColumnName(exl.ws.UsedRange.Columns.Count);
            dat = exl.ReadRange(6, exl.getMaxRowNumber(), "A", endcol);
            usedColumn = Enumerable.Repeat(false, imenaKolona.Count).ToList();
            emptyColumn = Enumerable.Repeat(false, imenaKolona.Count).ToList();
            results = new Dictionary<string, Dictionary<string, double>>();
            // Nadji novi partNo na kojem nisi bio
            int i;
            for (i = 1; i < listaExcel.Count; i++)
            {
                if (bio[i]) continue;
                bio[i] = true;
                break;           
            }
            // obradi mdb
            string path = m.getMDBpath(listaExcel[i]);
            if (path == null) return;
            listaMdb = m.getMDBdata(listaExcel[i]);
            m.CreateTable("mapping", listaExcel[i]);
            int PNordinal = listaMdb.Columns["reference"].Ordinal;
            string connectString = @"Provider = Microsoft.Jet.OLEDB.4.0; Data Source =" + m.folderPath+ Path.DirectorySeparatorChar + path;
            cn = new OleDbConnection(connectString);
            MDBpathLabel.Text = m.folderPath + Path.DirectorySeparatorChar + path;
            string dir = Path.GetDirectoryName(MDBpathLabel.Text) + Path.DirectorySeparatorChar;
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

            string sep = "";
            for (int x = 0; x < 30; x++) sep += '-';
            tw.WriteLine(sep);
            tw.WriteLine("Unique 100% matches: ");
            tw.WriteLine(sep);

            for (int k = 0; k < listaMdb.Columns.Count; ++k)
            {
                int count = exl.countColumns("A", exl.getColumnName(exl.ws.UsedRange.Columns.Count));
                List<int> candCol = Enumerable.Range(0, count).ToList();
                List<int> noOfMatches = Enumerable.Repeat(0, count).ToList();
                foreach (var item in listaMdb.AsEnumerable())
                {

                    var pn = (string)item.ItemArray[PNordinal];
                    if (listaExcel.Contains(pn))
                    {
                        bio[listaExcel.FindIndex(a => a == pn)] = true;
                        index = dat.FindIndex(drow => drow[0] == pn);
                        xmlData = dat[index];
                        string atr = item.ItemArray[k].ToString();
                        double num;
                        //if (candCol.Count == 0) break;
                        if (double.TryParse(atr, out num))
                        {
                            double num2;
                            for (int j = 0; j < xmlData.Count; j++)
                            {

                                if (double.TryParse(xmlData[j], out num2))
                                {
                                    if (num != num2)
                                    {
                                        candCol.Remove(j);

                                    }
                                    else
                                    {
                                        noOfMatches[j]++;
                                    }

                                }
                                else
                                {
                                    candCol.Remove(j);
                                }
                            }
                        }
                        else if (atr != "")
                        {
                            double num2;
                            for (int j = 0; j < xmlData.Count; j++)
                            {
                                if (!double.TryParse(xmlData[j], out num2) && xmlData[j] != "")
                                {
                                    RegexOptions options = RegexOptions.None;
                                    Regex regex = new Regex("[ ]{2,}", options);
                                    // ukloni visestruke razmake i razmake na pocetku i kraju
                                    atr = regex.Replace(atr, " ");
                                    atr = atr.Trim();
                                    // ukloni visestruke razmake i razmake na pocetku i kraju
                                    xmlData[j] = regex.Replace(xmlData[j], " ");
                                    xmlData[j] = xmlData[j].Trim();
                                    if (!string.Equals(atr, xmlData[j], StringComparison.CurrentCultureIgnoreCase))
                                    {
                                        candCol.Remove(j);
                                    }
                                    else
                                    {
                                        noOfMatches[j]++;
                                    }

                                }
                                else
                                {
                                    candCol.Remove(j);
                                }
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }

                    else if (!listaExcel.Contains(pn))
                    {
                        m.DeletePartNumber(cn, pn, MDBpathLabel.Text);
                        listaMdb.Rows.Remove(item);
                        continue;
                    }

                }
                //foreach (var cand in candCol) text += cand + ",";
                if (candCol.Count == 1)
                {
                    try
                    {
                        cn.Open();
                    }
                    catch (InvalidOperationException except) { }
                    string upit = "SELECT COUNT(mdb) FROM mapping WHERE mdb = '" + listaMdb.Columns[k].ColumnName + "'";
                    OleDbCommand cmd = new OleDbCommand(upit, cn);
                    int exists = (int)cmd.ExecuteScalar();
                    if (exists == 0)
                    {
                        m.AddExtraMapping(imenaKolona[candCol[0]],
                                            listaMdb.Columns[k].ColumnName,
                                            opisKolona[candCol[0]],
                                            cn);
                    }
                    //MessageBox.Show("Mapirao " + imenaKolona[candCol[0]]+" => "+" "+listaMdb.Columns[k]);
                    OutputTextbox.Text += "MDB: " + listaMdb.Columns[k] + " => " + "Excel: " + imenaKolona[candCol[0]] + "\r\n";
                    tw.WriteLine("MDB: " + listaMdb.Columns[k] + " => " + "Excel: " + imenaKolona[candCol[0]]);
                    usedColumn[candCol[0]] = true;
                    continue;
                }
                // SUMNJIVO mapiranje
                listBox2MDB.Items.Add(listaMdb.Columns[k].ColumnName);
                string text = listaMdb.Columns[k].ColumnName + " => ";
                foreach (var cand in candCol)
                {
                    text += imenaKolona[cand] + ", ";
                }
                int cnt = listaMdb.Rows.Count;
                text += "\n";
                string mdbCol = listaMdb.Columns[k].ColumnName;
                try
                {
                    results.Add(mdbCol, new Dictionary<string, double>());
                }
                catch { }
                for (int j = 0; j < count; j++)
                {
                    if (Math.Round(((double)noOfMatches[j] / cnt), 3) > 0.0)
                    {
                        try
                        {
                            results[mdbCol].Add(imenaKolona[j], (double)noOfMatches[j] / cnt);
                        }
                        catch { }
                        text += imenaKolona[j] + ": " + Math.Round(((double)noOfMatches[j] / cnt), 3) + "\n";
                    }
                }
                text += listaMdb.Rows.Count;
                //MessageBox.Show(text);
            }
            for (int z = 0; z < imenaKolona.Count; z++)
            {
                if (usedColumn[z] == false)
                {
                    bool empty = true;
                    foreach (var row in listaMdb.AsEnumerable())
                    {
                        var pn = (string)row.ItemArray[PNordinal];
                        index = dat.FindIndex(drow => drow[0] == pn);
                        List<string> xmlData = dat[index];
                        if (xmlData[z] != "")
                        {
                            empty = false;
                            break;
                        }
                    }
                    if (!empty) listBoxExcel.Items.Add(imenaKolona[z]);
                    else emptyColumn[z] = true;
                }
            }
            tw.WriteLine(sep);
            tw.WriteLine("User mapping: ");
            tw.WriteLine(sep);
            tw.Close();
            refreshStatusLabel();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            int index = 0;
            double numType;
            string type = "";
            List<string> listaMdbName = new List<string>();
            string name = "";
            string dir = Path.GetDirectoryName(MDBpathLabel.Text) + Path.DirectorySeparatorChar;
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
            string sep = "";
            for (int x = 0; x < 30; x++) sep += '-';
            tw.WriteLine(sep);
            tw.WriteLine("New Columns: ");
            tw.WriteLine(sep);
            for (int i = 0; i < listBoxExcel.Items.Count; i++)
            {
                listBoxExcel.SelectedIndex = i;
                listaMdbName.Clear();
                try
                {
                    string upit = "select top 1 * from data;";
                    OleDbCommand cmd = new OleDbCommand(upit, cn);
                    OleDbDataReader reader = cmd.ExecuteReader();
                    DataTable dt = new DataTable();
                    dt.Load(reader);
                    foreach (DataColumn col in dt.Columns)
                    {
                        listaMdbName.Add(col.ColumnName);
                    }
                }
                catch { }
                name = m.NameForMDB(listBoxExcel.Items[i].ToString());
                RegexOptions options = RegexOptions.None;
                Regex regex = new Regex("^" + name + "__([0-9]+)$", options);
                //int regmatch = 0;
                List<string> matches = new List<string>();
                for (int j = 0; j < listaMdbName.Count; j++)
                {
                    if (regex.IsMatch(listaMdbName[j]))
                    {
                        matches.Add(listaMdbName[j]);
                        //regmatch++;
                    }
                    else if (string.Equals(listaMdbName[j], name))
                    {
                        matches.Add(listaMdbName[j]);
                    }
                }
                if (matches.Count > 1)
                {
                    int number = -1;
                    for (int j = 0; j < matches.Count; j++)
                    {
                        if (regex.IsMatch(matches[j]))
                        {
                            string numberstr = regex.Replace(matches[j], "$1");
                            int x;
                            if (int.TryParse(numberstr, out x))
                            {
                                if (x > number) number = x;
                            }
                        }
                    }
                    number++;
                    name = name + "__" + number.ToString();
                } else if (listaMdb.Columns.Contains(name))
                {
                    name = name + "__2";
                }
                index = -1;
                for (int j = 0; j < imenaKolona.Count; ++j)
                {
                    if (imenaKolona[j] == listBoxExcel.Items[i].ToString()
                        && opisKolona[j] == textBox2.Text)
                    {
                        index = j;
                        break;
                    }
                }
                if (index == -1) throw new Exception("Index not found.");
                m.AddExtraMapping(imenaKolona[index], name, opisKolona[index], cn);
                for (int j = 0; j < dat.Count; ++j)
                {
                    if (double.TryParse(dat[j][index], out numType) == false && dat[j][index] != "")
                    {
                        type = "varchar";
                        break;
                    }
                    if (double.TryParse(dat[j][index], out numType))
                    {
                        type = "float";
                    }
                }
                if (type == "") throw new Exception("Type not correct");
                m.AddColumnsMDB(cn, name, type);
                LoadDataToMdb(name, index);
                tw.WriteLine("MDB: " + name + " => " + "Excel: " + imenaKolona[index]);
            }


            tw.WriteLine(sep);
            tw.WriteLine("tp_ Columns: ");
            tw.WriteLine(sep);
            tw.Close();
            listBoxExcel.Items.Clear();
        }

        private void LoadDataToMdb(string name, int indexCol)
        {
            int PNordinal = listaMdb.Columns["reference"].Ordinal;
            foreach (var item in listaMdb.AsEnumerable())
            {
                var pn = (string)item.ItemArray[PNordinal];

                if (listaExcel.Contains(pn))
                {
                    index = dat.FindIndex(drow => drow[0] == pn);
                    dataMdb = dat[index];
                    m.UpdateMDB(cn, name, dataMdb[indexCol], pn);
                }
                else
                {
                    continue;
                }
            }
        }

        private void refreshStatusLabel()
        {
            int res = bio.Count - 1;
            int cnt = 0;
            for (int i = 1; i < bio.Count; i++)
            {
                if (bio[i]) cnt++;
            }
            StatusLabel.Text = cnt.ToString() + " / " + res.ToString();
        }

        private List<string> getColumnNames(Excel e)
        {
            List<List<string>> listaKolone = e.ReadRange(4, 4, "A", exl.getColumnName(exl.ws.UsedRange.Columns.Count));
            for (int i = 0; i < listaKolone[0].Count; i++)
            {
                if (listaKolone[0][i] == "")
                {
                    if (opisKolona != null)
                    {
                        if (opisKolona[i] != "")
                            listaKolone[0][i] = opisKolona[i];
                        else
                        {
                            if (posebnaImena != null)
                                listaKolone[0][i] = posebnaImena[i];
                        }
                    }
                }
            }

            return listaKolone[0];

        }
        private List<string> getColumnDescription(Excel e)
        {
            List<List<string>> listaOpisa = e.ReadRange(3, 3, "A", exl.getColumnName(exl.ws.UsedRange.Columns.Count));

            return listaOpisa[0];
        }

        private List<string> getSpecialNames(Excel e)
        {
            List<List<string>> listaImena = e.ReadRange(5, 5, "A", exl.getColumnName(exl.ws.UsedRange.Columns.Count));

            return listaImena[0];
        }

        private void listBox2MDB_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Text = "";
            if (listBox2MDB.SelectedItem == null) return;
            string mdbName = listBox2MDB.SelectedItem.ToString();
            foreach (var exlName in results[mdbName])
            {
                textBox1.Text += exlName.Key + ": " + Math.Round(exlName.Value * 100, 2) + "%\r\n";
            }
        }

        private void listBoxExcel_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox2.Text = "";
            if (listBoxExcel.SelectedItem == null) return;
            int descIndex = 0, position = 0;
            for (int i = 0; i < listBoxExcel.Items.Count; i++)
            {
                if (i == listBoxExcel.SelectedIndex) break;
                if (listBoxExcel.SelectedItem.ToString() == listBoxExcel.Items[i].ToString())
                {
                    position++;
                }
            }
            for (int i = 0; i < imenaKolona.Count; i++)
            {
                if (usedColumn[i] || emptyColumn[i]) continue;
                if (imenaKolona[i] == listBoxExcel.SelectedItem.ToString())
                {
                    if (position > 0)
                    {
                        position--;
                    }
                    else
                    {
                        descIndex = i;
                        break;
                    }
                }
            }
            textBox2.Text = opisKolona[descIndex];
        }

        private void MDBpathLabel_DoubleClick(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
            {
                FileName = Path.GetDirectoryName(MDBpathLabel.Text) + Path.DirectorySeparatorChar,
                UseShellExecute = true,
                Verb = "open"
            });
        }

        private void OutputTextbox_TextChanged(object sender, EventArgs e)
        {
            OutputTextbox.SelectionStart = OutputTextbox.Text.Length;
            OutputTextbox.ScrollToCaret();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            exl.CloseExcel();
            cn.Close();
        }

        private void tp_btn_Click(object sender, EventArgs e)
        {
            m.AddTPName(cn, listBox2MDB, MDBpathLabel.Text);

        }

        private void btnAddFamily_Click(object sender, EventArgs e)
        {
            string dir = Path.GetDirectoryName(@"C:\NOVO\") + Path.DirectorySeparatorChar;
            TextWriter tw = null;
            if (!File.Exists(dir + "data.log"))
            {
                File.Create(dir + "data.log").Dispose();
                tw = new StreamWriter(dir + "data.log");
            }
            else if (File.Exists(dir + "data.log"))
            {
                tw = new StreamWriter(dir + "data.log", true);
            }
            string sep = "";
            for (int x = 0; x < 30; x++) sep += '-';
            tw.WriteLine(sep);
            tw.WriteLine("Part Number " + " ------  " + " Path");
            tw.WriteLine(sep);
            // exl.updateExcelRow();
            //string path = @"C:\NOVO\";
            excelDataZeros = new List<string>();
            excelDataOne = new List<string>();
            listofMatches = new List<string>();
            paths = exl.ReadColumn("B");
            partnumbers = exl.ReadColumn("C");
            listzeroOne = exl.ReadColumn("A");
            done = Enumerable.Repeat(false, listzeroOne.Count).ToList();
            groupingData = exl.ReadRange(6, exl.getMaxRowNumber(), "I", exl.getRangeParametar(imenaKolona));
            int count = exl.countColumns("I", exl.getRangeParametar(imenaKolona));
            for (int i = 0 ; i < listzeroOne.Count; i++)
            {
                if (listzeroOne.Contains("0"))
                {
                    if (listzeroOne[i] == "0")
                    {
                        excelDataZeros = groupingData[i];
                        for (int k = 0; k < listzeroOne.Count; k++)
                        {
                            int counter = 0;
                            if (listzeroOne[k] == "1")
                            {
                                excelDataOne = groupingData[k];
                                for (int j = 0; j < count; j++)
                                {
                                    if (excelDataOne[j] == excelDataZeros[j])
                                    {
                                        //listofMatches.Add(paths[k]);
                                        counter++;

                                    }
                                    else
                                    {
                                        continue;
                                    }
           
                                 }

                                if (counter == count)
                                {
                                    tw.WriteLine(partnumbers[i + 1] + "  ------  " + paths[k]);
                                    break;
                                }
                                                               
                            }
                        }

                    }
                }
                else
                {
                    MessageBox.Show("There is no part numbers to be grouped!");
                    break;
                }
              
            }
            MessageBox.Show("done");
           
            tw.Close();

        }

    }
}
