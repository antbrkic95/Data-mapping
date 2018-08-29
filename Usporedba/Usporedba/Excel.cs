using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Usporedba
{
    public class Excel
    {
        string path="";
        _Application excel = new _Excel.Application();
        public _Excel.Workbook wb = null;
        public _Excel.Worksheet ws;
        public int lastRow;
        public Excel()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            string fileName = "";
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                fileName = openFileDialog.FileName;
                //MessageBox.Show(fileName);
                this.path = fileName;
                this.wb = excel.Workbooks.Open(path);
                
            }
        }


        public void updateExcelRow()
        {
            object missing = Type.Missing;
            _Excel.Range range = ws.get_Range("A1", missing);
            range.Value2 = "ss";
            wb.Save();
           // ws.Cells[2, 1] = "j";
            MessageBox.Show("Success");
            
            /*//this.wb = excel.Workbooks.Open(path);
            string upit = "UPDATE [Sheet$1] SET [B]='" + mdbPathOne + "' where [B]='" + pn + "'";
            OleDbCommand cmd = new OleDbCommand();
            cmd.ExecuteNonQuery();*/
        }

        public Excel(string path, int sheet) {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }

        public void CloseExcel()
        {
            //System.Windows.Forms.MessageBox.Show("Destruktor");
            wb.Close(0);
            excel.Quit();
        }
       
        public Excel(string path)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
        }

        public int changeSheet(int sheet)
        {
            if (wb == null) return 0;
            ws = wb.Worksheets[sheet];
            return 1;
        }

        public string ReadCell(int red, int stupac) {
            if (ws.Cells[red, stupac].Value2 != null)
            {
                return ws.Cells[red, stupac].Value2.ToString();
            }
            else
                return "";
        }



        public List<string> ReadColumn(string col)
        {
            List<string> lista = new List<string>();
            List<List<string> > range = ReadRange(1, ws.UsedRange.Rows.Count, col, col);
            for (int i = 0; i < range.Count; i++)
            {
                if (range[i][0] != "")
                    lista.Add(range[i][0]);
                //if (i == range.Count - 1)
                //    MessageBox.Show(range[i][0]);
            }

            return lista;
        }

        public List<List<string>> ReadRange(int startRow, int endRow, string startColstr, string endColstr)
        {
            _Excel.Range range = ws.get_Range(startColstr + startRow.ToString(), 
                                                endColstr + endRow.ToString() );
            Array myvalues = (Array)range.Cells.Value;
            List<List<string>> ret = new List<List<string>>();
            //MessageBox.Show(myvalues.GetLength(0) + ", " + myvalues.GetLength(1));
            for (int i = 1; i <= myvalues.GetLength(0); i++)
            {
                List<string> row = new List<string>();
                for (int j = 1; j <= myvalues.GetLength(1); j++)
                {
                    if (myvalues.GetValue(i,j) == null)
                    {
                        row.Add("");
                    }
                    else
                    {
                        row.Add(myvalues.GetValue(i, j).ToString());
                    }
                }
                ret.Add(row);
            }

            return ret;
        }

        public int getMaxRowNumber()
        {
            Range usedRange = ws.UsedRange;
            this.lastRow = usedRange.Rows.Count;
            return lastRow;
        }

        public List< List< string > > ReadRange2(int startRow, int endRow, string startColstr, string endColstr)
        {
            List<List<string>> ret = new List<List<string>>();
            int startCol=0, endCol=0;
            for (int i = startColstr.Length-1; i >= 0; i--)
            {
                if (i == startColstr.Length - 1) startCol += (startColstr[i] - 'A');
                else 
                    startCol += (startColstr[i] - 'A'+1) * (startColstr.Length - 1 - i) * 26;
            }
            for (int i = endColstr.Length - 1; i >= 0; i--)
            {
                if (i == endColstr.Length - 1) endCol += (endColstr[i] - 'A');
                else
                    endCol += (endColstr[i] - 'A'+1) * (endColstr.Length - 1 - i) * 26;
            }
            for (int i = startRow; i <= endRow; i++)
            {
                List<string> row = new List<string>();
                for (int j = startCol; j <= endCol; j++)
                {
                    row.Add(ReadCell(i, j));
                }
                ret.Add(row);
            }
            return ret;
        }


        public List<List<string>> ReadRangeColumn(int startRow, int endRow, string startColstr, string endColstr)
        {
            List<List<string>> ret = new List<List<string>>();
            int startCol = 0, endCol = 0;
            for (int i = startColstr.Length - 1; i >= 0; i--)
            {
                if (i == startColstr.Length - 1) startCol += (startColstr[i] - 'A');
                else
                    startCol += (startColstr[i] - 'A' + 1) * (startColstr.Length - 1 - i) * 26;
            }
            for (int i = endColstr.Length - 1; i >= 0; i--)
            {
                if (i == endColstr.Length - 1) endCol += (endColstr[i] - 'A');
                else
                    endCol += (endColstr[i] - 'A' + 1) * (endColstr.Length - 1 - i) * 26;
            }
            for (int i = startCol; i <= endCol; i++)
            {
                List<string> col = new List<string>();
                for (int j = startRow; j <= endRow; j++)
                {
                    col.Add(ReadCell(i, j));
                }
                ret.Add(col);
            }
            return ret;
        }

        public int countColumns(string startColstr, string endColstr)
        {
            int startCol = 0, endCol = 0;
            for (int i = startColstr.Length - 1; i >= 0; i--)
            {
                if (i == startColstr.Length - 1) startCol += (startColstr[i] - 'A');
                else
                    startCol += (startColstr[i] - 'A' + 1) * (startColstr.Length - 1 - i) * 26;
            }
            for (int i = endColstr.Length - 1; i >= 0; i--)
            {
                if (i == endColstr.Length - 1) endCol += (endColstr[i] - 'A');
                else
                    endCol += (endColstr[i] - 'A' + 1) * (endColstr.Length - 1 - i) * 26;
            }
            return endCol - startCol + 1;
        }

        public int PartNoIndex(string pn)
        {
            for (int i = 1; i <= ws.UsedRange.Rows.Count; i++)
            {
                string x = ReadCell(i, 1);
                if (string.Equals(x, pn))
                {
                    return i;
                }
            }
            return -1;
        }

        public string getRangeParametar(List<string> imenaKolona)
        {
            string par = "";
            for (int i = 0; i < imenaKolona.Count; i++)
            {
                if (imenaKolona[i] == "d")
                {
                    par = getColumnName(i);
                    return par;
                    //break;
                }
                else if (imenaKolona[i] == "D")
                {
                    par = getColumnName(i);
                    return par;
                }
            }
            return "";
        }
        public string getColumnName(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }

        internal class Worksheet
        {
        }
    }
}
