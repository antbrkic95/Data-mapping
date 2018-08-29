using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using _Excel = Microsoft.Office.Interop.Excel;

namespace Usporedba
{
    public partial class SheetPicker : Form
    {
        public int sheetIndex = 0;
        Excel exl = null;

        public SheetPicker(Excel e)
        {
            InitializeComponent();
            exl = e;
            foreach (_Excel.Worksheet ws in exl.wb.Worksheets)
            {
                listBox1.Items.Add(ws.Name + "\r\n");
            }
        }

        private void listBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem != null)
                sheetIndex = listBox1.SelectedIndex + 1;
        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void OkBtn_Click(object sender, EventArgs e)
        {
            exl.changeSheet(sheetIndex);
            this.Close();
        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            exl.changeSheet(sheetIndex);
            this.Close();
        }
    }
}
