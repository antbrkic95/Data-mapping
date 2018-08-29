namespace Usporedba
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.listBoxExcel = new System.Windows.Forms.ListBox();
            this.btnDodaj = new System.Windows.Forms.Button();
            this.listBox2MDB = new System.Windows.Forms.ListBox();
            this.btnUsporedi = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.MDBpathLabel = new System.Windows.Forms.Label();
            this.OutputTextbox = new System.Windows.Forms.TextBox();
            this.btnAdd = new System.Windows.Forms.Button();
            this.StatusLabel = new System.Windows.Forms.Label();
            this.tp_btn = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnAddFamily = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // listBoxExcel
            // 
            this.listBoxExcel.FormattingEnabled = true;
            this.listBoxExcel.Location = new System.Drawing.Point(506, 49);
            this.listBoxExcel.Name = "listBoxExcel";
            this.listBoxExcel.Size = new System.Drawing.Size(238, 199);
            this.listBoxExcel.TabIndex = 0;
            this.listBoxExcel.SelectedIndexChanged += new System.EventHandler(this.listBoxExcel_SelectedIndexChanged);
            // 
            // btnDodaj
            // 
            this.btnDodaj.Location = new System.Drawing.Point(106, 280);
            this.btnDodaj.Name = "btnDodaj";
            this.btnDodaj.Size = new System.Drawing.Size(97, 23);
            this.btnDodaj.TabIndex = 1;
            this.btnDodaj.Text = "Map selected";
            this.btnDodaj.UseVisualStyleBackColor = true;
            this.btnDodaj.Click += new System.EventHandler(this.btnDodaj_Click);
            // 
            // listBox2MDB
            // 
            this.listBox2MDB.FormattingEnabled = true;
            this.listBox2MDB.Location = new System.Drawing.Point(25, 49);
            this.listBox2MDB.Name = "listBox2MDB";
            this.listBox2MDB.Size = new System.Drawing.Size(238, 199);
            this.listBox2MDB.TabIndex = 2;
            this.listBox2MDB.SelectedIndexChanged += new System.EventHandler(this.listBox2MDB_SelectedIndexChanged);
            // 
            // btnUsporedi
            // 
            this.btnUsporedi.Location = new System.Drawing.Point(25, 280);
            this.btnUsporedi.Name = "btnUsporedi";
            this.btnUsporedi.Size = new System.Drawing.Size(75, 23);
            this.btnUsporedi.TabIndex = 3;
            this.btnUsporedi.Text = "Find next";
            this.btnUsporedi.UseVisualStyleBackColor = true;
            this.btnUsporedi.Click += new System.EventHandler(this.btnUsporedi_Click);
            // 
            // textBox1
            // 
            this.textBox1.Enabled = false;
            this.textBox1.Location = new System.Drawing.Point(269, 49);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(179, 199);
            this.textBox1.TabIndex = 4;
            // 
            // textBox2
            // 
            this.textBox2.Enabled = false;
            this.textBox2.Location = new System.Drawing.Point(750, 49);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(179, 199);
            this.textBox2.TabIndex = 5;
            // 
            // MDBpathLabel
            // 
            this.MDBpathLabel.AutoSize = true;
            this.MDBpathLabel.Location = new System.Drawing.Point(131, 364);
            this.MDBpathLabel.Name = "MDBpathLabel";
            this.MDBpathLabel.Size = new System.Drawing.Size(0, 13);
            this.MDBpathLabel.TabIndex = 6;
            this.MDBpathLabel.DoubleClick += new System.EventHandler(this.MDBpathLabel_DoubleClick);
            // 
            // OutputTextbox
            // 
            this.OutputTextbox.Location = new System.Drawing.Point(506, 280);
            this.OutputTextbox.Multiline = true;
            this.OutputTextbox.Name = "OutputTextbox";
            this.OutputTextbox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.OutputTextbox.Size = new System.Drawing.Size(455, 199);
            this.OutputTextbox.TabIndex = 7;
            this.OutputTextbox.TextChanged += new System.EventHandler(this.OutputTextbox_TextChanged);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(209, 280);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(98, 23);
            this.btnAdd.TabIndex = 8;
            this.btnAdd.Text = "Add Remaining";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // StatusLabel
            // 
            this.StatusLabel.AutoSize = true;
            this.StatusLabel.Location = new System.Drawing.Point(77, 325);
            this.StatusLabel.Name = "StatusLabel";
            this.StatusLabel.Size = new System.Drawing.Size(0, 13);
            this.StatusLabel.TabIndex = 9;
            // 
            // tp_btn
            // 
            this.tp_btn.Location = new System.Drawing.Point(313, 280);
            this.tp_btn.Name = "tp_btn";
            this.tp_btn.Size = new System.Drawing.Size(124, 23);
            this.tp_btn.TabIndex = 10;
            this.tp_btn.Text = "Make custom (tp_)";
            this.tp_btn.UseVisualStyleBackColor = true;
            this.tp_btn.Click += new System.EventHandler(this.tp_btn_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.label1.Location = new System.Drawing.Point(26, 323);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(45, 15);
            this.label1.TabIndex = 11;
            this.label1.Text = "Done:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Monaco", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(22, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 16);
            this.label2.TabIndex = 12;
            this.label2.Text = "MDB";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Monaco", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(503, 22);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 16);
            this.label3.TabIndex = 13;
            this.label3.Text = "EXCEL";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.label4.Location = new System.Drawing.Point(26, 362);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(88, 15);
            this.label4.TabIndex = 14;
            this.label4.Text = "Path to mdb:";
            // 
            // btnAddFamily
            // 
            this.btnAddFamily.Location = new System.Drawing.Point(29, 404);
            this.btnAddFamily.Name = "btnAddFamily";
            this.btnAddFamily.Size = new System.Drawing.Size(119, 24);
            this.btnAddFamily.TabIndex = 15;
            this.btnAddFamily.Text = "Group PartNumbers";
            this.btnAddFamily.UseVisualStyleBackColor = true;
            this.btnAddFamily.Click += new System.EventHandler(this.btnAddFamily_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.ClientSize = new System.Drawing.Size(1000, 511);
            this.Controls.Add(this.btnAddFamily);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tp_btn);
            this.Controls.Add(this.StatusLabel);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.OutputTextbox);
            this.Controls.Add(this.MDBpathLabel);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btnUsporedi);
            this.Controls.Add(this.listBox2MDB);
            this.Controls.Add(this.btnDodaj);
            this.Controls.Add(this.listBoxExcel);
            this.Name = "Form1";
            this.Text = "Data mapping";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox listBoxExcel;
        private System.Windows.Forms.Button btnDodaj;
        private System.Windows.Forms.ListBox listBox2MDB;
        private System.Windows.Forms.Button btnUsporedi;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label MDBpathLabel;
        private System.Windows.Forms.TextBox OutputTextbox;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Label StatusLabel;
        private System.Windows.Forms.Button tp_btn;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnAddFamily;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
    }
}

