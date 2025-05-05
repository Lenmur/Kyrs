using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;


namespace Курсовая_Работа
{
    public partial class Form1 : Form
    {
        List<Participants> participantsList = new List<Participants>();
        
        int n = 0;

        public Form1()
        {
            InitializeComponent();
            Process.Start(Path());
            Thread.Sleep(1000);
            VisibleTrue();
            ProitiTest.Enabled = false;
            тестированиеToolStripMenuItem.Enabled = false;
        }
        public string Path()
        {
            var path = System.IO.Path.GetFullPath(@"Результаты опроса.xlsx");
            return path;
        }

        public int otk_exel()
        {
            int k = 0;
            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook ObjWorkBook = ObjWorkExcel.ActiveWorkbook;
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkExcel.Sheets[1];
            while (ObjWorkSheet.Cells[k+1,1].Value != null)
            {
                k++;
            }
            return k;
        }

        private void посмотретьПредыдущиеРезультатыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.MdiChildren.Any())
            {
                return;
            }
            else
            {
                тестированиеToolStripMenuItem.BackColor = Color.Transparent;
                посмотретьПредыдущиеРезультатыToolStripMenuItem.BackColor = Color.Red;
                VisibleFalse();
                Form3 child1 = new Form3();
                child1.MdiParent = this;
                child1.Location = new System.Drawing.Point(0, 0);
                child1.Dock = DockStyle.Fill;
                child1.Width = 1158;
                child1.Height = 651;
                child1.Show();
            }
        }

        private void закрытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ProitiTest_Click(object sender, EventArgs e)
        {
            n = otk_exel();
            if (comboBox1.Text == "муж.")
            {
                participantsList.Add(new Participants(textBox1.Text, Convert.ToInt32(numericUpDown1.Value), "Мужской", n));
                participantsList[participantsList.Count - 1].Save();
            }
            if (comboBox1.Text == "жен.")
            {
                participantsList.Add(new Participants(textBox1.Text, Convert.ToInt32(numericUpDown1.Value), "Женский", n));
                participantsList[participantsList.Count - 1].Save();
            }
            тестированиеToolStripMenuItem.Enabled = true;
            тестированиеToolStripMenuItem.BackColor = Color.Red;
            VisibleFalse();
            Form2 child = new Form2();
            child.Owner = this;
            child.Show();
            child.Dock = DockStyle.Fill;
            child.MdiParent = this;
            child.Width = 1158;
            child.Height = 658;
            child.Location = new System.Drawing.Point(0, 0);
        }
        public void VisibleFalse()
        {
            labelForm.Visible = false;
            labelText1.Visible = false;
            labelText2.Visible = false;
            labelText3.Visible = false;
            LabelData.Visible = false;
            ProitiTest.Visible = false;
            ГлавнаяКартинка.Visible = false;
            label1.Visible = false;
            label2.Visible = false;
            label3.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            textBox1.Visible = false;
            textBox1.Clear();
            numericUpDown1.Visible = false;
            numericUpDown1.Value = 1;
            comboBox1.Visible = false;
            comboBox1.Text = "";
        }
        public void VisibleTrue()
        {
            textBox1.Visible = true;
            numericUpDown1.Visible = true;
            comboBox1.Visible = true;
            labelText1.Visible = true;
            labelText2.Visible = true;
            labelText3.Visible = true;
            label1.Visible = true;
            label2.Visible = true;
            label3.Visible = true;
            label4.Visible = true;
            label5.Visible = true;
            LabelData.Text = "Дата: " + DateTime.Now.ToString("dd MMMM yyyy");
            LabelData.Visible = true;
            ProitiTest.Visible = true;
            ГлавнаяКартинка.Visible = true;
            labelForm.Visible = true;
        }
        private void comboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            ObjWorkExcel.Visible = false;
            if (comboBox1.Text != "" && textBox1.Text != "" && numericUpDown1.Value != 1)
            {
                ProitiTest.Enabled = true;
            }
            else
            {
                ProitiTest.Enabled = false;
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            if ((l < 'А' || l > 'я') && l != '\b' && l != ' ')
            {
                e.Handled = true;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
           if (comboBox1.Text != "" && textBox1.Text != "" && numericUpDown1.Value != 1)
           {
                ProitiTest.Enabled = true;
           }
           else
           {
                ProitiTest.Enabled = false;
           }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text != "" && textBox1.Text != "" && numericUpDown1.Value != 1)
            {
                ProitiTest.Enabled = true;
            }
            else
            {
                ProitiTest.Enabled = false;
            }
        }


        private void наГлавнуюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int i = 0;
            if (MdiChildren.Count()>0)
            {
                do
                {
                    MdiChildren[i].Close();
                }
                while (MdiChildren.Count() > 0);
            }
            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook ObjWorkBook = ObjWorkExcel.ActiveWorkbook;
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkExcel.Sheets[1];
            Excel.Worksheet ObjWorkSheet2 = (Excel.Worksheet)ObjWorkExcel.Sheets[2];
            if (ObjWorkSheet.Cells[n + 1, 1].Text != "" && ObjWorkSheet.Cells[n + 1, 6].Text == "")
            {
                for (int j = 1; j <= 5; j++)
                {
                    ObjWorkSheet.Cells[n + 1, j].Value = "";
                }
                for (int j = 1; j <= 15; j++)
                {
                    ObjWorkSheet2.Cells[n + 1, j].Value = "";
                }
            }
            тестированиеToolStripMenuItem.BackColor = Color.Transparent;
            посмотретьПредыдущиеРезультатыToolStripMenuItem.BackColor = Color.Transparent;
            VisibleTrue();
        }



        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                Excel.Workbook ObjWorkBook = ObjWorkExcel.ActiveWorkbook;
                Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkExcel.Sheets[1];
                Excel.Worksheet ObjWorkSheet2 = (Excel.Worksheet)ObjWorkExcel.Sheets[2];
                if (ObjWorkSheet.Cells[n + 1, 1].Text != "" && ObjWorkSheet.Cells[n + 1, 6].Text == "")
                {
                    for (int i = 1; i <= 5;i++ )
                    {
                        ObjWorkSheet.Cells[n + 1, i].Value = "";
                    }
                    for (int i = 1; i <= 15; i++)
                    {
                        ObjWorkSheet2.Cells[n + 1, i].Value = "";
                    }
                }
                ObjWorkBook.Save();
                ObjWorkExcel.Quit();
            }
            catch
            {
                this.Close();
            }
        }

        private void Form1_Click(object sender, EventArgs e)
        {
            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            ObjWorkExcel.Visible = false;
        }

        private void ГлавнаяКартинка_Click(object sender, EventArgs e)
        {
            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            ObjWorkExcel.Visible = false;
        }
    }
}
