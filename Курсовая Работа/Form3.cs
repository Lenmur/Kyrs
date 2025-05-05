using System;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Курсовая_Работа
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
            dataGridView1.ColumnCount = 4;
            int k = 0;
            chartRes.Visible = false;
            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook ObjWorkBook = ObjWorkExcel.ActiveWorkbook;
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkExcel.Sheets[1];
            while (ObjWorkSheet.Cells[k + 1, 1].Value != null)
            {
                k++;
            }
            dataGridView1.RowCount = k;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.Rows[0].DefaultCellStyle.BackColor = Color.BlueViolet;
            for (int i = 0; i < k; i++)
            {
                for (int j = 1; j <= 9; j++)
                {
                    if (j == 1)
                    {
                        dataGridView1.Rows[i].Cells[0].Value = ObjWorkSheet.Cells[i+1, j].Value.ToString();
                    }
                    if (j == 2)
                    {
                        dataGridView1.Rows[i].Cells[1].Value = ObjWorkSheet.Cells[i+1, j].Value.ToString();
                    }
                    if (j == 3)
                    {
                        dataGridView1.Rows[i].Cells[2].Value = ObjWorkSheet.Cells[i+1, j].Value.ToString();
                    }
                    if (j == 9)
                    {
                        dataGridView1.Rows[i].Cells[3].Style.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Bold);
                        dataGridView1.Rows[i].Cells[3].Value = ObjWorkSheet.Cells[i + 1, j].Text;
                    }

                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook ObjWorkBook = ObjWorkExcel.ActiveWorkbook;
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkExcel.Sheets[1];
            if (dataGridView1.CurrentRow.Index != 0)
            {
                int index = dataGridView1.CurrentRow.Index;
                int numer = Convert.ToInt32(dataGridView1.Rows[index].Cells[0].Value);

                double d1 = ObjWorkSheet.Cells[numer+1, 6].Value;
                double z1 = ObjWorkSheet.Cells[numer+1, 7].Value;
                double x1 = ObjWorkSheet.Cells[numer+1, 8].Value;
                chartRes.Visible = true;
                chartRes.Series[0].Points.Clear();
                chartRes.Series[0].Points.AddY(d1*100);
                chartRes.Series[0].Points[0].LegendText = "Цифровая компетенция";
                chartRes.Series[0].Points[0].Label = (d1 * 100) + "%";
                chartRes.Series[0].Points.AddY(z1*100);
                chartRes.Series[0].Points[1].LegendText = "Цифровое потребление";
                chartRes.Series[0].Points[1].Label = (z1 * 100) + "%";
                chartRes.Series[0].Points.AddY(x1*100);
                chartRes.Series[0].Points[2].LegendText = "Цифровая безопасность";
                chartRes.Series[0].Points[2].Label = (x1 * 100) + "%";
                chartRes.Series[0].IsValueShownAsLabel = true;
            }
            else
            {
                MessageBox.Show("Выберите другую строчку","Ошибка");
            }
        }
    }
}
