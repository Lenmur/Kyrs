using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace Курсовая_Работа
{
    class InExcel : Results
    {
        public string data { get; set; }
        public InExcel(double dc1, double ds1, double ddc1, double dl1, int no1, string fio1, int age1, string gender1, string data1) : base(dc1, ds1, ddc1, dl1, no1, fio1, age1, gender1)
        {
            data = data1;
        }
        public void Import_to_Excel()
        {
            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook ObjWorkBook = ObjWorkExcel.ActiveWorkbook;
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkExcel.Sheets[1];
            double result1 = (Digital_competencies / 4) * 100;
            double result2 = (Digital_security / 4) * 100;
            double result3 = (Digital_consumption / 4) * 100;
            double result4 = Math.Round(((result1 + result2 + result3) / 300) * 100);
            int k = Number_Opros + 1;

            ObjWorkSheet.Cells[k, 6].Value = result1 + "%";
            ObjWorkSheet.Cells[k, 7].Value = result3 + "%";

            ObjWorkSheet.Cells[k, 8].Value = result2 + "%";
            ObjWorkSheet.Cells[k, 8].Font.Bold = false;

            ObjWorkSheet.Cells[k, 9].Value = result4 + "%";
            ObjWorkSheet.Cells[k, 9].Font.Bold = false;

            ObjWorkSheet.Cells[Number_Opros + 2, 8].Value = "Общий показатель ИЦГН";
            ObjWorkSheet.Cells[Number_Opros + 2, 8].Font.Bold = true;
            
            ObjWorkSheet.Cells[Number_Opros + 2, 9].Formula = "=SUM(I" + 2 + ":I" + k + ")";
            Thread.Sleep(500);
            ObjWorkSheet.Cells[Number_Opros + 2, 9].Value = (ObjWorkSheet.Cells[Number_Opros + 2, 9].Value)/k;
            ObjWorkSheet.Cells[Number_Opros + 2, 9].Font.Bold = true;
            ObjWorkBook.Save(); 
        }
    }
}
