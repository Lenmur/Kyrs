using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace Курсовая_Работа
{
    public class Participants : Questionnaire
    {
        public string FIO { get; set; }
        public int Age { get; set; }
        public string Gender { get; set; }

        public Participants(string fio1, int age1, string gender1, int no1) : base(no1)
        { 
            FIO = fio1; 
            Age = age1; 
            Gender = gender1; 
        }

        public void Save()
        {
            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook ObjWorkBook = ObjWorkExcel.ActiveWorkbook;
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkExcel.Sheets[1];
            Excel.Worksheet ObjWorkSheet2 = (Excel.Worksheet)ObjWorkExcel.Sheets[2];

            ObjWorkSheet.Cells[Number_Opros + 1, 1].Value = Number_Opros;
            ObjWorkSheet.Cells[Number_Opros + 1, 2].Value = DateTime.Now;
            ObjWorkSheet.Cells[Number_Opros + 1, 3].Value = FIO;
            ObjWorkSheet.Cells[Number_Opros + 1, 4].Value = Age;
            ObjWorkSheet.Cells[Number_Opros + 1, 5].Value = Gender;

            ObjWorkSheet2.Cells[Number_Opros + 1, 1].Value = Number_Opros;
            ObjWorkSheet2.Cells[Number_Opros + 1, 2].Value = DateTime.Now;
            ObjWorkSheet2.Cells[Number_Opros + 1, 3].Value = FIO;

            ObjWorkBook.Save();
        }

    }
}
