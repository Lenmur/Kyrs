using Excel = Microsoft.Office.Interop.Excel;


namespace Курсовая_Работа
{
    class Task : Questionnaire
    {
       public int The_number_of_the_question { get; set; }
       public string Answer { get; set; }
        public Task(int no, int num, string answer ) : base( no)
        {
            The_number_of_the_question = num;
            Answer = answer;
        }
        public void Save_the_answer()
        {
            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook ObjWorkBook = ObjWorkExcel.ActiveWorkbook;
            Excel.Worksheet ObjWorkSheet2 = (Excel.Worksheet)ObjWorkExcel.Sheets[2];
            ObjWorkExcel.DisplayAlerts = false;
            ObjWorkSheet2.Cells[Number_Opros + 1,The_number_of_the_question+3].Value = Answer;
            ObjWorkBook.Save();
        }
    }
}
