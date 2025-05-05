using System;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using System.Windows.Forms;
using System.Diagnostics;

namespace Курсовая_Работа
{
    class OnPrint : Results
    {
        public OnPrint(double dc1, double ds1, double ddc1, double dl1, int no1, string fio1, int age1, string gender1, string data1) : base(dc1, ds1, ddc1, dl1, no1, fio1, age1, gender1) {}

        public void PrintWord()
        {
            string[] data = new string[4];
            var path = System.IO.Path.GetFullPath(@"Сертификат участника (Шаблон).docx");
            Application word = new Application();
            Document docW = word.Documents.Open(path);
            Bookmarks wBookmarks = docW.Bookmarks;
            Range wRange;

            data[0] = DateTime.Now.ToString("dd MMMM yyyy");
            data[1] = FIO;
            data[2] = Number_Opros.ToString();
            double result1 = (Digital_competencies / 4) * 100;
            double result2 = (Digital_security / 4) * 100;
            double result3 = (Digital_consumption / 4) * 100;
            double result4 = Math.Round(((result1 + result2 + result3) / 300) * 100);
            data[3] = result4.ToString();

            int i = 0;
            foreach (Bookmark mark in wBookmarks)
            {
                wRange = mark.Range;
                wRange.Text = data[i];
                i++;
            }
            var pathPDF = System.IO.Path.GetFullPath(@"Сертификаты PDF\\Сертификат PDF - " + data[2] + ".pdf");
            docW.ExportAsFixedFormat(pathPDF, WdExportFormat.wdExportFormatPDF);
            MessageBox.Show("Сертификат сформирован!\nВнимание! Включите принтер", "Успешно");
            word.Quit(false);

            try
            {
                Process p = new Process();
                p.StartInfo = new ProcessStartInfo()
                {
                    CreateNoWindow = true,
                    Verb = "print",
                    FileName = pathPDF,
                };
                p.Start();
            }
            catch
            {
                MessageBox.Show("Сертификат не распечатался ", "Ошибка");
                return;
            }
        }
    }
}
