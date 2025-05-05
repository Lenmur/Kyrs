using System;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;


namespace Курсовая_Работа
{
    class Results : Participants
    {
        public double Digital_competencies;
        public double Digital_consumption;
        public double Digital_security;
        public double Digital_literacy;
        public Results(double dc, double ds, double ddc, double dl, int no, string fio1, int age1, string gender1) : base(fio1,age1,gender1,no)
        {
            Digital_competencies = dc;
            Digital_security = ds;
            Digital_consumption = ddc;
            Digital_literacy = dl;
        }

        public void Processing_of_the_results (Label labelVopros, Chart chartRes, ProgressBar progressBar1)
        {
            double result1 = (Digital_competencies / 4) * 100;
            double result2 = (Digital_security / 4) * 100;
            double result3 = (Digital_consumption / 4) * 100;
            double result4 = Math.Round(((result1 + result2 + result3) / 300) * 100);
            Digital_competencies = result1;
            Digital_security = result2;
            Digital_consumption = result3;
            Digital_literacy = result4;
            chartRes.Visible = true;
            chartRes.Series[0].Points.AddY(Digital_competencies);
            chartRes.Series[0].Points[0].LegendText = "Цифровая компетенция";
            chartRes.Series[0].Points[0].Label = Digital_competencies + "%";
            chartRes.Series[0].Points.AddY(Digital_security);
            chartRes.Series[0].Points[1].LegendText = "Цифровая безопасность";
            chartRes.Series[0].Points[1].Label = Digital_security + "%";
            chartRes.Series[0].Points.AddY(Digital_consumption);
            chartRes.Series[0].Points[2].LegendText = "Цифровое потребление";
            chartRes.Series[0].Points[2].Label = Digital_consumption + "%";
            chartRes.Series[0].IsValueShownAsLabel = true;
            labelVopros.Location = progressBar1.Location;
            if (Digital_literacy == 100)
            {
                labelVopros.Text = "Ваш Индекс Цифровой Грамотности - " + Digital_literacy + "%" + "\nВы просто Мастер технологий!";
            }
            else if (Digital_literacy >= 80 && Digital_literacy < 100)
            {
                labelVopros.Text = "Ваш Индекс Цифровой Грамотности - " + Digital_literacy + "%" + "\nВы просто Мастер технологий! \nЗнать все невозможно, но почти все можно узнать и понять.\nВсе, что вам нужно, это немного времени и давний хороший приятель Google.";
            }
            else if (Digital_literacy >= 50 && Digital_literacy < 80)
            {
                labelVopros.Text = "Ваш Индекс Цифровой Грамотности - " + Digital_literacy + "%" + "\nВы Герой нашего времени! \nПоздравляем, вы являетесь частью 21 - го века! \nВы конечно не профи, но точно знаете больше, чем достаточно.";
            }
            else if (Digital_literacy < 50)
            {
                labelVopros.Text = "Ваш Индекс Цифровой Грамотности - " + Digital_literacy + "%" + "\nВы очень плохо подкованы в сфере технологий и безопасности.";
            }
            
        }


       
    }
}
