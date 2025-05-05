using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Курсовая_Работа
{
    public partial class Form2 : Form
    {
        byte k = 0;
        int n = 0;
        int i = 1;
        double dc = 0;
        double ddc = 0;
        double ds = 0;
        List<Results> resultsList = new List<Results>();
        List<InExcel> inexcelList = new List<InExcel>();
        List<OnPrint> wordprintList = new List<OnPrint>();
        List<Task> taskList = new List<Task>();
        public Form2()
        {
            InitializeComponent();
            VisibleCheckFalse();
            CheckedRadioFalse();
            VisibleLabelFalse();
            VisibleComboBoxFalse();
            textBox1.Visible = false;
            buttonZT.Visible = false;
            Q1();
            k++;
            buttonPrint.Visible = false;
            chartRes.Visible = false;
            dc = 0; ddc = 0; ds = 0; ds = 0;
            n = otk_exel2();
            label1Number.Text = "Опрос - " + (n-1).ToString();
            pictureBox3.Visible = false;
        }
        public int otk_exel2()
        {
            int n = 0;
            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            ObjWorkExcel.Visible = false;
            Excel.Workbook ObjWorkBook = ObjWorkExcel.ActiveWorkbook;
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkExcel.Sheets[1];
            while (ObjWorkSheet.Cells[n + 1, 1].Value != null)
            {
                n++;
            }
            return n;
        }

        public string Path()
        {
            i++;
            if (i == 6)
            {
                i+=2;
            }
            var path = System.IO.Path.GetFullPath(@"Картинки для опроса\\V" + i + ".jpg");
            return path;
        }
        public void Q1()
        {
            //1 вопрос//
            progressBar1.PerformStep();
            CheckedRadioFalse();
            pictureBox3.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox3.BorderStyle = BorderStyle.Fixed3D;
            labelVopros.Text = "Вы не можете подключиться к интернету. Что вы будете делать? ";
            radioButtonO1.Text = "Поищу решение в панели диагности интернет-соединения и перезагружу роутер";
            radioButtonO2.Text = "Опубликую яростный твит, как только появится связь";
            radioButtonO3.Text = "Обращусь к тому кто рубит в этом фишку";
            radioButtonO4.Text = "Да кто его знает, я даже не знаю что такое интернет";
            labelNumVopros.Text = "Вопрос - 1";
        }
        public void Q2()
        {
            //2 вопрос//
            progressBar1.PerformStep();
            CheckedRadioFalse();
            pictureBox3.Visible = true;
            Bitmap image1 = new Bitmap(Path());
            pictureBox3.Image = image1;
            pictureBox3.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox3.BorderStyle = BorderStyle.Fixed3D;
            labelVopros.Text = "Для защиты ваших данных в Сети вы используете:";
            radioButtonO1.Text = "Щит и Меч Конечно!";
            radioButtonO2.Text = "Легкий пароль такой же как и имя моей собаки";
            radioButtonO3.Text = "Пароль который состоит из букв, цифр, символов и знаков";
            radioButtonO4.Text = "Что? Ничего! Кто-то хочет украсть у меня список моих покупок?";
            labelNumVopros.Text = "Вопрос - 2";
        }
        public void Q3()
        {
            //3 вопрос//
            progressBar1.PerformStep();
            buttonSLV.Enabled = false;
            Bitmap image1 = new Bitmap(Path());
            pictureBox3.Image = image1;
            pictureBox3.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox3.BorderStyle = BorderStyle.Fixed3D;
            labelVopros.Text = "Устройство компьютера, выполняющее обработку информации";
            VisibleRadioFalse();
            textBox1.Location = radioButtonO1.Location;
            textBox1.Visible = true;
            labelNumVopros.Text = "Вопрос - 3";
        }
        public void Q4()
        {
            //4 вопрос//
            progressBar1.PerformStep();
            CheckedRadioFalse();
            Bitmap image1 = new Bitmap(Path());
            pictureBox3.Image = image1;
            pictureBox3.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox3.BorderStyle = BorderStyle.Fixed3D;
            textBox1.Visible = false;
            VisibleRadioTrue();
            labelVopros.Text = "Как вы общаетесь с друзьями которые живут далеко от вас?";
            radioButtonO1.Text = "Жду пока они ко мне приедут";
            radioButtonO2.Text = "Звоню по стационарному телефону";
            radioButtonO3.Text = "По старинке: письма, голуби, марки, конверты";
            radioButtonO4.Text = "В Skype, FaceTime, WhatsApp";
            labelNumVopros.Text = "Вопрос - 4";
        }
        public void Q5()
        {
            //5 вопрос//
            progressBar1.PerformStep();
            CheckedRadioFalse();
            Bitmap image1 = new Bitmap(Path());
            pictureBox3.Image = image1;
            pictureBox3.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox3.BorderStyle = BorderStyle.Fixed3D;
            labelVopros.Text = "Для того, чтобы компьютер заразился вирусом, необходимо";
            radioButtonO1.Text = " Хотя бы один раз выполнять программу, содержащую вирус.";
            radioButtonO2.Text = " Переписать на дискету информацию с компьютера.";
            radioButtonO3.Text = " Запустить программу Drweb.";
            radioButtonO4.Text = " Перезагрузить компьютер";
            labelNumVopros.Text = "Вопрос - 5";
        }
        public void Q6()
        {
            //6 вопрос//
            progressBar1.PerformStep();
            labelVopros.Text = "Устройство ввода информации\n(выберите несколько вариантов ответов)";
            pictureBox3.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox3.BorderStyle = BorderStyle.Fixed3D;
            labelVopros.Location = new Point(labelVopros.Location.X, labelVopros.Location.Y - 20);
            VisibleRadioFalse();
            VisibleCheckTrue();
            checkBoxO1.Location = radioButtonO1.Location;
            checkBoxO2.Location = radioButtonO2.Location;
            checkBoxO3.Location = radioButtonO3.Location;
            checkBoxO4.Location = radioButtonO4.Location;
            CheckedCheckFalse();
            buttonSLV.Enabled = false;
            checkBoxO1.Text = " Компьютерная мышь";
            checkBoxO2.Text = " Монитор";
            checkBoxO3.Text = " Клавиатура";
            checkBoxO4.Text = " Микрофон";
            checkBoxO5.Text = " Принтер";
            labelNumVopros.Text = "Вопрос - 6";
        }
        
        public void Q7()
        {
            //7 вопрос//
            pictureBox3.Visible = false;
            pictureBox3.Visible = false;
            progressBar1.PerformStep();
            labelVopros.Location = new Point(labelVopros.Location.X, labelVopros.Location.Y + 20);
            VisibleCheckFalse();
            CheckedRadioFalse();
            radioButtonO3.Visible = true;
            radioButtonO2.Visible = true;
            labelVopros.Text = "Об обслуживании своего ПК, вы вспоминаете только тогда, когда возникают проблема";
            radioButtonO2.Text = "Да";
            radioButtonO3.Text = "Нет";
            radioButtonO1.Visible = false;
            radioButtonO4.Visible = false;
            labelNumVopros.Text = "Вопрос - 7";
        }
        public void Q8()
        {
            //8 вопрос//
            progressBar1.PerformStep();
            CheckedRadioFalse();
            VisibleRadioTrue();
            pictureBox3.Visible = true;
            Bitmap image1 = new Bitmap(Path());
            pictureBox3.Image = image1;
            pictureBox3.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox3.BorderStyle = BorderStyle.Fixed3D;
            labelVopros.Text = "Какой пароль является самым надежным?";
            radioButtonO1.Text = " А1982 ";
            radioButtonO2.Text = " A19n!n82A# ";
            radioButtonO3.Text = " Anna_1982 ";
            radioButtonO4.Text = " A1n9n8a2 ";
            labelNumVopros.Text = "Вопрос - 8";
        }
        public void Q9()
        {
            //9 вопрос//
            progressBar1.PerformStep();
            pictureBox3.Visible = false;
            labelVopros.Text = "Сопоставьте программы с их функционалом";
            VisibleRadioFalse();
            buttonSLV.Enabled = false;
            labelO1.Location = radioButtonO1.Location;
            labelO2.Location = radioButtonO2.Location;
            labelO3.Location = radioButtonO3.Location;
            labelO4.Location = radioButtonO4.Location;
            labelO5.Location = checkBoxO5.Location;
            VisibleLabelTrue();
            labelO1.Text = "1) WinRAR";
            labelO2.Text = "2) Google";
            labelO3.Text = "3) Adobe Acrobat Reader";
            labelO4.Text = "4) OneDrive";
            labelO5.Text = "5) Zoom";
            label1V.Text = "Облачное хранилище";
            label4V.Text = "Приложение для общения и для создания масштабных конференций";
            label3V.Text = "Архиватор файлов";
            label2V.Text = "Поисковой Web-браузер";
            label5V.Text = "Программа для просмотра электронных публикаций в формате PDF";
            VisibleComboBoxTrue();
            labelNumVopros.Text = "Вопрос - 9";
        }
        public void Q10()
        {
            //10 вопрос//
            pictureBox3.Visible = false;
            VisibleComboBoxFalse();
            VisibleLabelFalse();
            progressBar1.PerformStep();
            buttonSLV.Enabled = false;
            CheckedRadioFalse();
            VisibleRadioTrue();
            labelVopros.Text = "Вы решились на покупку нового компьютера.Что для вас превыше всего? ";
            radioButtonO3.Visible = true;
            radioButtonO4.Visible = true;
            radioButtonO2.Visible = true;
            radioButtonO1.Visible = true;
            radioButtonO1.Text = "Возможность настройки и апгрейта";
            radioButtonO2.Text = "Надежность бренда как Asus, Hp, Dell, Apple";
            radioButtonO3.Text = "Сочетание цены и качества";
            radioButtonO4.Text = "Спасибо, но я использую крутую печатную машинку";
            labelNumVopros.Text = "Вопрос - 10";
        }
        public void Q11()
        {
            //11 вопрос//
            CheckedRadioFalse();
            progressBar1.PerformStep();
            pictureBox3.Visible = false;
            VisibleRadioTrue();
            labelVopros.Text = "Что такое фишинг?";
            radioButtonO1.Text = "Создание бесплатных программ";
            radioButtonO2.Text = "Вид мошенничества с целью получения доступа к конфиденциальным данным пользователей — логинам и паролям";
            radioButtonO3.Text = "Бесплатное антивирусное приложение для разблокировки компьютера";
            radioButtonO4.Text = "Переписка от чужого лица с целью вымогательства денежных средств";
            labelNumVopros.Text = "Вопрос - 11";
        }
        public void Q12()
        {
            //12 вопрос//
            textBox1.Clear();
            pictureBox3.Visible = false;
            buttonSLV.Visible = false;
            buttonZT.Enabled = false;
            progressBar1.PerformStep();
            CheckedRadioFalse();
            labelVopros.Text = "Копирование выделенного объекта производится при нажатой клавише";
            VisibleRadioFalse();
            textBox1.Location = radioButtonO1.Location;
            textBox1.Visible = true;
            labelNumVopros.Text = "Вопрос - 12";
            
        }

        public void VisibleRadioTrue()
        {
            radioButtonO1.Visible = true;
            radioButtonO2.Visible = true;
            radioButtonO3.Visible = true;
            radioButtonO4.Visible = true;
        }
        public void VisibleRadioFalse()
        {
            radioButtonO1.Visible = false;
            radioButtonO2.Visible = false;
            radioButtonO3.Visible = false;
            radioButtonO4.Visible = false;
        }
        public void VisibleCheckFalse()
        {
            checkBoxO1.Visible = false;
            checkBoxO2.Visible = false;
            checkBoxO3.Visible = false;
            checkBoxO4.Visible = false;
            checkBoxO5.Visible = false;
        }
        public void VisibleCheckTrue()
        {
            checkBoxO1.Visible = true;
            checkBoxO2.Visible = true;
            checkBoxO3.Visible = true;
            checkBoxO4.Visible = true;
            checkBoxO5.Visible = true;
        }
        public void CheckedCheckFalse()
        {
            checkBoxO1.Checked = false;
            checkBoxO2.Checked = false;
            checkBoxO3.Checked = false;
            checkBoxO4.Checked = false;
            checkBoxO5.Checked = false;
        }
        public void CheckedRadioFalse()
        {
            radioButtonO1.Checked = false;
            radioButtonO2.Checked = false;
            radioButtonO3.Checked = false;
            radioButtonO4.Checked = false;
        }
        public void VisibleLabelFalse()
        {
            labelO1.Visible = false;
            labelO2.Visible = false;
            labelO3.Visible = false;
            labelO4.Visible = false;
            labelO5.Visible = false;
            label1V.Visible = false;
            label2V.Visible = false;
            label3V.Visible = false;
            label4V.Visible = false;
            label5V.Visible = false;
        }
        public void VisibleLabelTrue()
        {
            labelO1.Visible = true;
            labelO2.Visible = true;
            labelO3.Visible = true;
            labelO4.Visible = true;
            labelO5.Visible = true;
            label1V.Visible = true;
            label2V.Visible = true;
            label3V.Visible = true;
            label4V.Visible = true;
            label5V.Visible = true;
        }
        public void VisibleComboBoxFalse()
        {
            comboBox1.Visible = false;
            comboBox2.Visible = false;
            comboBox3.Visible = false;
            comboBox4.Visible = false;
            comboBox5.Visible = false;
        }
        public void VisibleComboBoxTrue()
        {
            comboBox1.Visible = true;
            comboBox2.Visible = true;
            comboBox3.Visible = true;
            comboBox4.Visible = true;
            comboBox5.Visible = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            k++;
            if (k == 2)
            {
                //1 вопрос//
                if (radioButtonO1.Checked == true)
                {
                    ddc += 1;
                    taskList.Add(new Task(n - 1, 1, "+"));
                    taskList[taskList.Count - 1].Save_the_answer();
                }
                else
                {
                    taskList.Add(new Task(n - 1, 1, "-"));
                    taskList[taskList.Count - 1].Save_the_answer();
                }
                Q2(); 
            }
            if (k == 3)
            {
                //2 вопрос//
                if (radioButtonO3.Checked == true)
                {
                    ds += 1;
                    taskList.Add(new Task(n - 1, 2, "+"));
                    taskList[taskList.Count - 1].Save_the_answer();
                }
                else
                {
                    taskList.Add(new Task(n - 1, 2, "-"));
                    taskList[taskList.Count - 1].Save_the_answer();
                }
                Q3();
            }
            if (k == 4)
            {
                //3 вопрос//
                string otvet = textBox1.Text.ToLower();
                otvet = otvet.Replace(" ", string.Empty);
                if (otvet == "процессор" || otvet == "processor" || otvet == "процесор" || otvet == "процесссор")
                {
                    dc += 1;
                    taskList.Add(new Task(n - 1, 3, "+"));
                    taskList[taskList.Count - 1].Save_the_answer();
                }
                else
                {
                    taskList.Add(new Task(n - 1, 3, "-"));
                    taskList[taskList.Count - 1].Save_the_answer();
                }
                Q4();
            }
            if (k == 5)
            {
                //4 вопрос//
                if (radioButtonO4.Checked == true)
                {
                    ddc += 1;
                    taskList.Add(new Task(n - 1, 4, "+"));
                    taskList[taskList.Count - 1].Save_the_answer();
                }
                else
                {
                    taskList.Add(new Task(n - 1, 4, "-"));
                    taskList[taskList.Count - 1].Save_the_answer();
                }

                Q5();
            }
            if (k == 6)
            {
                //5 вопрос//
                Q6();
                if (radioButtonO1.Checked == true)
                {
                    ds += 1;
                    taskList.Add(new Task(n - 1, 5, "+"));
                    taskList[taskList.Count - 1].Save_the_answer();
                }
                else
                {
                    taskList.Add(new Task(n - 1, 5, "-"));
                    taskList[taskList.Count - 1].Save_the_answer();
                }
            }
            if (k == 7)
            {
                //6 вопрос//
                if (checkBoxO1.Checked == true && checkBoxO3.Checked == true && checkBoxO4.Checked == true && checkBoxO2.Checked == false && checkBoxO5.Checked == false)
                {
                    taskList.Add(new Task(n - 1, 6, "+"));
                    taskList[taskList.Count - 1].Save_the_answer();
                    dc += 1;
                }
                else if (checkBoxO1.Checked == true && checkBoxO3.Checked == true && checkBoxO4.Checked == true && (checkBoxO2.Checked == true || checkBoxO5.Checked == true))
                {
                    taskList.Add(new Task(n - 1, 6, "+/-"));
                    taskList[taskList.Count - 1].Save_the_answer();
                    dc += 0.8;
                }
                else if ((checkBoxO1.Checked == true && checkBoxO3.Checked == true && checkBoxO2.Checked == false && checkBoxO5.Checked == false) || (checkBoxO4.Checked == true && checkBoxO1.Checked == true && checkBoxO2.Checked == false && checkBoxO5.Checked == false) || (checkBoxO3.Checked == true && checkBoxO4.Checked == true && checkBoxO2.Checked == false && checkBoxO5.Checked == false))
                {
                    taskList.Add(new Task(n - 1, 6, "+/-"));
                    taskList[taskList.Count - 1].Save_the_answer();
                    dc +=0.67;
                }
                else if ((checkBoxO1.Checked == true && checkBoxO3.Checked == true && (checkBoxO2.Checked == true || checkBoxO5.Checked == true)) || (checkBoxO4.Checked == true && checkBoxO1.Checked == true && (checkBoxO2.Checked == true || checkBoxO5.Checked == true)) || checkBoxO3.Checked == true && checkBoxO4.Checked == true && (checkBoxO2.Checked == true || checkBoxO5.Checked == true))
                {
                    taskList.Add(new Task(n - 1, 6, "+/-"));
                    taskList[taskList.Count - 1].Save_the_answer();
                    dc += 0.6;
                }
                else if ((checkBoxO1.Checked == true && (checkBoxO2.Checked == false && checkBoxO5.Checked == false)) || (checkBoxO3.Checked == true && (checkBoxO2.Checked == false && checkBoxO5.Checked == false)) || (checkBoxO4.Checked == true && (checkBoxO2.Checked == false && checkBoxO5.Checked == false)))
                {
                    taskList.Add(new Task(n - 1, 6, "+/-"));
                    taskList[taskList.Count - 1].Save_the_answer();
                    dc += 0.33;
                }
                else if ((checkBoxO1.Checked == true && (checkBoxO2.Checked == true || checkBoxO5.Checked == true)) || (checkBoxO3.Checked == true && (checkBoxO2.Checked == true || checkBoxO5.Checked == true)) || (checkBoxO4.Checked == true && (checkBoxO2.Checked == true || checkBoxO5.Checked == true)))
                {
                    taskList.Add(new Task(n - 1, 6, "+/-"));
                    taskList[taskList.Count - 1].Save_the_answer();
                    dc += 0.25;
                }
                Q7();
            }
            if (k == 8)
            {
                //7 вопрос//
                if (radioButtonO3.Checked == true)
                {
                    ddc += 1;
                    taskList.Add(new Task(n - 1, 7, "+"));
                    taskList[taskList.Count - 1].Save_the_answer();
                }
                else
                {
                    taskList.Add(new Task(n - 1, 7, "-"));
                    taskList[taskList.Count - 1].Save_the_answer();
                }
                Q8();
            }
            if (k == 9)
            {
                //8 вопрос//
                if (radioButtonO2.Checked == true)
                {
                    ds += 1;
                    taskList.Add(new Task(n - 1, 8, "+"));
                    taskList[taskList.Count - 1].Save_the_answer();
                }
                else
                {
                    taskList.Add(new Task(n - 1, 8, "-"));
                    taskList[taskList.Count - 1].Save_the_answer();
                }
                Q9();
            }
            if (k == 10)
            {
                //9 вопрос//
                if (comboBox1.Text == "3")
                {
                    dc += 0.2;
                }
                if (comboBox2.Text == "5")
                {
                    dc += 0.2;
                }
                if (comboBox3.Text == "1")
                {
                    dc += 0.2;
                }
                if (comboBox4.Text == "2")
                {
                    dc += 0.2;
                }
                if (comboBox5.Text == "4")
                {
                    dc += 0.2;
                }
                string p = comboBox1.Text + ";" + comboBox2.Text + ";" + comboBox3.Text + ";" + comboBox4.Text + ";" + comboBox5.Text + ";";
                taskList.Add(new Task(n - 1, 9, p));
                taskList[taskList.Count - 1].Save_the_answer();
                Q10();
            }
            if (k == 11)
            {
                //10 вопрос//
                if (radioButtonO1.Checked == true)
                {
                    ddc += 1;
                    taskList.Add(new Task(n - 1, 10, "+"));
                    taskList[taskList.Count - 1].Save_the_answer();
                }
                else
                {
                    taskList.Add(new Task(n - 1, 10, "-"));
                    taskList[taskList.Count - 1].Save_the_answer();
                }
                Q11();
            }
            if (k == 12)
            {
                //11 вопрос//
                if(radioButtonO2.Checked == true)
                {
                    ds += 1;
                    taskList.Add(new Task(n - 1, 11, "+"));
                    taskList[taskList.Count - 1].Save_the_answer();
                }
                else
                {
                    taskList.Add(new Task(n - 1, 11, "-"));
                    taskList[taskList.Count - 1].Save_the_answer();
                }
                Q12();
                buttonSLV.Visible = false;
                buttonZT.Location = buttonSLV.Location;
                buttonZT.Visible = true;
            }
        }



        private void radioButtonO1_CheckedChanged(object sender, EventArgs e)
        {
            if(radioButtonO1.Checked ==true && radioButtonO1.Visible == true)
            {
                buttonSLV.Enabled = true;
            }
            else
            {
                buttonSLV.Enabled = false;
            }
        }

        private void radioButtonO2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonO2.Checked == true && radioButtonO2.Visible == true)
            {
                buttonSLV.Enabled = true;
            }
            else
            {
                buttonSLV.Enabled = false;
            }
        }

        private void radioButtonO3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonO3.Checked == true && radioButtonO3.Visible == true)
            {
                buttonSLV.Enabled = true;
            }
            else
            {
                buttonSLV.Enabled = false;
            }
        }

        private void radioButtonO4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonO4.Checked == true && radioButtonO4.Visible == true)
            {
                buttonSLV.Enabled = true;
            }
            else
            {
                buttonSLV.Enabled = false;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.TextLength < 3 )
            {
                buttonSLV.Enabled = false;
                buttonZT.Enabled = false;
            }
            else
            {
                buttonSLV.Enabled = true;
                buttonZT.Enabled = true;
            }
        }

        private void checkBoxO1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxO1.Checked == true && checkBoxO1.Visible == true)
            {
                buttonSLV.Enabled = true;
            }
            else if (checkBoxO1.Checked == false && checkBoxO1.Visible == true)
            {
                buttonSLV.Enabled = false;
            }
        }

        private void checkBoxO2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxO2.Checked == true && checkBoxO2.Visible == true)
            {
                buttonSLV.Enabled = true;
            }
            else if (checkBoxO2.Checked == false && checkBoxO2.Visible == true)
            {
                buttonSLV.Enabled = false;
            }
        }

        private void checkBoxO3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxO3.Checked == true && checkBoxO3.Visible == true)
            {
                buttonSLV.Enabled = true;
            }
            else if(checkBoxO3.Checked == false && checkBoxO3.Visible == true)
            {
                buttonSLV.Enabled = false;
            }
        }

        private void checkBoxO4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxO4.Checked == true && checkBoxO4.Visible == true)
            {
                buttonSLV.Enabled = true;
            }
            else if (checkBoxO4.Checked == false && checkBoxO4.Visible == true)
            {
                buttonSLV.Enabled = false;
            }
        }

        private void checkBoxO5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxO5.Checked == true && checkBoxO5.Visible == true)
            {
                buttonSLV.Enabled = true;
            }
            else if (checkBoxO5.Checked == false && checkBoxO5.Visible == true)
            {
                buttonSLV.Enabled = false;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox1.Text != comboBox2.Text && comboBox1.Text != comboBox3.Text && comboBox1.Text != comboBox4.Text && comboBox1.Text != comboBox5.Text)
            {
                comboBox1.ForeColor = Color.Black;
            }
            else
            {
                comboBox1.ForeColor = Color.Red;
                buttonSLV.Enabled = false;
            }
            if ((comboBox5.Text.Length != 0 && comboBox4.Text.Length != 0 && comboBox3.Text.Length != 0 && comboBox2.Text.Length != 0 && comboBox1.Text.Length != 0) && (comboBox5.ForeColor != Color.Red && comboBox4.ForeColor != Color.Red && comboBox3.ForeColor != Color.Red && comboBox2.ForeColor != Color.Red && comboBox1.ForeColor != Color.Red))
            {
                buttonSLV.Enabled = true;
            }
            comboBox1.SelectionLength = 0;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text != comboBox1.Text && comboBox2.Text != comboBox3.Text && comboBox2.Text != comboBox4.Text && comboBox2.Text != comboBox5.Text)
            {
                comboBox2.ForeColor = Color.Black;
            }
            else
            {
                comboBox2.ForeColor = Color.Red;
                buttonSLV.Enabled = false;
            }
            if ((comboBox5.Text.Length != 0 && comboBox4.Text.Length != 0 && comboBox3.Text.Length != 0 && comboBox2.Text.Length != 0 && comboBox1.Text.Length != 0) && (comboBox5.ForeColor != Color.Red && comboBox4.ForeColor != Color.Red && comboBox3.ForeColor != Color.Red && comboBox2.ForeColor != Color.Red && comboBox1.ForeColor != Color.Red))
            {
                buttonSLV.Enabled = true;
            }

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.Text != comboBox1.Text && comboBox3.Text != comboBox2.Text && comboBox3.Text != comboBox4.Text && comboBox3.Text != comboBox5.Text)
            {
                comboBox3.ForeColor = Color.Black;
            }
            else
            {
                comboBox3.ForeColor = Color.Red;
                buttonSLV.Enabled = false;
            }
            if ((comboBox5.Text.Length != 0 && comboBox4.Text.Length != 0 && comboBox3.Text.Length != 0 && comboBox2.Text.Length != 0 && comboBox1.Text.Length != 0) && (comboBox5.ForeColor != Color.Red && comboBox4.ForeColor != Color.Red && comboBox3.ForeColor != Color.Red && comboBox2.ForeColor != Color.Red && comboBox1.ForeColor != Color.Red))
            {
                buttonSLV.Enabled = true;
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox4.Text != comboBox1.Text && comboBox4.Text != comboBox2.Text && comboBox4.Text != comboBox3.Text && comboBox4.Text != comboBox5.Text)
            {
                comboBox4.ForeColor = Color.Black;
            }
            else
            {
                comboBox4.ForeColor = Color.Red;
                buttonSLV.Enabled = false;
            }
            if ((comboBox5.Text.Length != 0 && comboBox4.Text.Length != 0 && comboBox3.Text.Length != 0 && comboBox2.Text.Length != 0 && comboBox1.Text.Length != 0) && (comboBox5.ForeColor != Color.Red && comboBox4.ForeColor != Color.Red && comboBox3.ForeColor != Color.Red && comboBox2.ForeColor != Color.Red && comboBox1.ForeColor != Color.Red))
            {
                buttonSLV.Enabled = true;
            }
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox5.Text != comboBox1.Text && comboBox5.Text != comboBox2.Text && comboBox5.Text != comboBox3.Text && comboBox5.Text != comboBox4.Text)
            {
                comboBox5.ForeColor = Color.Black;
            }
            else
            {
                comboBox5.ForeColor = Color.Red;
                buttonSLV.Enabled = false;
            }
            if ((comboBox5.Text.Length != 0 && comboBox4.Text.Length != 0 && comboBox3.Text.Length != 0 && comboBox2.Text.Length != 0 && comboBox1.Text.Length != 0) && (comboBox5.ForeColor != Color.Red && comboBox4.ForeColor != Color.Red && comboBox3.ForeColor != Color.Red && comboBox2.ForeColor != Color.Red && comboBox1.ForeColor != Color.Red))
            {
                buttonSLV.Enabled = true;
            }
        }

        private void buttonZT_Click(object sender, EventArgs e)
        {
            //12 вопрос//
            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkExcel.Sheets[1];
            string otvet = textBox1.Text.ToLower();
            otvet = otvet.Replace(" ", string.Empty);
            if (otvet == "ctrl" || otvet == "стрл" || otvet == "контрл" || otvet == "ктрл")
            {
                dc += 1;
                taskList.Add(new Task(n - 1, 12, "+"));
                taskList[taskList.Count - 1].Save_the_answer();
            }
            else
            {
                taskList.Add(new Task(n - 1, 12, "-"));
                taskList[taskList.Count - 1].Save_the_answer();
            }
            AllVisibleFalse();
            pictureBox1.Visible = false;
            // конец 12 вопроса // 
            string z1 = ObjWorkSheet.Cells[n, 3].Value.ToString();
            int z2 = Convert.ToInt32(ObjWorkSheet.Cells[n, 4].Value);
            string z3 = ObjWorkSheet.Cells[n, 5].Value.ToString();
            resultsList.Add(new Results(dc,ds,ddc,0, n-1,z1 ,z2 ,z3));
            resultsList[resultsList.Count - 1].Processing_of_the_results(labelVopros, chartRes,progressBar1);
            inexcelList.Add(new InExcel(dc, ds, ddc, 0, n - 1, z1, z2, z3, DateTime.Now.ToString()));
            inexcelList[inexcelList.Count - 1].Import_to_Excel();
            Thread.Sleep(500);
            MessageBox.Show("Вы можете распечатать сертификат участника","Внимание");
            buttonPrint.Visible = true;
        }
        public void AllVisibleFalse()
        {
            buttonSLV.Visible = false;
            textBox1.Visible = false;
            buttonZT.Visible = false;
            progressBar1.Visible = false;
            labelNumVopros.Visible = false;
            pictureBox3.Visible = false;

        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!((e.KeyChar >= 'A' && e.KeyChar <= 'я') || e.KeyChar == 8))
            {
                e.Handled = true;
            }
        }

        private void comboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }
        private void comboBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }
        private void comboBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }
        private void comboBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }
        private void comboBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void buttonPrint_Click(object sender, EventArgs e)
        {
            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook ObjWorkBook = ObjWorkExcel.ActiveWorkbook;
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkExcel.Sheets[1];
            string z1 = ObjWorkSheet.Cells[n, 3].Value.ToString();
            int z2 = Convert.ToInt32(ObjWorkSheet.Cells[n, 4].Value);
            string z3 = ObjWorkSheet.Cells[n, 5].Value.ToString();
            wordprintList.Add(new OnPrint(dc, ds, ddc, 0, n - 1, z1, z2, z3, DateTime.Now.ToString()));
            wordprintList[0].PrintWord();
        }
    }
}
