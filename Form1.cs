//Есть баг с открытием файла, если пользователь закрывает окно открытия не выбрав файл, то 
//в редакторе программа виснет, но при еспользовании exe файла выдает ошибку и все ок
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Threading;
using System.IO;
using System.Globalization;
using System.Security.Policy;
using System.Security.Cryptography;
using System.Windows.Forms.DataVisualization.Charting;
using MathNet.Numerics;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        const int n = 5001;
        public Form1()
        {
            InitializeComponent();
        }

        //Физические константы
        // eps нужно для точности вычисления при паралельном соединении насосов
        const double g = 9.80665, pi = 3.1415926535, ei = 2.7182818284, eps = 0.001;
        const double deltaT = 30;       //Максимально возможный перепад изменения температур в трубе
        const double deltanu = 0.000001;  //Максимально возможный перепад изменения температур в трубе
        //Функция рандома
        private Random rnd = new Random();


        //Блок 1 - Магистральные и подпорные насосы
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        //Число насосов
        int n_pump = 0;     //Магистральных
        //Параметры насосов
        string[] mark_pump = new string[100];       //Марка насоса
        double[] rate_pump = new double[100];       //Подача ротора
        double[] disk_pump = new double[100];       //Диаметр диска ротора
        string[] engn_pump = new string[100];       //Тип привода (электродвигателя)
        double[] efcy_pump = new double[100];       //КПД насоса
        double[] capA_pump = new double[100];       //Коэффициент аппроксимации a, м
        double[] capB_pump = new double[100];       //Коэффициент аппроксимации b, м
        double[] kz = new double[100];              //Кавитационный запас
        //Для расчета через аппроксимацию 
        double a = 0, b = 0, k = 0, k1 = 0, Q = 0;
        

        //Считывание параметров насосов
        private void button2_Click(object sender, EventArgs e)
        {
            if (numericUpDown2.Value > 100)
            {
                numericUpDown2.Value = 100;
                MessageBox.Show("Количество насосов не может превышать 100", "Внимание", MessageBoxButtons.OK);
            }

            n_pump = Convert.ToInt32(numericUpDown2.Value);
            dataGridView2.ColumnCount = n_pump;

            for (int i = 0; i < n_pump; i++)
            {
                dataGridView2.Columns[i].Name = i.ToString();
                dataGridView2.Columns[i].DefaultCellStyle.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(63)))), ((int)(((byte)(63)))), ((int)(((byte)(70)))));
                dataGridView2.Columns[i].Width = 35;
            }

            numericUpDown1.Enabled = true;
            numericUpDown1.Maximum = numericUpDown2.Value;
            comboBox1.Enabled = true;
            button3.Enabled = true;
            for (int i = 0; i < n_pump; i++)
            {
                mark_pump[i] = comboBox1.Items[0].ToString();
                rate_pump[i] = 1250.0;
                disk_pump[i] = 440.0;
                engn_pump[i] = "СТД 1250-2";
                efcy_pump[i] = 82.0;
                capA_pump[i] = 331.0;
                capB_pump[i] = 0.451 / 10000;
                kz[i] = 20;
            }
            comboBox1.SelectedItem = mark_pump[0];
            textBox5.Text = rate_pump[0].ToString();
            textBox6.Text = disk_pump[0].ToString();
            textBox7.Text = engn_pump[0];
            textBox8.Text = efcy_pump[0].ToString();
            textBox9.Text = capA_pump[0].ToString();
            textBox10.Text = capB_pump[0].ToString();
            textBox21.Text = kz[0].ToString();
        }
        //Выбор насосов
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedEmployee = (string)comboBox1.SelectedItem;
            if (selectedEmployee != "Свой вариант (расширенный)")
            {
                textBox5.Visible = true;
                textBox5.Enabled = true;

                textBox6.Visible = true;
                textBox6.Enabled = true;

                textBox7.Visible = true;
                textBox7.Enabled = true;

                textBox8.Visible = true;
                textBox8.Enabled = true;

                textBox9.Visible = true;
                textBox9.Enabled = true;

                textBox10.Visible = true;
                textBox10.Enabled = true;

                textBox21.Enabled = true;

                label11.Visible = true;
                label12.Visible = true;
                label13.Visible = true;
                label14.Visible = true;
                label15.Visible = true;
                label16.Visible = true;

                button1.Visible = false;
                button1.Enabled = false;

                dataGridView1.Enabled = false;
                dataGridView1.Visible = false;

                button3.Enabled = true;

            }
            else
                if (selectedEmployee == "Свой вариант (расширенный)")
            {
                textBox5.Visible = false;
                textBox5.Enabled = false;

                textBox6.Visible = false;
                textBox6.Enabled = false;

                textBox7.Visible = false;
                textBox7.Enabled = false;

                textBox8.Visible = false;
                textBox8.Enabled = false;

                textBox9.Visible = false;
                textBox9.Enabled = false;

                textBox10.Visible = false;
                textBox10.Enabled = false;

                textBox21.Enabled = true;

                label11.Visible = false;
                label12.Visible = false;
                label13.Visible = false;
                label14.Visible = false;
                label15.Visible = false;
                label16.Visible = false;

                button1.Visible = true;
                button1.Enabled = true;

                dataGridView1.Visible = true;
                dataGridView1.Enabled = true;

                button3.Enabled = false;
            }


            switch (selectedEmployee)
            {
                case "НМ 1250-260":
                    textBox5.Text = "1250,0";
                    textBox6.Text = "440,0";
                    textBox7.Text = "СТД 1250-2";
                    textBox8.Text = "82,0";
                    textBox9.Text = "331,0";
                    textBox10.Text = (0.0000451).ToString();
                    textBox21.Text = "20";
                    break;
                case "НМ 1250-260 (900 м3/ч)":
                    textBox5.Text = "900,0";
                    textBox6.Text = "418,0";
                    textBox7.Text = "СТД 1250-2";
                    textBox8.Text = "82,0";
                    textBox9.Text = "273,0";
                    textBox10.Text = (0.000008).ToString();
                    textBox21.Text = "20";
                    break;
                case "НМ 2500-230":
                    textBox5.Text = "2500,0";
                    textBox6.Text = "430,0";
                    textBox7.Text = "СТД 2000-2";
                    textBox8.Text = "86,0";
                    textBox9.Text = "282,0";
                    textBox10.Text = (0.00000792).ToString();
                    textBox21.Text = "32";
                    break;
                case "НМ 2500-230 (1800 м3/ч)":
                    textBox5.Text = "1800,0";
                    textBox6.Text = "405,0";
                    textBox7.Text = "СТД 2000-2";
                    textBox8.Text = "86,0";
                    textBox9.Text = "251,0";
                    textBox10.Text = (0.00000812).ToString();
                    textBox21.Text = "27";
                    break;
                case "НМ 2500-230 (1250 м3/ч)":
                    textBox5.Text = "1250,0";
                    textBox6.Text = "425,0";
                    textBox7.Text = "СТД 2000-2";
                    textBox8.Text = "86,0";
                    textBox9.Text = "245,0";
                    textBox10.Text = (0.000016).ToString();
                    textBox21.Text = "25";
                    break;
                ////////////////////////////////////////////////
                case "НМ 3600-260":
                    textBox5.Text = "3600,0";
                    textBox6.Text = "450,0";
                    textBox7.Text = "СТД 2500-2";
                    textBox8.Text = "86,0";
                    textBox9.Text = "304,0";
                    textBox10.Text = (0.00000579).ToString();
                    textBox21.Text = "38";
                    break;
                case "НМ 3600-260 (2500 м3/ч)":
                    textBox5.Text = "2500,0";
                    textBox6.Text = "430,0";
                    textBox7.Text = "СТД 2500-2";
                    textBox8.Text = "84,0";
                    textBox9.Text = "285,0";
                    textBox10.Text = (0.00000644).ToString();
                    textBox21.Text = "35";
                    break;
                case "НМ 3600-260 (1800 м3/ч)":
                    textBox5.Text = "1800,0";
                    textBox6.Text = "450,0";
                    textBox7.Text = "СТД 2500-2";
                    textBox8.Text = "82,0";
                    textBox9.Text = "273,0";
                    textBox10.Text = (0.0000125).ToString();
                    textBox21.Text = "33";
                    break;

                ////////////////////////////////////////////////
                case "НМ 5000-210":
                    textBox5.Text = "5000,0";
                    textBox6.Text = "450,0";
                    textBox7.Text = "СТД 3200-2";
                    textBox8.Text = "86,0";
                    textBox9.Text = "272,0";
                    textBox10.Text = (0.0000026).ToString();
                    textBox21.Text = "42";
                    break;
                case "НМ 5000-210 (3500 м3/ч)":
                    textBox5.Text = "3500,0";
                    textBox6.Text = "470,0";
                    textBox7.Text = "СТД 3200-2";
                    textBox8.Text = "84,0";
                    textBox9.Text = "286,0";
                    textBox10.Text = (0.00000529).ToString();
                    textBox21.Text = "31";
                    break;
                case "НМ 5000-210 (2500 м3/ч)":
                    textBox5.Text = "2500,0";
                    textBox6.Text = "480,0";
                    textBox7.Text = "СТД 3200-2";
                    textBox8.Text = "82,0";
                    textBox9.Text = "236,0";
                    textBox10.Text = (0.00000484).ToString();
                    textBox21.Text = "27";
                    break;
                ////////////////////////////////////////////////
                case "НМ 7000-210":
                    textBox5.Text = "7000,0";
                    textBox6.Text = "455,0";
                    textBox7.Text = "СТД 5000-2";
                    textBox8.Text = "86,0";
                    textBox9.Text = "299,0";
                    textBox10.Text = (0.00000194).ToString();
                    textBox21.Text = "52";
                    break;
                case "НМ 7000-210 (5000 м3/ч)":
                    textBox5.Text = "5000,0";
                    textBox6.Text = "475,0";
                    textBox7.Text = "СТД 5000-2";
                    textBox8.Text = "84,0";
                    textBox9.Text = "281,0";
                    textBox10.Text = (0.00000249).ToString();
                    textBox21.Text = "45";
                    break;
                case "НМ 7000-210 (3500 м3/ч)":
                    textBox5.Text = "3500,0";
                    textBox6.Text = "476,0";
                    textBox7.Text = "СТД 5000-2";
                    textBox8.Text = "82,0";
                    textBox9.Text = "272,0";
                    textBox10.Text = (0.00000290).ToString();
                    textBox21.Text = "50";
                    break;
                ////////////////////////////////////////////////
                case "НМ 10000-210":
                    textBox5.Text = "10000,0";
                    textBox6.Text = "495,0";
                    textBox7.Text = "СТД 6300-2";
                    textBox8.Text = "90,0";
                    textBox9.Text = "307,0";
                    textBox10.Text = (0.000000975).ToString();
                    textBox21.Text = "65";
                    break;
                case "НМ 10000-210 (7000 м3/ч)":
                    textBox5.Text = "10000,0";
                    textBox6.Text = "505,0";
                    textBox7.Text = "СТД 6300-2";
                    textBox8.Text = "84,0";
                    textBox9.Text = "305,0";
                    textBox10.Text = (0.00000208).ToString();
                    textBox21.Text = "60";
                    break;
                case "НМ 10000-210 (5000 м3/ч)":
                    textBox5.Text = "10000,0";
                    textBox6.Text = "475,0";
                    textBox7.Text = "СТД 6300-2";
                    textBox8.Text = "82,0";
                    textBox9.Text = "263,0";
                    textBox10.Text = (0.00000197).ToString();
                    textBox21.Text = "45";
                    break;
                case "Свой вариант (a-bQQ)":
                    textBox5.Text = "";
                    textBox6.Text = "";
                    textBox7.Text = "";
                    textBox8.Text = "";
                    textBox9.Text = "";
                    textBox10.Text = "";
                    textBox21.Text = "";
                    break;

                case "Свой вариант (расширенный)":
                    textBox5.Text = "";
                    textBox6.Text = "";
                    textBox7.Text = "";
                    textBox8.Text = "";
                    textBox9.Text = "";
                    textBox10.Text = "";
                    textBox21.Text = "";
                    break;
            }
        }
        //Расчет коэффициентов для методом точек (перепроверить расчеты)
        private void button1_Click_1(object sender, EventArgs e)
        {
            int t = Convert.ToInt32(numericUpDown1.Value) - 1;
            bool checkin = false;

            if (dataGridView1.RowCount == 6)
            {
                for (int i = 0; i < 5; i++)
                    for (int j = 1; j < 4; j++)
                        if (dataGridView1.Rows[i].Cells[j].Value == null || Double.TryParse(dataGridView1.Rows[i].Cells[j].Value.ToString(), out double check) == false)
                        {
                            MessageBox.Show("Один из коэфф. введен неправильно", "Внимание", MessageBoxButtons.OK);
                            checkin = true;
                            break;
                        }
            }
            else
            {
                MessageBox.Show("Введено недостаточно или много коэффициентов", "Внимание", MessageBoxButtons.OK);
                dataGridView1.RowCount = 6;
                checkin = true;
            }
            double s1 = 0, s2 = 0, s3 = 0;
            if (checkin == false)
            {
                //Вычисляем a
                for (int i = 0; i < 5; i++)
                {
                    s1 = s1 + Math.Pow(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value), 2);
                    s2 = s2 + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) * Math.Pow(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value), 2);
                }
                s3 = s1 * s2;
                
                s1 = 0; s2 = 0;
                for (int i = 0; i < 5; i++)
                {
                    s1 = s1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                    s2 = s2 + Math.Pow(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value), 4);
                }
                s3 = s3 - s1 * s2;
                s1 = 0; s2 = 0;

                for (int i = 0; i < 5; i++)
                {
                    s1 = s1 + Math.Pow(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value), 2);
                    s2 = s2 + Math.Pow(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value), 4);
                }
                s3 = s3 / (Math.Pow(s1, 2) - 5 * s2);
                a = s3; s3 = 0; s1 = 0; s2 = 0;

                //Вычисляем b
                for (int i = 0; i < 5; i++)
                {
                    s1 = s1 + Math.Pow(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value), 2);
                    s2 = s2 + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                    s3 = s3 + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) * Math.Pow(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value), 2);
                }
                s3 = 5 * s3 - s1 * s2;
                s1 = 0; s2 = 0;
                for (int i = 0; i < 5; i++)
                {
                    s1 = s1 + Math.Pow(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value), 2);
                    s2 = s2 + Math.Pow(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value), 4);
                }
                s3 = s3 / (Math.Pow(s1, 2) - 5 * s2);
                b = s3; s3 = 0; s1 = 0; s2 = 0;

                //Вычисляем k
                for (int i = 0; i < 5; i++)
                {
                    s1 = s1 + Math.Pow(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value), 4);
                    s2 = s2 + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value);
                }
                s3 = s1 * s2;
                s1 = 0; s2 = 0;

                for (int i = 0; i < 5; i++)
                {
                    s1 = s1 + Math.Pow(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value), 3);
                    s2 = s2 + Math.Pow(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value), 2) * Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value);
                }
                s3 = s3 - s1 * s2;
                k = s3;
                s1 = 0; s2 = 0; s3 = 0;

                for (int i = 0; i < 5; i++)
                {
                    s1 = s1 + Math.Pow(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value), 2);
                    s2 = s2 + Math.Pow(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value), 4);
                    s3 = s3 + Math.Pow(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value), 3);
                }
                k = k / (s1 * s2 - Math.Pow(s3, 2));
                s1 = 0; s2 = 0; s3 = 0;

                //Вычисляем k1
                for (int i = 0; i < 5; i++)
                {
                    s1 = s1 + Math.Pow(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value), 3);
                    s2 = s2 + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value);
                    s3 = s3 + Math.Pow(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value), 2);
                }
                k1 = s1 * s2 - Math.Pow(s3, 2);
                s1 = 0; s2 = 0; s3 = 0;

                for (int i = 0; i < 5; i++)
                {
                    s1 = s1 + Math.Pow(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value), 2);
                    s2 = s2 + Math.Pow(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value), 4);
                    s3 = s3 + Math.Pow(Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value), 3);
                }
                k1 = k1 / (s1 * s2 - Math.Pow(s3, 2));

                dataGridView1.RowCount = 7;

                dataGridView1.Rows[5].Cells[0].Value = "a";
                dataGridView1.Rows[5].Cells[1].Value = "b";
                dataGridView1.Rows[5].Cells[2].Value = "k";
                dataGridView1.Rows[5].Cells[3].Value = "k1";

                dataGridView1.Rows[6].Cells[0].Value = a.ToString();
                dataGridView1.Rows[6].Cells[1].Value = b.ToString();
                dataGridView1.Rows[6].Cells[2].Value = k.ToString();
                dataGridView1.Rows[6].Cells[3].Value = k1.ToString();

                button3.Enabled = true;
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {

            string selectedEmployee = (string)comboBox1.SelectedItem;
            if (selectedEmployee != "Свой вариант (расширенный)")
            {
                int i = Convert.ToInt32(numericUpDown1.Value) - 1;
                
                mark_pump[i] = comboBox1.SelectedItem.ToString();

                if (Double.TryParse(textBox5.Text, out rate_pump[i]) == false)
                    textBox5.Text = "Неправильно введена подача";
                else
                    rate_pump[i] = Convert.ToDouble(textBox5.Text);

                if (Double.TryParse(textBox6.Text, out disk_pump[i]) == false)
                    textBox6.Text = "Неправильно введен диаметр";
                else
                    disk_pump[i] = Convert.ToDouble(textBox6.Text);

                if (String.IsNullOrEmpty(textBox7.Text) == true)
                    textBox7.Text = "Неправильно тип привода";
                else
                    engn_pump[i] = textBox7.Text;

                if (Double.TryParse(textBox8.Text, out efcy_pump[i]) == false)
                    textBox8.Text = "Неправильно введен КПД";
                else
                    efcy_pump[i] = Convert.ToDouble(textBox8.Text);

                if (Double.TryParse(textBox9.Text, out capA_pump[i]) == false)
                    textBox9.Text = "Неправильно введен коэфф.";
                else
                    capA_pump[i] = Convert.ToDouble(textBox9.Text);

                if (Double.TryParse(textBox10.Text, out capB_pump[i]) == false)
                    textBox10.Text = "Неправильно введен коэфф.";
                else
                    capB_pump[i] = Convert.ToDouble(textBox10.Text);

                if (Double.TryParse(textBox21.Text, out kz[i]) == false)
                    textBox21.Text = "Неправильно введен запас.";
                else
                    kz[i] = Convert.ToDouble(textBox21.Text);
            }
            else
                if (selectedEmployee == "Свой вариант (расширенный)")
                {
                    int i = Convert.ToInt32(numericUpDown1.Value) - 1;
                    if (Double.TryParse(textBox11.Text, out Q) == false) //??????????????????????????
                        textBox11.Text = "Неправильно введен расход";
                    else
                    {
                        Q = Convert.ToDouble(textBox11.Text);
                        mark_pump[i] = comboBox1.SelectedItem.ToString();
                        efcy_pump[i] = Q * k - k1 * Math.Pow(Q, 2);
                        capA_pump[i] = a;
                        capB_pump[i] = b;
                        textBox13.Text = efcy_pump[i].ToString();
                        if (Double.TryParse(textBox21.Text, out kz[i]) == false)
                            textBox21.Text = "Неправильно введен запас.";
                        else
                            kz[i] = Convert.ToDouble(textBox21.Text);
                    }
                }
        }
        //Тип насосов
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            int i = Convert.ToInt32(numericUpDown1.Value) - 1;
            string selectedEmployee = mark_pump[i];
            if (selectedEmployee != "Свой вариант (расширенный)")
            {
                if (textBox5.Enabled == false)
                {
                    textBox5.Visible = true;
                    textBox5.Enabled = true;

                    textBox6.Visible = true;
                    textBox6.Enabled = true;

                    textBox7.Visible = true;
                    textBox7.Enabled = true;

                    textBox8.Visible = true;
                    textBox8.Enabled = true;

                    textBox9.Visible = true;
                    textBox9.Enabled = true;

                    textBox10.Visible = true;
                    textBox10.Enabled = true;

                    label11.Visible = true;
                    label12.Visible = true;
                    label13.Visible = true;
                    label14.Visible = true;
                    label15.Visible = true;
                    label16.Visible = true;
                    button1.Visible = false;
                    button1.Enabled = false;

                    dataGridView1.Enabled = false;
                    dataGridView1.Visible = false;

                    button3.Enabled = true;
                }
                comboBox1.SelectedItem = mark_pump[i].ToString();
                textBox5.Text = rate_pump[i].ToString();
                textBox6.Text = disk_pump[i].ToString();
                textBox7.Text = engn_pump[i].ToString();
                textBox8.Text = efcy_pump[i].ToString();
                textBox9.Text = capA_pump[i].ToString();
                textBox10.Text = capB_pump[i].ToString();
            }
            else
            if (selectedEmployee == "Свой вариант (расширенный)")
            {
                if (textBox5.Enabled == true)
                {
                    textBox5.Visible = false;
                    textBox5.Enabled = false;

                    textBox6.Visible = false;
                    textBox6.Enabled = false;

                    textBox7.Visible = false;
                    textBox7.Enabled = false;

                    textBox8.Visible = false;
                    textBox8.Enabled = false;

                    textBox9.Visible = false;
                    textBox9.Enabled = false;

                    textBox10.Visible = false;
                    textBox10.Enabled = false;

                    label11.Visible = false;
                    label12.Visible = false;
                    label13.Visible = false;
                    label14.Visible = false;
                    label15.Visible = false;
                    label16.Visible = false;

                    button1.Visible = true;
                    button1.Enabled = true;

                    dataGridView1.Visible = true;
                    dataGridView1.Enabled = true;

                    button3.Enabled = false;
                }
                comboBox1.SelectedItem = mark_pump[i].ToString();
            }
        }

        //Число подпорных насосов
        int n_pump_p = 0;   //Подпорные
        //Параметры подпорных насосов
        string[] mark_pump_p = new string[100];
        double[] rate_pump_p = new double[100];
        double[] disk_pump_p = new double[100];
        string[] engn_pump_p = new string[100];
        double[] efcy_pump_p = new double[100];
        double[] capA_pump_p = new double[100];
        double[] capB_pump_p = new double[100];
        double[] kz_p = new double[100];
        //Для расчета через аппроксимацию 
        double a_p = 0, b_p = 0, k_p = 0, k1_p = 0;



        //Считывание параметров насосов
        private void button6_Click(object sender, EventArgs e)
        {
            if (numericUpDown3.Value - 1 >= 100)
            {
                numericUpDown3.Value = 100;
                MessageBox.Show("Количество насосов не может превышать 100", "Внимание", MessageBoxButtons.OK);
            }

            n_pump_p = Convert.ToInt32(numericUpDown3.Value);
            dataGridView3.ColumnCount = n_pump_p;
            textBox22.Enabled = true;


            for (int i = 0; i < n_pump_p; i++)
            {
                dataGridView3.Columns[i].Name = i.ToString();
                dataGridView3.Columns[i].DefaultCellStyle.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(63)))), ((int)(((byte)(63)))), ((int)(((byte)(70)))));
                dataGridView3.Columns[i].Width = 35;
            }

            numericUpDown4.Enabled = true;
            numericUpDown4.Maximum = numericUpDown3.Value;
            comboBox2.Enabled = true;
            button5.Enabled = true;

            for (int i = 0; i <= n_pump_p; i++)
            {
                mark_pump_p[i] = comboBox2.Items[0].ToString();
                rate_pump_p[i] = 600.0;
                disk_pump_p[i] = 445.0;
                engn_pump_p[i] = "ВАОВ560М-4У1";
                efcy_pump_p[i] = 77.0;
                capA_pump_p[i] = 74.7;
                capB_pump_p[i] = 4.2600 / 100000;
                kz_p[i] = 4;
            }

            comboBox2.SelectedItem = mark_pump_p[0];
            textBox17.Text = rate_pump_p[0].ToString();
            textBox16.Text = disk_pump_p[0].ToString();
            textBox15.Text = engn_pump_p[0];
            textBox20.Text = efcy_pump_p[0].ToString();
            textBox19.Text = capA_pump_p[0].ToString();
            textBox18.Text = capB_pump_p[0].ToString();
            textBox22.Text = kz_p[0].ToString();
        }
        //Выбор насосов
        private void comboBox2_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            string selectedEmployee = (string)comboBox2.SelectedItem;
            if (selectedEmployee != "Свой вариант (расширенный)")
            {
                textBox15.Visible = true;
                textBox15.Enabled = true;

                textBox16.Visible = true;
                textBox16.Enabled = true;

                textBox17.Visible = true;
                textBox17.Enabled = true;

                textBox18.Visible = true;
                textBox18.Enabled = true;

                textBox19.Visible = true;
                textBox19.Enabled = true;

                textBox20.Visible = true;
                textBox20.Enabled = true;

                label18.Visible = true;
                label19.Visible = true;
                label20.Visible = true;
                label21.Visible = true;
                label22.Visible = true;
                label23.Visible = true;
                button7.Visible = false;
                button7.Enabled = false;

                dataGridView4.Enabled = false;
                dataGridView4.Visible = false;

                button5.Enabled = true;

            }
            else
                if (selectedEmployee == "Свой вариант (расширенный)")
            {
                textBox15.Visible = false;
                textBox15.Enabled = false;

                textBox16.Visible = false;
                textBox16.Enabled = false;

                textBox17.Visible = false;
                textBox17.Enabled = false;

                textBox18.Visible = false;
                textBox18.Enabled = false;

                textBox19.Visible = false;
                textBox19.Enabled = false;

                textBox20.Visible = false;
                textBox20.Enabled = false;

                label18.Visible = false;
                label19.Visible = false;
                label20.Visible = false;
                label21.Visible = false;
                label22.Visible = false;
                label23.Visible = false;

                button7.Visible = true;
                button7.Enabled = true;

                dataGridView4.Visible = true;
                dataGridView4.Enabled = true;

                button5.Enabled = false;
            }
            switch (selectedEmployee)
            {
                case "НПВ 600-60":
                    textBox17.Text = "600,0";
                    textBox16.Text = "445,0";
                    textBox15.Text = "ВАОВ560М-4У1";
                    textBox20.Text = "77,0";
                    textBox19.Text = "74,7";
                    textBox18.Text = (4.2600 / 100000).ToString();
                    textBox22.Text = "4";
                    break;
                case "НПВ 600-60 (215 м3/ч)":
                    textBox17.Text = "215,0";
                    textBox16.Text = "400,0";
                    textBox15.Text = "ВАОВ560М-4У1";
                    textBox20.Text = "77,0";
                    textBox19.Text = "62,2";
                    textBox18.Text = (4.7568 / 100000).ToString();
                    textBox22.Text = "4";
                    break;
                case "НПВ 1250-60":
                    textBox17.Text = "1250,0";
                    textBox16.Text = "525,0";
                    textBox15.Text = "ВАОВ-5К-315-6 УХЛ1";
                    textBox20.Text = "82,0";
                    textBox19.Text = "77,4";
                    textBox18.Text = (1.1368 / 100000).ToString();
                    textBox22.Text = "2,2";
                    break;
                case "НПВ 1250-60 (900 м3/ч)":
                    textBox17.Text = "900,0";
                    textBox16.Text = "500,0";
                    textBox15.Text = "ВАОВ-5К-315-6 УХЛ1";
                    textBox20.Text = "80,0";
                    textBox19.Text = "68,5";
                    textBox18.Text = (1.0448 / 100000).ToString();
                    textBox22.Text = "2,2";
                    break;
                case "НПВ 1250-60 (360 м3/ч)":
                    textBox17.Text = "360,0";
                    textBox16.Text = "475,0";
                    textBox15.Text = "ВАОВ-5К-315-6 УХЛ1";
                    textBox20.Text = "78,0";
                    textBox19.Text = "61,2";
                    textBox18.Text = (9.3754 / 1000000).ToString();
                    textBox22.Text = "2,2";
                    break;
                ////////////////////////////////////////////////
                case "НПВ 2500-80":
                    textBox17.Text = "2500,0";
                    textBox16.Text = "540,0";
                    textBox15.Text = "ВАОВ-5К-800-6";
                    textBox20.Text = "84,0";
                    textBox19.Text = "102,4";
                    textBox18.Text = (3.7584 / 1000000).ToString();
                    textBox22.Text = "2,8";
                    break;
                case "НПВ 2500-80 (1900 м3/ч)":
                    textBox17.Text = "1900,0";
                    textBox16.Text = "515,0";
                    textBox15.Text = "ВАОВ-5К-800-6";
                    textBox20.Text = "82,0";
                    textBox19.Text = "94,6";
                    textBox18.Text = (4.0791 / 1000000).ToString();
                    textBox22.Text = "2,8";
                    break;
                case "НПВ 2500-80 (1100 м3/ч)":
                    textBox17.Text = "1100,0";
                    textBox16.Text = "487,0";
                    textBox15.Text = "ВАОВ-5К-800-6";
                    textBox20.Text = "80,0";
                    textBox19.Text = "85,0 ";
                    textBox18.Text = (4.0795 / 1000000).ToString();
                    textBox22.Text = "2,8";
                    break;
                ////////////////////////////////////////////////
                case "НПВ 3600-90":
                    textBox17.Text = "3600,0";
                    textBox16.Text = "610,0";
                    textBox15.Text = "ВАОВ-5К-1250-6 ";
                    textBox20.Text = "84,0";
                    textBox19.Text = "126,1";
                    textBox18.Text = (2.8040 / 1000000).ToString();
                    textBox22.Text = "3,2";
                    break;
                case "НПВ 3600-90 (2950 м3/ч)":
                    textBox17.Text = "2950,0";
                    textBox16.Text = "580,0";
                    textBox15.Text = "ВАОВ-5К-1250-6 ";
                    textBox20.Text = "82,0";
                    textBox19.Text = "116,2";
                    textBox18.Text = (3.0021 / 1000000).ToString();
                    textBox22.Text = "3,2";
                    break;
                case "НПВ 3600-90 (2200  м3/ч)":
                    textBox17.Text = "2200,0";
                    textBox16.Text = "550,0";
                    textBox15.Text = "ВАОВ-5К-1250-6 ";
                    textBox20.Text = "80,0";
                    textBox19.Text = "104.1";
                    textBox18.Text = (2.9749 / 1000000).ToString();
                    textBox22.Text = "3,2";
                    break;
                ////////////////////////////////////////////////
                case "НПВ 5000-120":
                    textBox17.Text = "5000,0";
                    textBox16.Text = "645,0";
                    textBox15.Text = "ВАОВ-5К-2250-6";
                    textBox20.Text = "85,0";
                    textBox19.Text = "151,8";
                    textBox18.Text = (1.2760 / 1000000).ToString();
                    textBox22.Text = "5";
                    break;
                case "НПВ 5000-120 (3700 м3/ч)":
                    textBox17.Text = "3700,0";
                    textBox16.Text = "613,0";
                    textBox15.Text = "ВАОВ-5К-2250-6";
                    textBox20.Text = "83,0";
                    textBox19.Text = "137,7";
                    textBox18.Text = (1.2839 / 1000000).ToString();
                    textBox22.Text = "5";
                    break;
                case "НПВ 5000-120 (1600 м3/ч)":
                    textBox17.Text = "1600,0";
                    textBox16.Text = "580,0";
                    textBox15.Text = "ВАОВ-5К-2250-6";
                    textBox20.Text = "81,0";
                    textBox19.Text = "123,1";
                    textBox18.Text = (1.2315 / 1000000).ToString();
                    textBox22.Text = "5";
                    break;
                ////////////////////////////////////////////////
                case "Свой вариант (a-bQQ)":
                    textBox17.Text = "";
                    textBox16.Text = "";
                    textBox15.Text = "";
                    textBox20.Text = "";
                    textBox19.Text = "";
                    textBox18.Text = "";
                    textBox22.Text = "";
                    break;
                case "Свой вариант (расширенный)":
                    textBox17.Text = "";
                    textBox16.Text = "";
                    textBox15.Text = "";
                    textBox20.Text = "";
                    textBox19.Text = "";
                    textBox18.Text = "";
                    textBox22.Text = "";
                    break;
            }
        }
        //Тип насосов
        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            int i = Convert.ToInt32(numericUpDown4.Value) - 1;
            string selectedEmployee = mark_pump_p[i];
            if (selectedEmployee != "Свой вариант (расширенный)")
            {
                if (textBox17.Enabled == false)
                {
                    textBox17.Visible = true;
                    textBox17.Enabled = true;

                    textBox16.Visible = true;
                    textBox16.Enabled = true;

                    textBox15.Visible = true;
                    textBox15.Enabled = true;

                    textBox20.Visible = true;
                    textBox20.Enabled = true;

                    textBox19.Visible = true;
                    textBox19.Enabled = true;

                    textBox18.Visible = true;
                    textBox18.Enabled = true;

                    label20.Visible = true;
                    label19.Visible = true;
                    label18.Visible = true;
                    label21.Visible = true;
                    label22.Visible = true;
                    label23.Visible = true;

                    button7.Visible = false;
                    button7.Enabled = false;

                    dataGridView4.Enabled = false;
                    dataGridView4.Visible = false;

                    button5.Enabled = true;
                }
                comboBox2.SelectedItem = mark_pump_p[i].ToString();
                textBox17.Text = rate_pump_p[i].ToString();
                textBox16.Text = disk_pump_p[i].ToString();
                textBox15.Text = engn_pump_p[i].ToString();
                textBox20.Text = efcy_pump_p[i].ToString();
                textBox19.Text = capA_pump_p[i].ToString();
                textBox18.Text = capB_pump_p[i].ToString();
            }
            else
            if (selectedEmployee == "Свой вариант (расширенный)")
            {
                if (textBox17.Enabled == true)
                {
                    textBox17.Visible = false;
                    textBox17.Enabled = false;

                    textBox16.Visible = false;
                    textBox16.Enabled = false;

                    textBox15.Visible = false;
                    textBox15.Enabled = false;

                    textBox20.Visible = false;
                    textBox20.Enabled = false;

                    textBox19.Visible = false;
                    textBox19.Enabled = false;

                    textBox18.Visible = false;
                    textBox18.Enabled = false;

                    label18.Visible = false;
                    label19.Visible = false;
                    label20.Visible = false;
                    label21.Visible = false;
                    label22.Visible = false;
                    label23.Visible = false;

                    button7.Visible = true;
                    button7.Enabled = true;

                    dataGridView4.Visible = true;
                    dataGridView4.Enabled = true;

                    button5.Enabled = false;
                }
                comboBox2.SelectedItem = mark_pump_p[i].ToString();
            }
        }
        //Расчет насосов
        private void button7_Click(object sender, EventArgs e)
        {
            int t = Convert.ToInt32(numericUpDown2.Value) - 1;
            bool checkin = false;
            if (dataGridView4.RowCount == 6)
            {
                for (int i = 0; i < 5; i++)
                    for (int j = 1; j < 4; j++)
                        if (dataGridView4.Rows[i].Cells[j].Value == null || Double.TryParse(dataGridView4.Rows[i].Cells[j].Value.ToString(), out double check) == false)
                        {
                            MessageBox.Show("Один из коэфф. введен неправильно", "Внимание", MessageBoxButtons.OK);
                            checkin = true;
                            break;
                        }
            }
            else
            {
                MessageBox.Show("Введено недостаточно коэффициентов", "Внимание", MessageBoxButtons.OK);
                dataGridView4.RowCount = 6;
                checkin = true;
            }
            double s1 = 0, s2 = 0, s3 = 0;
            if (checkin == false)
            {
                //Вычисляем a
                for (int i = 0; i < 5; i++)
                {
                    s1 = s1 + Math.Pow(Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value), 2);
                    s2 = s2 + Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) * Math.Pow(Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value), 2);
                }
                s3 = s1 * s2;

                s1 = 0; s2 = 0;
                for (int i = 0; i < 5; i++)
                {
                    s1 = s1 + Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value);
                    s2 = s2 + Math.Pow(Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value), 4);
                }
                s3 = s3 - s1 * s2;
                s1 = 0; s2 = 0;

                for (int i = 0; i < 5; i++)
                {
                    s1 = s1 + Math.Pow(Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value), 2);
                    s2 = s2 + Math.Pow(Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value), 4);
                }
                s3 = s3 / (Math.Pow(s1, 2) - 5 * s2);
                a_p = s3; s3 = 0; s1 = 0; s2 = 0;
                //Вычисляем b
                for (int i = 0; i < 5; i++)
                {
                    s1 = s1 + Math.Pow(Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value), 2);
                    s2 = s2 + Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value);
                    s3 = s3 + Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) * Math.Pow(Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value), 2);
                }
                s3 = 5 * s3 - s1 * s2;
                s1 = 0; s2 = 0;
                for (int i = 0; i < 5; i++)
                {
                    s1 = s1 + Math.Pow(Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value), 2);
                    s2 = s2 + Math.Pow(Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value), 4);
                }
                s3 = s3 / (Math.Pow(s1, 2) - 5 * s2);
                b_p = s3; s3 = 0; s1 = 0; s2 = 0;

                //Вычисляем k
                for (int i = 0; i < 5; i++)
                {
                    s1 = s1 + Math.Pow(Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value), 4);
                    s2 = s2 + Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value) * Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value);
                }
                s3 = s1 * s2;
                s1 = 0; s2 = 0;

                for (int i = 0; i < 5; i++)
                {
                    s1 = s1 + Math.Pow(Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value), 3);
                    s2 = s2 + Math.Pow(Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value), 2) * Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value);
                }
                s3 = s3 - s1 * s2;
                k_p = s3;
                s1 = 0; s2 = 0; s3 = 0;

                for (int i = 0; i < 5; i++)
                {
                    s1 = s1 + Math.Pow(Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value), 2);
                    s2 = s2 + Math.Pow(Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value), 4);
                    s3 = s3 + Math.Pow(Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value), 3);
                }
                k_p = k_p / (s1 * s2 - Math.Pow(s3, 2));
                s1 = 0; s2 = 0; s3 = 0;

                //Вычисляем k1
                for (int i = 0; i < 5; i++)
                {
                    s1 = s1 + Math.Pow(Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value), 3);
                    s2 = s2 + Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value) * Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value);
                    s3 = s3 + Math.Pow(Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value), 2);
                }
                k1_p = s1 * s2 - Math.Pow(s3, 2);
                s1 = 0; s2 = 0; s3 = 0;

                for (int i = 0; i < 5; i++)
                {
                    s1 = s1 + Math.Pow(Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value), 2);
                    s2 = s2 + Math.Pow(Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value), 4);
                    s3 = s3 + Math.Pow(Convert.ToDouble(dataGridView4.Rows[i].Cells[1].Value), 3);
                }
                k1_p = k1_p / (s1 * s2 - Math.Pow(s3, 2));

                dataGridView4.RowCount = 7;

                dataGridView4.Rows[5].Cells[0].Value = "a";
                dataGridView4.Rows[5].Cells[1].Value = "b";
                dataGridView4.Rows[5].Cells[2].Value = "k";
                dataGridView4.Rows[5].Cells[3].Value = "k1";

                dataGridView4.Rows[6].Cells[0].Value = a_p.ToString();
                dataGridView4.Rows[6].Cells[1].Value = b_p.ToString();
                dataGridView4.Rows[6].Cells[2].Value = k_p.ToString();
                dataGridView4.Rows[6].Cells[3].Value = k1_p.ToString();
                button5.Enabled = true;
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            string selectedEmployee = (string)comboBox2.SelectedItem;
            if (selectedEmployee != "Свой вариант (расширенный)")
            {
                int i = Convert.ToInt32(numericUpDown4.Value) - 1;
                mark_pump_p[i] = comboBox2.SelectedItem.ToString();

                if (Double.TryParse(textBox17.Text, out rate_pump_p[i]) == false)
                    textBox17.Text = "Неправильно введена подача";
                else
                    rate_pump_p[i] = Convert.ToDouble(textBox17.Text);

                if (Double.TryParse(textBox16.Text, out disk_pump_p[i]) == false)
                    textBox16.Text = "Неправильно введен диаметр";
                else
                    disk_pump_p[i] = Convert.ToDouble(textBox16.Text);

                if (Double.TryParse(textBox20.Text, out efcy_pump_p[i]) == false)
                    textBox20.Text = "Неправильно введен КПД";
                else
                    efcy_pump_p[i] = Convert.ToDouble(textBox20.Text);

                if (Double.TryParse(textBox19.Text, out capA_pump_p[i]) == false)
                    textBox19.Text = "Неправильно введен коэфф.";
                else
                    capA_pump_p[i] = Convert.ToDouble(textBox19.Text);

                if (Double.TryParse(textBox18.Text, out capB_pump_p[i]) == false)
                    textBox18.Text = "Неправильно введен коэфф.";
                else
                    capB_pump_p[i] = Convert.ToDouble(textBox18.Text);

                if (Double.TryParse(textBox22.Text, out kz_p[i]) == false)
                    textBox22.Text = "Неправильно введен запас.";
                else
                    kz_p[i] = Convert.ToDouble(textBox22.Text);

                if (String.IsNullOrEmpty(textBox15.Text) == true)
                    textBox15.Text = "Неправильно тип привода";
                else
                    engn_pump_p[i] = textBox15.Text;

            }
            else
                if (selectedEmployee == "Свой вариант (расширенный)")
            {
                int i = Convert.ToInt32(numericUpDown4.Value) - 1;
                if (Double.TryParse(textBox11.Text, out Q) == false)
                    textBox11.Text = "Неправильно введен расход";
                else
                {
                    Q = Convert.ToDouble(textBox11.Text);
                    mark_pump_p[i] = comboBox2.SelectedItem.ToString();
                    efcy_pump_p[i] = Q * k_p - k1_p * Math.Pow(Q, 2);
                    capA_pump_p[i] = a_p;
                    capB_pump_p[i] = b_p;
                    textBox12.Text = efcy_pump_p[i].ToString();
                    if (Double.TryParse(textBox22.Text, out kz_p[i]) == false)
                        textBox22.Text = "Неправильно введен запас.";
                    else
                        kz_p[i] = Convert.ToDouble(textBox22.Text);
                }
            }
        }


        //Функция для параллельного расчета (магистральный)
        private double f(int r, int h, double x)
        {
            double v = 0;
            int j = 0;
            for (int i = 0; i < h; i++)
            {
                j = Convert.ToInt32(dataGridView2.Rows[i].Cells[r].Value) - 1;
                if (capA_pump[j] >= x)
                    v = v + Math.Sqrt((capA_pump[j] - x) / capB_pump[j]);
                else
                    textBox13.Text = "Схема магистральных насосов неверна";
            }
            v = v - Q * 3600;
            return v;
        }
        //Функция для параллельного расчета (подпорный)
        private double f_p(int r, int h, double x)
        {
            double v = 0;
            int j = 0;
            for (int i = 0; i < h; i++)
            {
                j = Convert.ToInt32(dataGridView3.Rows[i].Cells[r].Value) - 1;
                if (capA_pump_p[j] >= x)
                    v = v + Math.Sqrt((capA_pump_p[j] - x) / capB_pump_p[j]);
                else
                    textBox12.Text = "Схема магистральных насосов неверна";
            }
            v = v - Q * 3600;
            return v;
        }
        //Итоговый расчет насосов (напор и расход НПА)
        private void button4_Click(object sender, EventArgs e)
        {
            double[] h_pump = new double[n_pump];
            bool[] check_pump = new bool[n_pump];

            double[] h_pump_p = new double[n_pump_p];
            bool[] check_pump_p = new bool[n_pump_p];

            double h_sum = 0, h_sum_p = 0;
            int t = 0;


            textBox12.Text = "0";
            textBox13.Text = "0";

            for (int i = 0; i < n_pump; i++)
            {
                h_pump[i] = 0;
                check_pump[i] = false;
            }
            for (int i = 0; i < n_pump_p; i++)
            {
                h_pump_p[i] = 0;
                check_pump_p[i] = false;
            }

            if (Double.TryParse(textBox11.Text, out Q) == false)
                textBox11.Text = "Неправильно введен расход";

            //Проверка что номер насоса вводимый пользователем принимает значения от 1 до n_pump
            double checkin = 0;

            //Магистральные насосы
            for (int j = 0; j < dataGridView2.ColumnCount; j++)
            {

                //Последовательное соединение
                if (dataGridView2.Rows[0].Cells[j].Value != null && Double.TryParse(dataGridView2.Rows[0].Cells[j].Value.ToString(), out checkin) == true 
                    && dataGridView2.Rows[1].Cells[j].Value == null && checkin > 0 && checkin <= n_pump)
                {
                    if (check_pump[Convert.ToInt32(dataGridView2.Rows[0].Cells[j].Value) - 1] == false)
                    {
                        check_pump[Convert.ToInt32(dataGridView2.Rows[0].Cells[j].Value) - 1] = true;
                        //Если в первом ряду насос 1, то его запас кавитации равен общему
                        if (j == 0)
                            textBox23.Text = kz[Convert.ToInt32(dataGridView2.Rows[0].Cells[0].Value) - 1].ToString();
                        h_pump[t] = capA_pump[Convert.ToInt32(dataGridView2.Rows[0].Cells[j].Value) - 1] - capB_pump[Convert.ToInt32(dataGridView2.Rows[0].Cells[j].Value) - 1] * Q * Q * 3600 * 3600;
                        h_sum = h_sum + h_pump[t];
                        t = t + 1;
                    }
                    else //Проверка от насосов с одинаковыми номерами 
                    {
                        textBox13.Text = "Схема магистральных насосов неверна";
                        break;
                    }
                }
                //Если номер насоса не в диапазоне или введены абракадабра вместо номера насоса
                else if ((dataGridView2.Rows[0].Cells[j].Value != null && Double.TryParse(dataGridView2.Rows[0].Cells[j].Value.ToString(), out checkin) == false)
                || checkin < 0 || checkin > n_pump)
                {
                    textBox13.Text = "Схема магистральных насосов неверна";
                    break;
                }

                t = 0;

                //Параллельное соединение
                if (dataGridView2.Rows[0].Cells[j].Value != null && dataGridView2.Rows[1].Cells[j].Value != null)
                {
                    int num = 0;
                    for (int i = 0; i < dataGridView2.RowCount - 1; i++)
                        if (dataGridView2.Rows[i].Cells[j].Value != null && Double.TryParse(dataGridView2.Rows[i].Cells[j].Value.ToString(), out checkin) == true
                            && checkin > 0 && checkin <= n_pump)
                            num = num + 1;
                        else
                            if (checkin <= 0 || checkin > n_pump)
                            textBox13.Text = "Схема магистральных насосов неверна";

                    if (textBox13.Text == "0")
                    {
                        double k1 = 0, k2 = 0, c = 0;
                        for (int i = 0; i < num; i++)
                        {
                            if (capA_pump[Convert.ToInt32(dataGridView2.Rows[i].Cells[j].Value) - 1] >= k2)
                                k2 = capA_pump[Convert.ToInt32(dataGridView2.Rows[i].Cells[j].Value) - 1];
                        }

                        //Кав. запас линии параллельно соединенных насосов равен наибольшему кав. запасу из этого ряда
                        if (j == 0)
                            for (int i = 0; i < num; i++)
                                if (Convert.ToDouble(textBox23.Text) < kz[Convert.ToInt32(dataGridView2.Rows[i].Cells[0].Value) - 1])
                                    textBox23.Text = kz[Convert.ToInt32(dataGridView2.Rows[i].Cells[0].Value) - 1].ToString();

                        for (int i = 0; i < num; i++)
                            if (check_pump[Convert.ToInt32(dataGridView2.Rows[i].Cells[j].Value) - 1] == false)
                                check_pump[Convert.ToInt32(dataGridView2.Rows[i].Cells[j].Value) - 1] = true;
                            else
                            {
                                textBox13.Text = "Схема магистральных насосов неверна";
                                break;
                            }

                        if (f(j, num, k1) * f(j, num, k2) <= 0)
                        {
                            do
                            {
                                c = (k1 + k2) / 2;
                                if (f(j, num, k1) * f(j, num, c) <= 0)
                                    k2 = c;
                                else
                                    k1 = c;
                            } while (Math.Abs(k1 - k2) > eps);
                            h_pump[t] = c;
                            h_sum = h_sum + h_pump[t];
                            t = t + 1;
                        }
                        else
                        {
                            textBox13.Text = "Схема магистральных насосов неверна";
                            break;
                        }
                    }
                }
            
            }


            //Подпорные насосы
            for (int j = 0; j < dataGridView3.ColumnCount; j++)
            {
                checkin = 0;
                t = 0;
                //Последовательное соединение
                if (dataGridView3.Rows[0].Cells[j].Value != null && Double.TryParse(dataGridView3.Rows[0].Cells[j].Value.ToString(), out checkin) == true
                    && dataGridView3.Rows[1].Cells[j].Value == null && checkin > 0 && checkin <= n_pump_p)
                {
                    if (check_pump_p[Convert.ToInt32(dataGridView3.Rows[0].Cells[j].Value) - 1] == false)
                    {
                        check_pump_p[Convert.ToInt32(dataGridView3.Rows[0].Cells[j].Value) - 1] = true;
                        //Если в первом ряду насос 1, то его запас кавитации равен общему
                        if (j == 0)
                            textBox24.Text = kz_p[Convert.ToInt32(dataGridView3.Rows[0].Cells[0].Value) - 1].ToString();
                        h_pump_p[t] = capA_pump_p[Convert.ToInt32(dataGridView3.Rows[0].Cells[j].Value) - 1] - capB_pump_p[Convert.ToInt32(dataGridView3.Rows[0].Cells[j].Value) - 1] * Q * Q * 3600 * 3600;
                        h_sum_p = h_sum_p + h_pump_p[t];
                        t = t + 1;
                    }
                    else //Проверка от насосов с одинаковыми номерами 
                    {
                        textBox12.Text = "Схема магистральных насосов неверна";
                        break;
                    }
                }
                //Если номер насоса не в диапазоне или введены абракадабра вместо номера насоса
                else if ((dataGridView3.Rows[0].Cells[j].Value != null && Double.TryParse(dataGridView3.Rows[0].Cells[j].Value.ToString(), out checkin) == false)
                || checkin < 0 || checkin > n_pump_p)
                {
                    textBox12.Text = "Схема магистральных насосов неверна";
                    break;
                }

                t = 0;

                //Параллельное соединение
                if (dataGridView3.Rows[0].Cells[j].Value != null && dataGridView3.Rows[1].Cells[j].Value != null)
                {
                    int num = 0;
                    for (int i = 0; i < dataGridView3.RowCount - 1; i++)
                        if (dataGridView3.Rows[i].Cells[j].Value != null && Double.TryParse(dataGridView3.Rows[i].Cells[j].Value.ToString(), out checkin) == true
                            && checkin > 0 && checkin <= n_pump_p)
                            num = num + 1;
                        else
                            if (checkin <= 0 || checkin > n_pump_p)
                            textBox12.Text = "Схема магистральных насосов неверна";

                    if (textBox12.Text == "0")
                    {
                        double k1 = 0, k2 = 0, c = 0;
                        for (int i = 0; i < num; i++)
                        {
                            if (capA_pump_p[Convert.ToInt32(dataGridView3.Rows[i].Cells[j].Value) - 1] >= k2)
                                k2 = capA_pump_p[Convert.ToInt32(dataGridView3.Rows[i].Cells[j].Value) - 1];
                        }

                        //Кав. запас линии параллельно соединенных насосов равен наибольшему кав. запасу из этого ряда
                        if (j == 0)
                            for (int i = 0; i < num; i++)
                                if (Convert.ToDouble(textBox24.Text) < kz_p[Convert.ToInt32(dataGridView3.Rows[i].Cells[0].Value) - 1])
                                    textBox24.Text = kz_p[Convert.ToInt32(dataGridView3.Rows[i].Cells[0].Value) - 1].ToString();

                        for (int i = 0; i < num; i++)
                            if (check_pump_p[Convert.ToInt32(dataGridView3.Rows[i].Cells[j].Value) - 1] == false)
                                check_pump_p[Convert.ToInt32(dataGridView3.Rows[i].Cells[j].Value) - 1] = true;
                            else
                            {
                                textBox12.Text = "Схема магистральных насосов неверна";
                                break;
                            }

                        if (f_p(j, num, k1) * f_p(j, num, k2) <= 0)
                        {
                            do
                            {
                                c = (k1 + k2) / 2;
                                if (f_p(j, num, k1) * f_p(j, num, c) <= 0)
                                    k2 = c;
                                else
                                    k1 = c;
                            } while (Math.Abs(k1 - k2) > eps);
                            h_pump_p[t] = c;
                            h_sum_p = h_sum_p + h_pump_p[t];
                            t = t + 1;
                        }
                        else
                        {
                            textBox12.Text = "Схема магистральных насосов неверна";
                            break;
                        }
                    }
                }

            }

            if (textBox13.Text == "0")
                textBox13.Text = h_sum.ToString();

            if (textBox12.Text == "0")
                textBox12.Text = h_sum_p.ToString();

            double wr;
            if (double.TryParse(textBox12.Text, out wr) == true && double.TryParse(textBox13.Text, out wr) == true && Convert.ToDouble(textBox12.Text) > Convert.ToDouble(textBox23.Text))
                textBox14.Text = (h_sum + h_sum_p).ToString();
            else
                textBox14.Text = "h_кав > h_под";
            if (h_sum < 0)
                textBox13.Text = "Насосы подобраны не верно";
            if (h_sum_p < 0)
                textBox12.Text = "Насосы подобраны не верно";

        }

        //Блок 2 - Параметры НП и трассы
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        bool check_lin = false;     //Проверка ввода параметров линейки
        bool check_fuels = false;     //Проверка ввода параметров НП
        //Параметры трассы
        //Вводимые
        double[] m_x = new double[5000];
        double[] m_z = new double[5000];
        double[] m_D = new double[5000];
        double[] m_sig = new double[5000];
        double[] m_del = new double[5000];
        double[] m_Tvn = new double[5000];
        double[] m_lyam_tr = new double[5000];
        double[] m_lyam_gr = new double[5000];
        double[] m_h0 = new double[5000];
        double[] m_Tst= new double[5000];
        //Расчитываемые
        double[] m_s = new double[5000];
        double[] m_vV = new double[5000];
        double[] m_eps = new double[5000];
        int[] m_ps = new int[5000];
        //Параметры жидкости
        double[] ro = new double[5000];
        double[] nu = new double[5000];
        double[] T = new double[5000];
        string[] fuels = new string[5000];
        double[] mass = new double[5000];
        double[] cp = new double[5000];
        double[] lyam = new double[5000];
        double[] a0 = new double[5000];
        double[] u0 = new double[5000];
        double[] T0 = new double[5000];
        //Расчитываемые
        double[] vol = new double[5000];
        //Параметры графики
        Color[] cl_h = new Color[1000];
        Color[] cl_p = new Color[1000];
        Color[] cl_ro = new Color[1000];
        Color[] cl_T = new Color[1000];
        Color Back = Color.FromArgb(63, 63, 70);        //Цвет фона
        const float grid = 25, lw = 2;                  //Размер ячейки и толщина линии графиков
        Color cl_g = Color.FromArgb(103, 103, 110);     //Цвет граинцы ячейки
        double Tvnmax = 0, Tvnmin = 0;                  //Максимальная и минимальная температуры грунта
        double Tstmax = 0, Tstmin = 0;                  //Максимальная и минимальная температуры стенки
        double Tmax = 0, Tmin = 0;                      //Максимальная и минимальная НП
        double romax = 0,  romin = 0;                   //Максимальная и минимальная плотности НП
        double numax = 0, numin = 0;                    //Максимальная и минимальная вязкости НП
        double length = 0;                              //Длина трубы трубопровода
        Color T_c = Color.FromArgb(66,133,244);         //Цвет графика температуры окр. среды
        Color z_c = Color.FromArgb(244, 244, 133);      //Цвет графика температуры окр. среды                                                                                     //Для расчета температуры
        int num_fluid = 0;                              //Номер расчитываемой жидкости
        double[] x = new double[n];
        double[] z = new double[n];
        double[] Tvn = new double[n];
        double deltax = 0;
        double[] D = new double[n];
        double[] sig = new double[n];  
        
        double[] a2 = new double[n];
        double[,] cp_st = new double[1000, n];
        double[,] ro_st = new double[1000, n];
        double[,] nu_st = new double[1000, n];
        double[,] lyam_st = new double[1000, n];
        double[,] pr_st = new double[1000, n];
        double[,] cp_rl = new double[1000, n];
        double[,] ro_rl = new double[1000, n];
        double[,] nu_rl = new double[1000, n];
        double[,] lyam_rl = new double[1000, n];
        double[,] pr_rl = new double[1000, n];
        double[] v = new double[n];
        double[,] a1 = new double[1000, n];
        double[,] k_sp = new double[1000, n];
        double Qm = 0;
        double[,] Shu = new double[1000, n];
        double[,] Re = new double[1000, n];
        double[] eps_sp = new double[n];
        double[,] lyam_sp = new double[1000, n];
        double[,] i_sp = new double[1000, n];
        double[,] b_sp = new double[1000, n];
        double[,] Tk = new double[1000, n];

        //Расчет местоположения жидкости;
        double[] x_s = new double[1000];
        double[] x_f = new double[1000];
        //Расчет давления и напора жидкостей
        //double[] H = new double[1000];
        //double[] P = new double[1000];
        //Расчет объема смеси
        double[] vol_s = new double[1000];

        double[] s_smes = new double[1000];         //Начало смеси
        double[] f_smes = new double[1000];         //Конец смеси
        const  int n_sq = 100;                      //Количество отрезков смеси 
        int i_i = 1;                                //Номер расчитываемой смеси

        bool pumping = false;
        int key_p = 0;

        //Ввод данных трассы
        private void button13_Click(object sender, EventArgs e)
        {
            check_lin = false;
            double res = 0;

            Microsoft.Office.Interop.Excel.Application ObjExcel;
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Файл Excel|*.XLSX;*.XLS";
            openDialog.ShowDialog();
            ObjExcel = new Microsoft.Office.Interop.Excel.Application();

            if (File.Exists(openDialog.FileName) == true)
            {
                ObjWorkBook = ObjExcel.Workbooks.Open(openDialog.FileName);
                ObjWorkSheet = ObjExcel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            }
            else
            {
                ObjWorkBook = null;
                ObjWorkSheet = null;
            }
            if (Double.TryParse(textBox11.Text, out Q) == true)
            {
                try
                {
                    Microsoft.Office.Interop.Excel.Range rg = null;
                    Int32 row = 1;
                    dataGridView5.Rows.Clear();
                    List<String> arr = new List<string>();
                    while (ObjWorkSheet.get_Range("a" + row, "a" + row).Value != null && check_lin == false)
                    {
                        rg = ObjWorkSheet.get_Range("a" + row, "u" + row);
                        foreach (Microsoft.Office.Interop.Excel.Range item in rg)
                        {
                            try
                            {
                                arr.Add(item.Value.ToString().Trim());
                            }
                            catch { arr.Add(""); }
                        }
                        for (int i = 0; i < 10; i++)
                            if ((i != 0 && Double.TryParse(arr[i], out res) == false))
                                check_lin = true;
                        dataGridView5.Rows.Add(arr[0], arr[1], arr[2], arr[3], arr[4], arr[5], arr[6], arr[7], arr[8], arr[9]);
                        arr.Clear();
                        row++;
                    }

                    if (check_lin == false)
                    {
                        //Расчет

                        for (int i = 0; i < dataGridView5.RowCount - 1; i++)
                        {
                            //Ввод
                            m_x[i] = Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value);
                            m_z[i] = Convert.ToDouble(dataGridView5.Rows[i].Cells[1].Value);
                            m_D[i] = Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value);
                            m_sig[i] = Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value);
                            m_del[i] = Convert.ToDouble(dataGridView5.Rows[i].Cells[4].Value);
                            m_Tvn[i] = Convert.ToDouble(dataGridView5.Rows[i].Cells[5].Value);
                            m_lyam_tr[i] = Convert.ToDouble(dataGridView5.Rows[i].Cells[6].Value);
                            m_lyam_gr[i] = Convert.ToDouble(dataGridView5.Rows[i].Cells[7].Value);
                            m_h0[i] = Convert.ToDouble(dataGridView5.Rows[i].Cells[8].Value);
                            m_Tst[i] = Convert.ToDouble(dataGridView5.Rows[i].Cells[9].Value);

                            if (i == 0)
                            {
                                Tvnmax = m_Tvn[i];
                                Tvnmin = m_Tvn[i];

                                Tstmax = m_Tst[i];
                                Tstmin = m_Tst[i];
                            }
                            else
                            {
                                if (Tvnmax < m_Tvn[i])
                                    Tvnmax = m_Tvn[i];
                                if (Tvnmin > m_Tvn[i])
                                    Tvnmin = m_Tvn[i];

                                if (Tstmax < m_Tst[i])
                                    Tstmax = m_Tst[i];
                                if (Tstmin > m_Tst[i])
                                    Tstmin = m_Tst[i];
                            }

                            if (i == dataGridView5.RowCount - 2)
                                length = m_x[i];
                        }

                        MessageBox.Show("Файл успешно считан!", "Чтение файла", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                    }
                }
                catch (Exception ex) { MessageBox.Show("Ошибка: " + ex.Message, "Ошибка чтения файла", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                {
                    if (ObjWorkBook != null)
                        ObjWorkBook.Close(false, "", null);
                    ObjExcel.Quit();
                    ObjWorkBook = null;
                    ObjWorkSheet = null;
                    ObjExcel = null;
                }

            }
            else
            {
                MessageBox.Show("Неправильно введен расход", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox11.Text = "Неправильно введен расход";
            }
            if (check_lin == true)
            {
                MessageBox.Show("Ошибка чтения файла, переменные введены неверно ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //Обновление данных трассы
        private void button8_Click(object sender, EventArgs e)
        {
            double res = 0;
            for (int i = 0; i < dataGridView5.RowCount - 1; i++)
                for (int j = 0; j < 10; j++)
                    if ((Double.TryParse(dataGridView5.Rows[i].Cells[j].Value.ToString(), out res) == false && j != 0) || res < 0)
                        check_lin = true;

            if (check_lin == false)
            {
                for (int i = 0; i < dataGridView5.RowCount - 1; i++)
                {
                    m_x[i] = Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value);
                    m_z[i] = Convert.ToDouble(dataGridView5.Rows[i].Cells[1].Value);
                    m_D[i] = Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value);
                    m_sig[i] = Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value);
                    m_del[i] = Convert.ToDouble(dataGridView5.Rows[i].Cells[4].Value);
                    m_Tvn[i] = Convert.ToDouble(dataGridView5.Rows[i].Cells[5].Value);
                    m_lyam_tr[i] = Convert.ToDouble(dataGridView5.Rows[i].Cells[6].Value);
                    m_lyam_gr[i] = Convert.ToDouble(dataGridView5.Rows[i].Cells[7].Value);
                    m_h0[i] = Convert.ToDouble(dataGridView5.Rows[i].Cells[8].Value);
                    m_Tst[i] = Convert.ToDouble(dataGridView5.Rows[i].Cells[9].Value);

                    if (i == 0)
                    {
                        Tvnmax = m_Tvn[i];
                        Tvnmin = m_Tvn[i];
                    }
                    else
                    {
                        if (Tvnmax < m_Tvn[i])
                            Tvnmax = m_Tvn[i];
                        if (Tvnmin > m_Tvn[i])
                            Tvnmin = m_Tvn[i];
                    }
                    if (i == dataGridView5.RowCount - 2)
                        length = m_x[i];
                }
                MessageBox.Show("Данные успешно обновлены!", "Данные", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                MessageBox.Show("Ошибка чтения данных, переменные введены неверно ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        //Ввод данных НП
        private void button15_Click(object sender, EventArgs e)
        {
            chart1.Series.Clear();
            chart4.Series.Clear();
            chart5.Series.Clear();

            chart1.Series.Add("z");
            chart1.Series["z"].ChartType = SeriesChartType.Spline;
            chart1.Series["z"].Color = z_c;
            chart1.Series["z"].BorderWidth = 2;

            double res = 0;

            Microsoft.Office.Interop.Excel.Application ObjExcel;
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Файл Excel|*.XLSX;*.XLS";
            openDialog.ShowDialog();
            ObjExcel = new Microsoft.Office.Interop.Excel.Application();

            if (File.Exists(openDialog.FileName) == true)
            {
                ObjWorkBook = ObjExcel.Workbooks.Open(openDialog.FileName);
                ObjWorkSheet = ObjExcel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            }
            else
            {
                ObjWorkBook = null;
                ObjWorkSheet = null;
            }
            try
            {
                Microsoft.Office.Interop.Excel.Range rg = null;
                Int32 row = 1;
                dataGridView6.Rows.Clear();
                List<String> arr = new List<string>();
                while (ObjWorkSheet.get_Range("a" + row, "a" + row).Value != null && check_fuels == false)
                {
                    rg = ObjWorkSheet.get_Range("a" + row, "u" + row);
                    foreach (Microsoft.Office.Interop.Excel.Range item in rg)
                    {
                        try
                        {
                            arr.Add(item.Value.ToString().Trim());
                        }
                        catch { arr.Add(""); }
                    }
                    for (int i = 0; i < 10; i++)
                        if ((i != 3 && Double.TryParse(arr[i], out res) == false) || (res < 0))
                            check_fuels = true;
                    dataGridView6.Rows.Add(arr[0], arr[1], arr[2], arr[3], arr[4], arr[5], arr[6], arr[7], arr[8], arr[9]);
                    arr.Clear();
                    row++;
                }
                if (check_fuels == false)
                {
                    //Расчет

                    for (int i = 0; i < dataGridView6.RowCount - 1; i++)
                    {
                        ro[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[0].Value);
                        nu[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[1].Value);
                        T[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[2].Value);
                        fuels[i] = dataGridView6.Rows[i].Cells[3].Value.ToString();
                        mass[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[4].Value);
                        cp[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[5].Value);
                        lyam[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[6].Value);
                        a0[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[7].Value);
                        u0[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[8].Value);
                        T0[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[9].Value);
                        //Расчет
                        if (i == 0)
                        {
                            romax = ro[i];
                            romin = ro[i];

                            Tmax = T0[i];
                            Tmin = T0[i];

                            numax = nu[i];
                            numin = nu[i];
                        }
                        else
                        {
                            if (romax < ro[i])
                                romax = ro[i];
                            if (romax > ro[i])
                                romax = ro[i];

                            if (Tmax < T0[i])
                                Tmax = T0[i];
                            if (Tmin > T0[i])
                                Tmin = T0[i];

                            if (numax < nu[i])
                                numax = nu[i];
                            if (numin > nu[i])
                                numin = nu[i];
                        }


                        vol[i] = mass[i] * 1000 / ro[i];
                        cl_h[i] = Color.FromArgb(rnd.Next(0, 64), rnd.Next(128, 255), rnd.Next(0, 64));         //green
                        cl_p[i] = Color.FromArgb(rnd.Next(128, 255), rnd.Next(0, 64), rnd.Next(0, 64));         //red
                        cl_ro[i] = Color.FromArgb(rnd.Next(128, 255), rnd.Next(128, 255), rnd.Next(128, 255));  //any
                        cl_T[i] = Color.FromArgb(rnd.Next(0, 64), rnd.Next(0, 64), rnd.Next(128, 255));         //blue

                        chart1.Series.Add(fuels[i] + "_H");
                        chart1.Series[fuels[i] + "_H"].ChartType = SeriesChartType.Line;
                        chart1.Series[fuels[i] + "_H"].Color = cl_h[i];
                        chart1.Series[fuels[i] + "_H"].BorderWidth = 2;

                        chart1.Series.Add(fuels[i] + "_P");
                        chart1.Series[fuels[i] + "_P"].YAxisType = AxisType.Secondary;
                        chart1.Series[fuels[i] + "_P"].ChartType = SeriesChartType.Line;
                        chart1.Series[fuels[i] + "_P"].Color = cl_p[i];
                        chart1.Series[fuels[i] + "_P"].BorderWidth = 2;

                        chart4.Series.Add(fuels[i] + "_ρ");
                        chart4.Series[fuels[i] + "_ρ"].YAxisType = AxisType.Primary;
                        chart4.Series[fuels[i] + "_ρ"].ChartType = SeriesChartType.Line;
                        chart4.Series[fuels[i] + "_ρ"].Color = cl_ro[i];
                        chart4.Series[fuels[i] + "_ρ"].BorderWidth = 2;

                        chart4.Series.Add(fuels[i] + "_T");
                        chart4.Series[fuels[i] + "_T"].YAxisType = AxisType.Secondary;
                        chart4.Series[fuels[i] + "_T"].ChartType = SeriesChartType.Line;
                        chart4.Series[fuels[i] + "_T"].Color = cl_T[i];
                        chart4.Series[fuels[i] + "_T"].BorderWidth = 2;

                        chart5.Series.Add(fuels[i] + "_ν");
                        chart5.Series[fuels[i] + "_ν"].YAxisType = AxisType.Primary;
                        chart5.Series[fuels[i] + "_ν"].ChartType = SeriesChartType.Line;
                        chart5.Series[fuels[i] + "_ν"].Color = cl_ro[i];
                        chart5.Series[fuels[i] + "_ν"].BorderWidth = 2;

                        chart5.Series.Add(fuels[i] + "_T");
                        chart5.Series[fuels[i] + "_T"].YAxisType = AxisType.Secondary;
                        chart5.Series[fuels[i] + "_T"].ChartType = SeriesChartType.Line;
                        chart5.Series[fuels[i] + "_T"].Color = cl_T[i];
                        chart5.Series[fuels[i] + "_T"].BorderWidth = 2;
                    }
                    textBox2.Text = fuels[num_fluid];
                    MessageBox.Show("Файл успешно считан!", "Чтение файла", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex) { MessageBox.Show("Ошибка: " + ex.Message, "Ошибка чтения файла", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            {
                if (ObjWorkBook != null)
                    ObjWorkBook.Close(false, "", null);
                ObjExcel.Quit();
                ObjWorkBook = null;
                ObjWorkSheet = null;
                ObjExcel = null;
            }

            if (check_fuels == true)
            {
                MessageBox.Show("Ошибка: ", "Ошибка чтения файла, переменные введены неверно", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        //Обновление данных НП
        private void button14_Click(object sender, EventArgs e)
        {
            chart1.Series.Clear();
            chart4.Series.Clear();

            chart1.Series.Add("z");
            chart1.Series["z"].ChartType = SeriesChartType.Spline;
            chart1.Series["z"].Color = z_c;
            chart1.Series["z"].BorderWidth = 2;

            check_fuels = false;
            double res = 0;
            for (int i = 0; i < dataGridView6.RowCount - 1; i++)
                for (int j = 0; j < 10; j++)
                    if ((Double.TryParse(dataGridView6.Rows[i].Cells[j].Value.ToString(), out res) == false && j != 3) || res < 0)
                        check_fuels = true;
            if (check_fuels == true)
                MessageBox.Show("Ошибка чтения данных, переменные введены неверно ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                for (int i = 0; i < dataGridView6.RowCount - 1; i++)
                {
                    ro[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[0].Value);
                    nu[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[1].Value);
                    T[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[2].Value);
                    fuels[i] = dataGridView6.Rows[i].Cells[3].Value.ToString();
                    mass[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[4].Value);
                    cp[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[5].Value);
                    lyam[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[6].Value);
                    a0[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[7].Value);
                    u0[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[8].Value);
                    T0[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[9].Value);
                    //Расчет
                    if (i == 0)
                    {
                        romax = ro[i];
                        romin = ro[i];

                        Tmax = T0[i];
                        Tmin = T0[i];

                        numax = nu[i];
                        numin = nu[i];
                    }
                    else
                    {
                        if (romax < ro[i])
                            romax = ro[i];
                        if (romax > ro[i])
                            romax = ro[i];

                        if (Tmax < T[i])
                            Tmax = T0[i];
                        if (Tmin > T[i])
                            Tmin = T0[i];

                        if (numax < nu[i])
                            numax = nu[i];
                        if (numin > nu[i])
                            numin = nu[i];
                    }

                    vol[i] = mass[i] * 1000 / ro[i];
                    cl_h[i] = Color.FromArgb(rnd.Next(0, 64), rnd.Next(128, 255), rnd.Next(0, 64));         //green
                    cl_p[i] = Color.FromArgb(rnd.Next(128, 255), rnd.Next(0, 64), rnd.Next(0, 64));         //red
                    cl_ro[i] = Color.FromArgb(rnd.Next(128, 255), rnd.Next(128, 255), rnd.Next(128, 255));  //any
                    cl_T[i] = Color.FromArgb(rnd.Next(0, 64), rnd.Next(0, 64), rnd.Next(128, 255));         //blue

                    chart1.Series.Add(fuels[i] + "_H");
                    chart1.Series[fuels[i] + "_H"].ChartType = SeriesChartType.Line;
                    chart1.Series[fuels[i] + "_H"].Color = cl_h[i];
                    chart1.Series[fuels[i] + "_H"].BorderWidth = 2;

                    chart1.Series.Add(fuels[i] + "_P");
                    chart1.Series[fuels[i] + "_P"].YAxisType = AxisType.Secondary;
                    chart1.Series[fuels[i] + "_P"].ChartType = SeriesChartType.Line;
                    chart1.Series[fuels[i] + "_P"].Color = cl_p[i];
                    chart1.Series[fuels[i] + "_P"].BorderWidth = 2;

                    chart4.Series.Add(fuels[i] + "_ρ");
                    chart4.Series[fuels[i] + "_ρ"].YAxisType = AxisType.Primary;
                    chart4.Series[fuels[i] + "_ρ"].ChartType = SeriesChartType.Line;
                    chart4.Series[fuels[i] + "_ρ"].Color = cl_ro[i];
                    chart4.Series[fuels[i] + "_ρ"].BorderWidth = 2;

                    chart4.Series.Add(fuels[i] + "_T");
                    chart4.Series[fuels[i] + "_T"].YAxisType = AxisType.Secondary;
                    chart4.Series[fuels[i] + "_T"].ChartType = SeriesChartType.Line;
                    chart4.Series[fuels[i] + "_T"].Color = cl_T[i];
                    chart4.Series[fuels[i] + "_T"].BorderWidth = 2;

                    chart5.Series.Add(fuels[i] + "_ν");
                    chart5.Series[fuels[i] + "_ν"].YAxisType = AxisType.Primary;
                    chart5.Series[fuels[i] + "_ν"].ChartType = SeriesChartType.Line;
                    chart5.Series[fuels[i] + "_ν"].Color = cl_ro[i];
                    chart5.Series[fuels[i] + "_ν"].BorderWidth = 2;

                    chart5.Series.Add(fuels[i] + "_T");
                    chart5.Series[fuels[i] + "_T"].YAxisType = AxisType.Secondary;
                    chart5.Series[fuels[i] + "_T"].ChartType = SeriesChartType.Line;
                    chart5.Series[fuels[i] + "_T"].Color = cl_T[i];
                    chart5.Series[fuels[i] + "_T"].BorderWidth = 2;
                }
                MessageBox.Show("Данные успешно обновлены!", "Данные", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        //Отрисовка температуры окружающей среды (грунта)
        private void button16_Click_1(object sender, EventArgs e)
        { 
            chart2.Series.Clear();
            chart2.Series.Add("T_грунта");
            chart2.Series["T_грунта"].YAxisType = AxisType.Primary;
            chart2.Series["T_грунта"].ChartType = SeriesChartType.Spline;
            chart2.Series["T_грунта"].Color = T_c;
            chart2.Series["T_грунта"].BorderWidth = 2;
            var chart = chart2.ChartAreas[0];

            foreach (var series in chart2.Series)
            {
                series.Points.Clear();
            }

            chart2.Legends[0].Enabled = true;
            chart2.Legends[0].Font = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart2.ChartAreas[0].AxisY.Title = "T, K";
            chart2.ChartAreas[0].AxisY.TitleFont = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart2.ChartAreas[0].AxisY.MajorGrid.LineColor = cl_g;
            chart2.ChartAreas[0].AxisY.MinorGrid.LineColor = cl_g;
            chart2.ChartAreas[0].AxisY.LineColor = cl_g;
            chart2.ChartAreas[0].AxisY.LabelStyle.ForeColor = Color.White;
            chart2.ChartAreas[0].AxisY.MajorTickMark.LineColor = cl_g;
            chart2.ChartAreas[0].AxisY.MinorTickMark.LineColor = cl_g;
            chart2.ChartAreas[0].AxisY.LabelStyle.Format = "0.0";
            chart2.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Arial", 9f, FontStyle.Bold);

            chart2.ChartAreas[0].AxisX.Title = "x, км";
            chart2.ChartAreas[0].AxisX.TitleFont = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart2.ChartAreas[0].AxisX.MajorGrid.LineColor = cl_g;
            chart2.ChartAreas[0].AxisX.MinorGrid.LineColor = cl_g;
            chart2.ChartAreas[0].AxisX.LineColor = cl_g;
            chart2.ChartAreas[0].AxisX.LabelStyle.ForeColor = Color.White;
            chart2.ChartAreas[0].AxisX.MajorTickMark.LineColor = cl_g;
            chart2.ChartAreas[0].AxisX.MinorTickMark.LineColor = cl_g;
            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0.00";
            chart2.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Arial", 9f, FontStyle.Bold);

            chart.AxisX.Minimum = 0;
            chart.AxisX.Maximum = Convert.ToDouble(dataGridView5.Rows[dataGridView5.Rows.Count - 2].Cells[0].Value);
            chart2.ChartAreas[0].AxisX.Interval = chart.AxisX.Maximum / 15;

            chart.AxisY.Minimum = Tvnmin - 5;
            chart.AxisY.Maximum = Tvnmax + 5;
            chart2.ChartAreas[0].AxisY.Interval = (Tvnmax - Tvnmin) / 15;
            
            for (int i = 0; i < dataGridView5.RowCount - 1; i++)
            { 
                chart2.Series["T_грунта"].Points.AddXY(m_x[i], m_Tvn[i]);
            }

        }
        //Отрисовка температуры стенки трубопровода
        private void button20_Click_1(object sender, EventArgs e)
        {
            chart3.Series.Clear();
            chart3.Series.Add("T_трубы");
            chart3.Series["T_трубы"].YAxisType = AxisType.Primary;
            chart3.Series["T_трубы"].ChartType = SeriesChartType.Spline;
            chart3.Series["T_трубы"].Color = T_c;
            chart3.Series["T_трубы"].BorderWidth = 2;
            var chart = chart3.ChartAreas[0];

            foreach (var series in chart3.Series)
            {
                series.Points.Clear();
            }

            chart3.Legends[0].Enabled = true;

            chart3.Legends[0].Font = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart3.ChartAreas[0].AxisY.Title = "T, K";
            chart3.ChartAreas[0].AxisY.TitleFont = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart3.ChartAreas[0].AxisY.MajorGrid.LineColor = cl_g;
            chart3.ChartAreas[0].AxisY.MinorGrid.LineColor = cl_g;
            chart3.ChartAreas[0].AxisY.LineColor = cl_g;
            chart3.ChartAreas[0].AxisY.LabelStyle.ForeColor = Color.White;
            chart3.ChartAreas[0].AxisY.MajorTickMark.LineColor = cl_g;
            chart3.ChartAreas[0].AxisY.MinorTickMark.LineColor = cl_g;
            chart3.ChartAreas[0].AxisY.LabelStyle.Format = "0.0";
            chart3.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Arial", 9f, FontStyle.Bold);

            chart3.ChartAreas[0].AxisX.Title = "x, км";
            chart3.ChartAreas[0].AxisX.TitleFont = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart3.ChartAreas[0].AxisX.MajorGrid.LineColor = cl_g;
            chart3.ChartAreas[0].AxisX.MinorGrid.LineColor = cl_g;
            chart3.ChartAreas[0].AxisX.LineColor = cl_g;
            chart3.ChartAreas[0].AxisX.LabelStyle.ForeColor = Color.White;
            chart3.ChartAreas[0].AxisX.MajorTickMark.LineColor = cl_g;
            chart3.ChartAreas[0].AxisX.MinorTickMark.LineColor = cl_g;
            chart3.ChartAreas[0].AxisX.LabelStyle.Format = "0.00";
            chart3.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Arial", 9f, FontStyle.Bold);

            chart.AxisX.Minimum = 0;
            chart.AxisX.Maximum = Convert.ToDouble(dataGridView5.Rows[dataGridView5.Rows.Count - 2].Cells[0].Value);
            chart3.ChartAreas[0].AxisX.Interval = chart.AxisX.Maximum / 15;

            chart.AxisY.Minimum = Tstmin - 5;
            chart.AxisY.Maximum = Tstmax + 5;
            chart3.ChartAreas[0].AxisY.Interval = (Tstmax - Tstmin) / 15;

            for (int i = 0; i < dataGridView5.RowCount - 1; i++)
            {
                chart3.Series["T_трубы"].Points.AddXY(m_x[i], m_Tst[i]);
            }
        }
        //Температурный расчет (изменение температуры нефтепродукта по течению вдоль оси трубопровода)
        private void button17_Click(object sender, EventArgs e)
        {
            Qm = 0;
            for (int m = 0; m < n; m++)
            {
                x[m] = 0;
                z[m] = 0;
                a2[m] = 0;
                v[m] = 0;
                eps_sp[m] = 0;
                for (int j = 0; j < 1000; j++)
                {
                    cp_st[j, m] = 0;
                    ro_st[j, m] = 0;
                    nu_st[j, m] = 0;
                    lyam_st[j, m] = 0;
                    pr_st[j, m] = 0;

                    cp_rl[j, m] = 0;
                    ro_rl[j, m] = 0;
                    nu_rl[j, m] = 0;
                    lyam_rl[j, m] = 0;
                    pr_rl[j, m] = 0;


                    a1[j, m] = 0;
                    k_sp[j, m] = 0;
                    Shu[j, m] = 0;
                    Re[j, m] = 0;

                    lyam_sp[j, m] = 0;
                    i_sp[j, m] = 0;
                    b_sp[j, m] = 0;
                    Tk[j, m] = 0;
                }
            }

            dataGridView8.Rows.Clear();
            //double x = m_x[0], deltax = Convert.ToDouble(textBox1.Text);
            //double a2 = 0, cp_st = 0, ro_st = 0, nu_st = 0, lyam_st = 0, pr_st = 0;
            //double cp_rl = 0, ro_rl = 0, nu_rl = 0, lyam_rl = 0, pr_rl = 0;
            //double v = 0, a1 = 0, k = 0, Qm = 0, Shu = 0, Re = 0, eps = 0;
            //double lyam_sp = 0, i_sp = 0, b = 0, Tk = 0;
            int i = 0;
            deltax = (m_x[dataGridView5.Rows.Count - 2] - m_x[0]) / (n - 1);
            textBox1.Text = deltax.ToString();
            //int num_loc = num_fluid;
            for (int num_loc = 0; num_loc < dataGridView6.Rows.Count - 1; num_loc++)
            {
                x[0] = m_x[0];
                z[0] = m_z[0];
                Tvn[0] = m_Tvn[0];
                //Convert.ToDouble(textBox1.Text);
                //a2 = 0; cp_st = 0; ro_st = 0; nu_st = 0; lyam_st = 0; pr_st = 0;
                //cp_rl = 0; ro_rl = 0; nu_rl = 0; lyam_rl = 0; pr_rl = 0;
                //v = 0; a1 = 0; k = 0; Qm = 0; Shu = 0; Re = 0; eps = 0;
                //lyam_sp = 0; i_sp = 0; b = 0; Tk = 0;
                //do
                for (int j = 0; j < n; j++)
                {
                    if (j != 0)
                        x[j] = x[j - 1] + deltax;
                    //m_ps[i] = m_ps[i] + 1;
                    for (i = 0; i < dataGridView5.Rows.Count - 2; i++)
                        if (x[j] >= m_x[i] && x[j] < m_x[i + 1])
                        {
                            break;
                        }
                    if (i != dataGridView5.Rows.Count - 2)
                    {
                        z[j] = m_z[i] + (m_z[i + 1] - m_z[i]) / (m_x[i + 1] - m_x[i]) * (x[j] - m_x[i]);
                        Tvn[j] = m_Tvn[i] + (m_Tvn[i + 1] - m_Tvn[i]) / (m_x[i + 1] - m_x[i]) * (x[j] - m_x[i]);
                    }
                    else
                    { 
                        z[j] = m_z[i];
                        Tvn[j] = m_Tvn[i];
                    }
                    //if (x[j] >= m_x[i] && x[j - 1] < m_x[i])
                    if (x[j] == m_x[i])
                        z[j] = m_z[i];
                    D[j] = m_D[i];
                    sig[j] = m_sig[i];

                    a2[j] = 2 * m_lyam_gr[i] / (D[j] / 1000 * Math.Log(2 * m_h0[i] * 1000 / D[j] + Math.Sqrt(Math.Pow(m_h0[i] * 1000 / D[j], 2) - 1)));
                    cp_st[num_loc, j] = (1.5072 + (m_Tst[i] - 223) / 100 * (1.7182 - 1.5072 * ro[num_loc] / 1000)) * 1000;
                    ro_st[num_loc, j] = ro[num_loc] + a0[num_loc] * (293.15 - m_Tst[i]);
                    nu_st[num_loc, j] = nu[num_loc] * Math.Exp(-u0[num_loc] * (m_Tst[i] - 293.15));
                    lyam_st[num_loc, j] = lyam[num_loc] * (1 - a0[num_loc] * (m_Tst[i] - 293.15));
                    pr_st[num_loc, j] = cp_st[num_loc, j] * ro_st[num_loc, j] * nu_st[num_loc, j] / lyam_st[num_loc, j];
                    //
                    if (x[j] == m_x[0])
                    {
                        cp_rl[num_loc, j] = (1.5072 + (T0[num_loc] - 223) / 100 * (1.7182 - 1.5072 * ro[num_loc] / 1000)) * 1000;
                        ro_rl[num_loc, j] = ro[num_loc] + a0[num_loc] * (293.15 - T0[num_loc]);
                        nu_rl[num_loc, j] = nu[num_loc] * Math.Exp(-u0[num_loc] * (T0[num_loc] - 293.15));
                        lyam_rl[num_loc, j] = lyam[num_loc] * (1 - a0[num_loc] * (T0[num_loc] - 293.15));
                    }
                    else
                    {
                        cp_rl[num_loc, j] = (1.5072 + (Tk[num_loc, j - 1] - 223) / 100 * (1.7182 - 1.5072 * ro[num_loc] / 1000)) * 1000;
                        ro_rl[num_loc, j] = ro[num_loc] + a0[num_loc] * (293.15 - Tk[num_loc, j - 1]);
                        nu_rl[num_loc, j] = nu[num_loc] * Math.Exp(-u0[num_loc] * (Tk[num_loc, j - 1] - 293.15));
                        lyam_rl[num_loc, j] = lyam[num_loc] * (1 - a0[num_loc] * (Tk[num_loc, j - 1] - 293.15));
                    }
                    pr_rl[num_loc, j] = cp_rl[num_loc, j] * ro_rl[num_loc, j] * nu_rl[num_loc, j] / lyam_rl[num_loc, j];
                    //
                    v[j] = Q * 4 / (pi * Math.Pow(((D[j] - 2 * sig[j]) / 1000), 2));
                    a1[num_loc, j] = 0.021 * Math.Pow(v[j], 0.8) * Math.Pow(cp_rl[num_loc, j], 0.68) * Math.Pow(ro_rl[num_loc, j], 0.68) * Math.Pow(lyam_rl[num_loc, j], 0.32)
                        * 1 / (Math.Pow(((D[j] - 2 * sig[j]) / 1000), 0.2) * Math.Pow(nu_rl[num_loc, j], 0.12)) * Math.Pow(1 / pr_st[num_loc, j], 0.25);
                    k_sp[num_loc, j] = 1 / (1000 / (a1[num_loc, j] * (D[j] - 2 * sig[j])) + 1000 / (a2[j] * D[j]) + Math.Log(D[j] / (D[j] - 2 * sig[j])) / (2 * m_lyam_tr[i]));
                    
                    Qm = Q * ro[num_fluid];
                    Shu[num_loc, j] = pi * k_sp[num_loc, j] * (D[j] - 2 * sig[j]) / 1000 * x[j] * 1000 / (Qm * cp_rl[num_loc, j]);
                    Re[num_loc, j] = v[j] * (D[j] - 2 * sig[j]) / (nu_rl[num_loc, j] * 1000);

                    eps_sp[j] = m_del[i]  / (D[j] - 2 * sig[j]);
                    //
                    if (Re[num_loc, j] < 10 / eps_sp[j])
                        lyam_sp[num_loc, j] = 0.3164 / Math.Pow(Re[num_loc, j], 0.25);
                    else
                        if (Re[num_loc, j] >= 10 / eps_sp[j] && Re[num_loc, j] < 500 / eps_sp[j])
                        lyam_sp[num_loc, j] = 0.11 * Math.Pow((eps_sp[j] + 68 / Re[num_loc, j]), 0.25);
                    else
                        if (Re[num_loc, j] >= 500 / eps_sp[j])
                        lyam_sp[num_loc, j] = 0.11 * Math.Pow(eps_sp[j], 0.25);
                    i_sp[num_loc, j] = lyam_sp[num_loc, j] * 8 * Math.Pow(1000 / (D[j] - 2 * sig[j]), 5) * Math.Pow(Q, 2) / (Math.Pow(pi, 2) * g);

                    b_sp[num_loc, j] = Q / ro_rl[num_loc, j] * i_sp[num_loc, j] / (pi * k_sp[num_loc, j] * (D[j] - 2 * sig[j]) / 1000);
                    Tk[num_loc, j] = Tvn[j] + (T0[num_fluid] - Tvn[j]) * Math.Exp(-Shu[num_loc, j]) + b * (1 - Math.Exp(-Shu[num_loc, j]));
                    //
                    if (num_loc == num_fluid)
                        dataGridView8.Rows.Add(x[j].ToString("0.00"), D[j], sig[j],
                                                Tvn[j], m_lyam_gr[i], m_lyam_tr[i], m_h0[i], m_Tst[i],
                                                a2[j], cp_st[num_loc, j], ro_st[num_loc, j], nu_st[num_loc, j], lyam_st[num_loc, j], pr_st[num_loc, j],
                                                cp_rl[num_loc, j], ro_rl[num_loc, j], nu_rl[num_loc, j], lyam_rl[num_loc, j], pr_rl[num_loc, j],
                                                v[j], a1[num_loc, j], k_sp[num_loc, j], Shu[num_loc, j], Re[num_loc, j], eps_sp[j], lyam_sp[num_loc, j], i_sp[num_loc, j], b_sp[num_loc, j], Tk[num_loc, j]);
                }
            }
        }
        //Переключение номера высчитываемого нефтепродукта
        private void button18_Click(object sender, EventArgs e)
        {
            if (num_fluid > 0)
            {
                num_fluid--;
                textBox2.Text = fuels[num_fluid];
                dataGridView8.Rows.Clear();
                int i = 0;
                for (int j = 0; j < n; j++)
                {
                    for (i = 0; i < dataGridView5.Rows.Count - 2; i++)
                        if (x[j] >= m_x[i] && x[j] < m_x[i + 1])
                        {
                            break;
                        }
                    dataGridView8.Rows.Add(x[j].ToString("0.00"), m_D[i], m_sig[i],
                        Tvn[j], m_lyam_gr[i], m_lyam_tr[i], m_h0[i], m_Tst[i],
                        a2[j], cp_st[num_fluid, j], ro_st[num_fluid, j], nu_st[num_fluid, j], lyam_st[num_fluid, j], pr_st[num_fluid, j],
                        cp_rl[num_fluid, j], ro_rl[num_fluid, j], nu_rl[num_fluid, j], lyam_rl[num_fluid, j], pr_rl[num_fluid, j],
                     v[j], a1[num_fluid, j], k_sp[num_fluid, j], Shu[num_fluid, j], Re[num_fluid, j], eps_sp[j], lyam_sp[num_fluid, j], i_sp[num_fluid, j], b_sp[num_fluid, j], Tk[num_fluid, j]);
            
                }
            }
        }
        //Переключение номера высчитываемого нефтепродукта
        private void button19_Click(object sender, EventArgs e)
        {
            if (num_fluid < dataGridView6.Rows.Count - 2)
            {
                num_fluid++;
                textBox2.Text = fuels[num_fluid];
                dataGridView8.Rows.Clear();
                int i = 0;
                for (int j = 0; j < n; j++)
                {
                    for (i = 0; i < dataGridView5.Rows.Count - 2; i++)
                        if (x[j] >= m_x[i] && x[j] < m_x[i + 1])
                        {
                            break;
                        }
                    dataGridView8.Rows.Add(x[j].ToString("0.00"), m_D[i], m_sig[i],
                        Tvn[j], m_lyam_gr[i], m_lyam_tr[i], m_h0[i], m_Tst[i],
                        a2[j], cp_st[num_fluid, j], ro_st[num_fluid, j], nu_st[num_fluid, j], lyam_st[num_fluid, j], pr_st[num_fluid, j],
                        cp_rl[num_fluid, j], ro_rl[num_fluid, j], nu_rl[num_fluid, j], lyam_rl[num_fluid, j], pr_rl[num_fluid, j],
                     v[j], a1[num_fluid, j], k_sp[num_fluid, j], Shu[num_fluid, j], Re[num_fluid, j], eps_sp[j], lyam_sp[num_fluid, j], i_sp[num_fluid, j], b_sp[num_fluid, j], Tk[num_fluid, j]);

                }
            }
        }
        //Расчет местоположения жидкостей в трубопроводе
        private void button22_Click(object sender, EventArgs e)
        {
            int i = 0;

            for (i = 0; i < dataGridView6.Rows.Count - 1; i++)
                mass[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[4].Value);
            dataGridView9.Rows.Clear();
            
            x_f[dataGridView6.Rows.Count - 2] = m_x[dataGridView5.Rows.Count - 2];
            x_s[dataGridView6.Rows.Count - 2] = x_f[dataGridView6.Rows.Count - 2];

            i = dataGridView6.Rows.Count - 2;
            int j = 0;

            for (j = n - 1; j >= 0; j--)
                if (mass[i] - ro_rl[i, j] * (deltax) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2) >= 0)
                {
                    if (j == 0)
                        MessageBox.Show("true", "Чтение файла", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (x_s[i] - deltax >= 0)
                    { 
                        mass[i] = mass[i] - ro_rl[i, j] * (deltax) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2);
                        x_s[i] = x_s[i] - deltax;
                    }
                    else
                    {
                        MessageBox.Show("true", "Чтение файла", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        x_s[i] = 0;
                        mass[i] = mass[i] - ro_rl[i, j] * (x_s[i]) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2);
                    }
                }
                else
                {
                    if (x_s[i] - (mass[i]) / (ro_rl[i, j] * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2)) >= 0)
                        x_s[i] = x_s[i] - (deltax) * (mass[i]) / (ro_rl[i, j] * (deltax) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2));
                    else
                        x_s[i] = 0;
                    if (i > 0)
                    { 
                        i--;
                        x_f[i] = x_s[i + 1];
                        x_s[i] = x_f[i] - (deltax) * (1 - mass[i + 1] / (ro_rl[i + 1, j] * (deltax) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2)));
                        mass[i + 1] = 0;
                        mass[i] = mass[i] - ro_rl[i, j] * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2) * ((deltax) * (1 - mass[i + 1] / (ro_rl[i + 1, j] * (deltax) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2))));
                    }
                }

                for (i = dataGridView6.Rows.Count - 2; i >= 0; i--)
                {
                    //mass[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[4].Value);
                    dataGridView9.Rows.Add(fuels[i], mass[i], vol[i], 
                    x_s[i], x_f[i], 0, 0, 0, 0);
                }

            for (i = 0; i < dataGridView6.Rows.Count - 1; i++)
                mass[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[4].Value);
            //Определяем первую закачиваемую в трубопровод жидкость
            for (i = 0; i <= dataGridView6.Rows.Count - 2; i++)
                if (x_f[i] == m_x[dataGridView5.Rows.Count - 2] && mass[i] > 0)
                {
                    key_p = i;
                    break;
                }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        //Итоговый расчет
        private void button21_Click(object sender, EventArgs e)
        {
            int kol = dataGridView6.Rows.Count - 2;
            dataGridView7.Rows.Clear();
            double H = 0, P = 0;
            //Определяем последнюю закачиваемую в трубопровод жидкость
            int i = 0;
            for (i = 0; i < dataGridView6.Rows.Count - 1; i++)
                if (x_s[i] == 0 && mass[i] > 0)
                        break;
            int l = i;
            for (int j = 0; j < n ; j++)
            {
                if (x_s[l] <= x[j] && x_f[l] >= x[j])
                {
                    if (j == 0)
                    {
                        H = Convert.ToDouble(textBox14.Text);
                    }
                    else
                    {
                        //MessageBox.Show( (dataGridView7.Rows[j - 1].Cells[10].Value).ToString(), "Чтение файла", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        H = H - i_sp[l, j - 1] * deltax * 1000;
                    }
                    P = (H - Math.Pow(v[j], 2) / (2 * g) - z[j]) * ro_rl[l, j] * g;

                    if (x[j] == x_s[l])
                    {
                        dataGridView9.Rows[kol - l].Cells[5].Value = H.ToString();
                        dataGridView9.Rows[kol - l].Cells[7].Value = (P * Math.Pow(10, -6)).ToString();
                    }
                    dataGridView7.Rows.Add(x[j].ToString("0.0000"), z[j], D[j], sig[j], Tk[l, j],
                        ro_rl[l, j], nu_rl[l, j], v[j], eps_sp[j], Re[l, j], lyam_sp[l, j], i_sp[l, j], H, (P * Math.Pow(10, -6)));  
                }
                else 
                {
                    if (j != n - 1)
                        l++;
                    if (x_s[l] <= x[j] && x_f[l - 1] < x[j] && j != n - 1)
                    {
                        H = H - i_sp[l, j - 1] * (x[j] - x_f[l - 1]) * 1000;
                        P = (H - Math.Pow(v[j], 2) / (2 * g) - z[j]) * ro_rl[l - 1, j] * g;

                        dataGridView9.Rows[kol + 1 -l].Cells[6].Value = H.ToString();
                        dataGridView9.Rows[kol + 1 - l].Cells[8].Value = (P * Math.Pow(10, -6)).ToString();

                        double zs = z[j - 1] + (z[j] - z[j - 1]) / (x[j] - x[j - 1]) * (x_f[l - 1] - x[j - 1]);
                        dataGridView7.Rows.Add(x_f[l - 1].ToString("0.0000"), zs, D[j], sig[j], Tk[l, j],
                            ro_rl[l - 1, j], nu_rl[l - 1, j], v[j], eps_sp[j], Re[l - 1, j], lyam_sp[l - 1, j], i_sp[l - 1, j], H, (P * Math.Pow(10, -6)));

                        P = P;
                        H = P / (ro_rl[l, j] * g) + Math.Pow(v[j], 2) / (2 * g) + z[j];

                        dataGridView9.Rows[kol - l].Cells[5].Value = H.ToString();
                        dataGridView9.Rows[kol - l].Cells[7].Value = (P * Math.Pow(10, -6)).ToString();

                        dataGridView7.Rows.Add(x_s[l].ToString("0.0000"), zs, D[j], sig[j], Tk[l, j],
                            ro_rl[l, j], nu_rl[l, j], v[j], eps_sp[j], Re[l, j], lyam_sp[l, j], i_sp[l, j], H, (P * Math.Pow(10, -6)));
                    }

                    if (j != n - 1)
                    {
                        H = H - i_sp[l, j - 1] * (x[j] - x_s[l]) * 1000;
                        P = (H - Math.Pow(v[j], 2) / (2 * g) - z[j]) * ro_rl[l, j] * g;
                    }
                    else
                    {
                        H = H - i_sp[l, j - 1] * (deltax) * 1000;
                        P = (H - Math.Pow(v[j], 2) / (2 * g) - z[j]) * ro_rl[l, j] * g;
                        dataGridView9.Rows[kol - l].Cells[6].Value = H.ToString();
                        dataGridView9.Rows[kol - l].Cells[8].Value = (P * Math.Pow(10, -6)).ToString();
                    }
                    dataGridView7.Rows.Add(x[j].ToString("0.0000"), z[j], D[j], sig[j], Tk[l, j],
                        ro_rl[l, j], nu_rl[l, j], v[j], eps_sp[j], Re[l, j], lyam_sp[l, j], i_sp[l, j], H, (P * Math.Pow(10, -6)));
                }
            }
        }
        
        //График H(x)-P(x)
        private void button24_Click_1(object sender, EventArgs e)
        {
            double x_rl = 0;
            //Определяем последнюю закачиваемую в трубопровод жидкость
            int i = 0;
            for (i = 0; i < dataGridView6.Rows.Count - 1; i++)
                if (x_s[i] == 0 && mass[i] > 0)
                    break;
            int l = i;

            foreach (var series in chart1.Series)
            {
                series.Points.Clear();
            }

            var chart = chart1.ChartAreas[0];

            chart1.Legends[0].Enabled = true;
            chart1.Legends[0].Font = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart1.ChartAreas[0].AxisY2.Title = "P, Мпа";
            chart1.ChartAreas[0].AxisY2.TitleFont = new System.Drawing.Font("Arial", 10f, FontStyle.Bold);
            chart1.ChartAreas[0].AxisY2.MajorGrid.LineColor = cl_g;
            chart1.ChartAreas[0].AxisY2.MinorGrid.LineColor = cl_g;
            chart1.ChartAreas[0].AxisY2.LineColor = cl_g;
            chart1.ChartAreas[0].AxisY2.LabelStyle.ForeColor = Color.White;
            chart1.ChartAreas[0].AxisY2.MajorTickMark.LineColor = cl_g;
            chart1.ChartAreas[0].AxisY2.MinorTickMark.LineColor = cl_g;
            chart1.ChartAreas[0].AxisY2.LabelStyle.Format = "0.00";
            chart1.ChartAreas[0].AxisY2.LabelStyle.Font = new System.Drawing.Font("Arial", 9f, FontStyle.Bold);

            chart1.ChartAreas[0].AxisY.Title = "H, м";
            chart1.ChartAreas[0].AxisY.TitleFont = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart1.ChartAreas[0].AxisY.MajorGrid.LineColor = cl_g;
            chart1.ChartAreas[0].AxisY.MinorGrid.LineColor = cl_g;
            chart1.ChartAreas[0].AxisY.LineColor = cl_g;
            chart1.ChartAreas[0].AxisY.LabelStyle.ForeColor = Color.White;
            chart1.ChartAreas[0].AxisY.MajorTickMark.LineColor = cl_g;
            chart1.ChartAreas[0].AxisY.MinorTickMark.LineColor = cl_g;
            chart1.ChartAreas[0].AxisY.LabelStyle.Format = "0.0";
            chart1.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Arial", 9f, FontStyle.Bold);

            chart1.ChartAreas[0].AxisX.Title = "x, км";
            chart1.ChartAreas[0].AxisX.TitleFont = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart1.ChartAreas[0].AxisX.MajorGrid.LineColor = cl_g;
            chart1.ChartAreas[0].AxisX.MinorGrid.LineColor = cl_g;
            chart1.ChartAreas[0].AxisX.LineColor = cl_g;
            chart1.ChartAreas[0].AxisX.LabelStyle.ForeColor = Color.White;
            chart1.ChartAreas[0].AxisX.MajorTickMark.LineColor = cl_g;
            chart1.ChartAreas[0].AxisX.MinorTickMark.LineColor = cl_g;
            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0.00";
            chart1.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Arial", 9f, FontStyle.Bold);

            chart.AxisX.Minimum = 0;
            chart.AxisX.Maximum = Convert.ToDouble(dataGridView5.Rows[dataGridView5.Rows.Count - 2].Cells[0].Value);
            chart1.ChartAreas[0].AxisX.Interval = chart.AxisX.Maximum / 15;

            chart.AxisY.Minimum = 0;
            chart.AxisY.Maximum = Convert.ToDouble(dataGridView7.Rows[0].Cells[12].Value) + 200;
            chart1.ChartAreas[0].AxisY.Interval = chart.AxisY.Maximum / 15;


            chart.AxisY2.Minimum = 0;
            chart.AxisY2.Maximum = Convert.ToDouble(dataGridView7.Rows[0].Cells[13].Value) + 2;
            chart1.ChartAreas[0].AxisY2.Interval = chart.AxisY2.Maximum / 15;

            for (int j = 0; j < dataGridView7.Rows.Count - 1; j++)
            {
                x_rl = Convert.ToDouble(dataGridView7.Rows[j].Cells[0].Value);
                if (x_s[l] <= x_rl && x_f[l] >= x_rl)
                {
                    chart1.Series[fuels[l] + "_H"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[12].Value));
                    chart1.Series[fuels[l] + "_P"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[13].Value));
                    chart1.Series["z"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[1].Value));
                }
                else
                {
                    chart1.Series[fuels[l] + "_H"].Points.AddXY(x_f[l], Convert.ToDouble(dataGridView7.Rows[j].Cells[12].Value));
                    chart1.Series[fuels[l] + "_P"].Points.AddXY(x_f[l], Convert.ToDouble(dataGridView7.Rows[j].Cells[13].Value));
                    chart1.Series["z"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[1].Value));

                    if (j != n - 1)
                    {
                        l++;
                        chart1.Series[fuels[l] + "_H"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[12].Value));
                        chart1.Series[fuels[l] + "_P"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[13].Value));
                        chart1.Series["z"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[1].Value));
                    }
                }
            }

            button26.Enabled = true;
            textBox3.Text = (chart.AxisX.Minimum).ToString();
            textBox4.Text = (chart.AxisX.Maximum).ToString();

            textBox26.Text = (chart.AxisY.Minimum).ToString();
            textBox25.Text = (chart.AxisY.Maximum).ToString();

            textBox28.Text = (chart.AxisY2.Minimum).ToString();
            textBox27.Text = (chart.AxisY2.Maximum).ToString();

            textBox35.Text = chart1.ChartAreas[0].AxisX.Interval.ToString();
            textBox36.Text = chart1.ChartAreas[0].AxisY.Interval.ToString();
            textBox37.Text = chart1.ChartAreas[0].AxisY2.Interval.ToString();

            textBox3.Enabled = true;
            textBox4.Enabled = true;
            textBox26.Enabled = true;
            textBox25.Enabled = true;
            textBox28.Enabled = true;
            textBox27.Enabled = true;
            textBox35.Enabled = true;
            textBox36.Enabled = true;
            textBox37.Enabled = true;
        }
        //График ρ(T(x))
        private void button23_Click_1(object sender, EventArgs e)
        {
            double x_rl = 0;
            //Определяем последнюю закачиваемую в трубопровод жидкость
            int i = 0;
            for (i = 0; i < dataGridView6.Rows.Count - 1; i++)
                if (x_s[i] == 0 && mass[i] > 0)
                    break;
            int l = i;

            foreach (var series in chart4.Series)
            {
                series.Points.Clear();
            }

            var chart = chart4.ChartAreas[0];
            chart4.Legends[0].Enabled = true;

            chart4.ChartAreas[0].AxisY.Title = "ρ, кг/м^3";
            chart4.ChartAreas[0].AxisY.TitleFont = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart4.ChartAreas[0].AxisY.MajorGrid.LineColor = cl_g;
            chart4.ChartAreas[0].AxisY.MinorGrid.LineColor = cl_g;
            chart4.ChartAreas[0].AxisY.LineColor = cl_g;
            chart4.ChartAreas[0].AxisY.LabelStyle.ForeColor = Color.White;
            chart4.ChartAreas[0].AxisY.MajorTickMark.LineColor = cl_g;
            chart4.ChartAreas[0].AxisY.MinorTickMark.LineColor = cl_g;
            chart4.ChartAreas[0].AxisY.LabelStyle.Format = "0.0";
            chart4.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Arial", 9f, FontStyle.Bold);

            chart4.ChartAreas[0].AxisX.Title = "x, км";
            chart4.ChartAreas[0].AxisX.TitleFont = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart4.ChartAreas[0].AxisX.MajorGrid.LineColor = cl_g;
            chart4.ChartAreas[0].AxisX.MinorGrid.LineColor = cl_g;
            chart4.ChartAreas[0].AxisX.LineColor = cl_g;
            chart4.ChartAreas[0].AxisX.LabelStyle.ForeColor = Color.White;
            chart4.ChartAreas[0].AxisX.MajorTickMark.LineColor = cl_g;
            chart4.ChartAreas[0].AxisX.MinorTickMark.LineColor = cl_g;
            chart4.ChartAreas[0].AxisX.LabelStyle.Format = "0.00";
            chart4.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Arial", 9f, FontStyle.Bold);

            chart4.Legends[0].Font = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart4.ChartAreas[0].AxisY2.Title = "T, K";
            chart4.ChartAreas[0].AxisY2.TitleFont = new System.Drawing.Font("Arial", 10f, FontStyle.Bold);
            chart4.ChartAreas[0].AxisY2.MajorGrid.LineColor = cl_g;
            chart4.ChartAreas[0].AxisY2.MinorGrid.LineColor = cl_g;
            chart4.ChartAreas[0].AxisY2.LineColor = cl_g;
            chart4.ChartAreas[0].AxisY2.LabelStyle.ForeColor = Color.White;
            chart4.ChartAreas[0].AxisY2.MajorTickMark.LineColor = cl_g;
            chart4.ChartAreas[0].AxisY2.MinorTickMark.LineColor = cl_g;
            chart4.ChartAreas[0].AxisY2.LabelStyle.Format = "0.00";
            chart4.ChartAreas[0].AxisY2.LabelStyle.Font = new System.Drawing.Font("Arial", 9f, FontStyle.Bold);

            chart.AxisX.Minimum = 0;
            chart.AxisX.Maximum = Convert.ToDouble(dataGridView5.Rows[dataGridView5.Rows.Count - 2].Cells[0].Value);
            chart4.ChartAreas[0].AxisX.Interval = chart.AxisX.Maximum / 15;

            chart.AxisY.Minimum = romin - 10;
            chart.AxisY.Maximum = romax + 10;
            chart4.ChartAreas[0].AxisY.Interval = (romax-romin + 100) / 15;

            chart.AxisY2.Minimum = Tmin - deltaT;
            chart.AxisY2.Maximum = Tmax + deltaT;
            chart4.ChartAreas[0].AxisY2.Interval = (Tmax- Tmin + 2 * deltaT) / 15;

            for (int j = 0; j < dataGridView7.Rows.Count - 1; j++)
            {
                x_rl = Convert.ToDouble(dataGridView7.Rows[j].Cells[0].Value);
                if (x_s[l] <= x_rl && x_f[l] >= x_rl)
                {
                    chart4.Series[fuels[l] + "_ρ"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[5].Value));
                    chart4.Series[fuels[l] + "_T"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[4].Value));
                }
                else
                {
                    if (j != n - 1)
                    {
                        l++;
                        chart4.Series[fuels[l] + "_ρ"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[5].Value));
                        chart4.Series[fuels[l] + "_T"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[4].Value));
                    }
                }
            }

            button27.Enabled = true;
            textBox34.Text = (chart.AxisX.Minimum).ToString();
            textBox33.Text = (chart.AxisX.Maximum).ToString();

            textBox32.Text = (chart.AxisY.Minimum).ToString();
            textBox31.Text = (chart.AxisY.Maximum).ToString();

            textBox30.Text = (chart.AxisY2.Minimum).ToString();
            textBox29.Text = (chart.AxisY2.Maximum).ToString();

            textBox40.Text = chart4.ChartAreas[0].AxisX.Interval.ToString();
            textBox39.Text = chart4.ChartAreas[0].AxisY.Interval.ToString();
            textBox38.Text = chart4.ChartAreas[0].AxisY2.Interval.ToString("0.00");

            textBox34.Enabled = true;
            textBox33.Enabled = true;
            textBox32.Enabled = true;
            textBox31.Enabled = true;
            textBox30.Enabled = true;
            textBox29.Enabled = true;
            textBox40.Enabled = true;
            textBox39.Enabled = true;
            textBox38.Enabled = true;
        }
        //График nu(x)
        private void button25_Click_1(object sender, EventArgs e)
        {
            double x_rl = 0;
            //Определяем последнюю закачиваемую в трубопровод жидкость
            int i = 0;
            for (i = 0; i < dataGridView6.Rows.Count - 1; i++)
                if (x_s[i] == 0 && mass[i] > 0)
                    break;
            int l = i;

            foreach (var series in chart5.Series)
            {
                series.Points.Clear();
            }

            var chart = chart5.ChartAreas[0];
            chart5.Legends[0].Enabled = true;

            chart5.ChartAreas[0].AxisY.Title = "ν, сСт";
            chart5.ChartAreas[0].AxisY.TitleFont = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart5.ChartAreas[0].AxisY.MajorGrid.LineColor = cl_g;
            chart5.ChartAreas[0].AxisY.MinorGrid.LineColor = cl_g;
            chart5.ChartAreas[0].AxisY.LineColor = cl_g;
            chart5.ChartAreas[0].AxisY.LabelStyle.ForeColor = Color.White;
            chart5.ChartAreas[0].AxisY.MajorTickMark.LineColor = cl_g;
            chart5.ChartAreas[0].AxisY.MinorTickMark.LineColor = cl_g;
            //chart5.ChartAreas[0].AxisY.LabelStyle.Format = "0.0";
            chart5.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Arial", 9f, FontStyle.Bold);

            chart5.ChartAreas[0].AxisX.Title = "x, км";
            chart5.ChartAreas[0].AxisX.TitleFont = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart5.ChartAreas[0].AxisX.MajorGrid.LineColor = cl_g;
            chart5.ChartAreas[0].AxisX.MinorGrid.LineColor = cl_g;
            chart5.ChartAreas[0].AxisX.LineColor = cl_g;
            chart5.ChartAreas[0].AxisX.LabelStyle.ForeColor = Color.White;
            chart5.ChartAreas[0].AxisX.MajorTickMark.LineColor = cl_g;
            chart5.ChartAreas[0].AxisX.MinorTickMark.LineColor = cl_g;
            chart5.ChartAreas[0].AxisX.LabelStyle.Format = "0.00";
            chart5.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Arial", 9f, FontStyle.Bold);

            chart5.Legends[0].Font = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart5.ChartAreas[0].AxisY2.Title = "T, K";
            chart5.ChartAreas[0].AxisY2.TitleFont = new System.Drawing.Font("Arial", 10f, FontStyle.Bold);
            chart5.ChartAreas[0].AxisY2.MajorGrid.LineColor = cl_g;
            chart5.ChartAreas[0].AxisY2.MinorGrid.LineColor = cl_g;
            chart5.ChartAreas[0].AxisY2.LineColor = cl_g;
            chart5.ChartAreas[0].AxisY2.LabelStyle.ForeColor = Color.White;
            chart5.ChartAreas[0].AxisY2.MajorTickMark.LineColor = cl_g;
            chart5.ChartAreas[0].AxisY2.MinorTickMark.LineColor = cl_g;
            chart5.ChartAreas[0].AxisY2.LabelStyle.Format = "0.00";
            chart5.ChartAreas[0].AxisY2.LabelStyle.Font = new System.Drawing.Font("Arial", 9f, FontStyle.Bold);

            chart.AxisX.Minimum = 0;
            chart.AxisX.Maximum = Convert.ToDouble(dataGridView5.Rows[dataGridView5.Rows.Count - 2].Cells[0].Value);
            chart5.ChartAreas[0].AxisX.Interval = chart.AxisX.Maximum / 15;

            chart.AxisY.Minimum = numin - deltanu;
            chart.AxisY.Maximum = numax + deltanu;
            chart5.ChartAreas[0].AxisY.Interval = (numax - numin + 2 * deltanu) / 15;

            chart.AxisY2.Minimum = Tmin - deltaT;
            chart.AxisY2.Maximum = Tmax + deltaT;
            chart5.ChartAreas[0].AxisY2.Interval = (Tmax - Tmin + 2 * deltaT) / 15;

            for (int j = 0; j < dataGridView7.Rows.Count - 1; j++)
            {
                x_rl = Convert.ToDouble(dataGridView7.Rows[j].Cells[0].Value);
                if (x_s[l] <= x_rl && x_f[l] >= x_rl)
                {
                    chart5.Series[fuels[l] + "_ν"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[6].Value));
                    chart5.Series[fuels[l] + "_T"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[4].Value));
                }
                else
                {
                    if (j != n - 1)
                    {
                        l++;
                        chart5.Series[fuels[l] + "_ν"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[6].Value));
                        chart5.Series[fuels[l] + "_T"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[4].Value));
                    }
                }
            }
            button28.Enabled = true;
            textBox49.Text = (chart.AxisX.Minimum).ToString();
            textBox48.Text = (chart.AxisX.Maximum).ToString();

            textBox47.Text = (chart.AxisY.Minimum).ToString();
            textBox46.Text = (chart.AxisY.Maximum).ToString();

            textBox45.Text = (chart.AxisY2.Minimum).ToString();
            textBox44.Text = (chart.AxisY2.Maximum).ToString();

            textBox43.Text = chart5.ChartAreas[0].AxisX.Interval.ToString();
            textBox42.Text = chart5.ChartAreas[0].AxisY.Interval.ToString();
            textBox41.Text = chart5.ChartAreas[0].AxisY2.Interval.ToString("0.00");

            textBox49.Enabled = true;
            textBox48.Enabled = true;
            textBox47.Enabled = true;
            textBox46.Enabled = true;
            textBox45.Enabled = true;
            textBox44.Enabled = true;
            textBox43.Enabled = true;
            textBox42.Enabled = true;
            textBox41.Enabled = true;
        }
        //Масштабирование H(x)-P(x)
        private void button26_Click(object sender, EventArgs e)
        {
            var chart = chart1.ChartAreas[0];

            if (Double.TryParse(textBox3.Text.ToString(), out double check) == true)
                chart.AxisX.Minimum = Convert.ToDouble(textBox3.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел x_min ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox4.Text.ToString(), out check) == true)
                chart.AxisX.Maximum = Convert.ToDouble(textBox4.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел x_max ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox26.Text.ToString(), out check) == true)
                chart.AxisY.Minimum = Convert.ToDouble(textBox26.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел H_min ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox25.Text.ToString(), out check) == true)
                chart.AxisY.Maximum = Convert.ToDouble(textBox25.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел H_max ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox28.Text.ToString(), out check) == true)
                chart.AxisY2.Minimum = Convert.ToDouble(textBox28.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел P_min ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox27.Text.ToString(), out check) == true)
                chart.AxisY2.Maximum = Convert.ToDouble(textBox27.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел P_max ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox35.Text.ToString(), out check) == true)
                chart1.ChartAreas[0].AxisX.Interval = Convert.ToDouble(textBox35.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел Δx ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox36.Text.ToString(), out check) == true)
                chart1.ChartAreas[0].AxisY.Interval = Convert.ToDouble(textBox36.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел ΔH ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox37.Text.ToString(), out check) == true)
                chart1.ChartAreas[0].AxisY2.Interval = Convert.ToDouble(textBox37.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел ΔP ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        //Масштабирование ρ(T(x))
        private void button27_Click(object sender, EventArgs e)
        {
            var chart = chart4.ChartAreas[0];

            if (Double.TryParse(textBox34.Text.ToString(), out double check) == true)
                chart.AxisX.Minimum = Convert.ToDouble(textBox34.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел x_min ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox33.Text.ToString(), out check) == true)
                chart.AxisX.Maximum = Convert.ToDouble(textBox33.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел x_max ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox32.Text.ToString(), out check) == true)
                chart.AxisY.Minimum = Convert.ToDouble(textBox32.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел ρ_min ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox31.Text.ToString(), out check) == true)
                chart.AxisY.Maximum = Convert.ToDouble(textBox31.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел ρ_max ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox30.Text.ToString(), out check) == true)
                chart.AxisY2.Minimum = Convert.ToDouble(textBox30.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел T_min ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox29.Text.ToString(), out check) == true)
                chart.AxisY2.Maximum = Convert.ToDouble(textBox29.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел T_max ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox40.Text.ToString(), out check) == true)
                chart4.ChartAreas[0].AxisX.Interval = Convert.ToDouble(textBox40.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел Δx ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox39.Text.ToString(), out check) == true)
                chart4.ChartAreas[0].AxisY.Interval = Convert.ToDouble(textBox39.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел Δρ ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox38.Text.ToString(), out check) == true)
                chart4.ChartAreas[0].AxisY2.Interval = Convert.ToDouble(textBox38.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел ΔT ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        //Масштабирование ν(T(x))
        private void button28_Click(object sender, EventArgs e)
        {
            var chart = chart5.ChartAreas[0];

            if (Double.TryParse(textBox49.Text.ToString(), out double check) == true)
                chart.AxisX.Minimum = Convert.ToDouble(textBox49.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел x_min ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox48.Text.ToString(), out check) == true)
                chart.AxisX.Maximum = Convert.ToDouble(textBox48.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел x_max ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox47.Text.ToString(), out check) == true)
                chart.AxisY.Minimum = Convert.ToDouble(textBox47.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел ν_min ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox46.Text.ToString(), out check) == true)
                chart.AxisY.Maximum = Convert.ToDouble(textBox46.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел ν_max ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox45.Text.ToString(), out check) == true)
                chart.AxisY2.Minimum = Convert.ToDouble(textBox45.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел T_min ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox44.Text.ToString(), out check) == true)
                chart.AxisY2.Maximum = Convert.ToDouble(textBox44.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел T_max ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox43.Text.ToString(), out check) == true)
                chart5.ChartAreas[0].AxisX.Interval = Convert.ToDouble(textBox43.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел Δx ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox42.Text.ToString(), out check) == true)
                chart5.ChartAreas[0].AxisY.Interval = Convert.ToDouble(textBox42.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел Δν ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (Double.TryParse(textBox41.Text.ToString(), out check) == true)
                chart5.ChartAreas[0].AxisY2.Interval = Convert.ToDouble(textBox41.Text.ToString());
            else
                MessageBox.Show("Не правильно введен предел ΔT ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        //Расчет начальных параметров смеси
        private void button29_Click(object sender, EventArgs e)
        {
            dataGridView10.Rows.Clear();
            double prom = 0;
            double h1 = 0, h2 = 0, h3 = 0;
            double p1 = 0, p2 = 0, p3 = 0;
            double x1 = 0, x2 = 0;
            double z1 = 0, z2 = 0, z3 = 0;
            double T1 = 0, T2 = 0, T3 = 0;
            double ro1 = 0, ro2 = 0, ro3 = 0;
            double nu1 = 0, nu2 = 0, nu3 = 0;
            double lyam1 = 0, lyam2 = 0, lyam3 = 0;
            for (int i = dataGridView9.Rows.Count - 2; i >= 1; i--)
            {
                //Выписаваем НП участвующие в контакте
                dataGridView10.Rows.Add(fuels[i] + "+" + fuels[i - 1], 0, 0, 0, 0, 0, 0, 0);
                //if (i == 1)
                //{
                //    dataGridView10.Rows.Add(fuels[i - 1] + "+" + fuels[dataGridView9.Rows.Count - 2], 0, 0, 0, 0, 0, 0, 0);
                //}
                vol_s[i] = 0;
                double vol_dop = 0;

                if (x_s[i] > 0 && x_s[i] < Convert.ToDouble(dataGridView5.Rows[dataGridView5.Rows.Count - 2].Cells[0].Value))
                {
                    int m = 0, m1= 0;
                    //Считаем объем смеси
                    while (x[m] < x_s[i])
                    {
                        vol_s[i] = vol_s[i] + Math.Pow((1000 * (Math.Pow(lyam_sp[i, m], 1.8) + Math.Pow(lyam_sp[i - 1, m], 1.8)) * (pi / 4) *
                            Math.Pow(((D[m] - 2 * sig[m]) / 1000), 1.57) * Math.Pow(deltax * 1000, 0.57)), (1 / 0.57));

                        //vol_dop = vol_dop + Math.Pow(4 * 1.645 *  Math.Pow((D[m] - 2 * sig[m]) / 1000, 2) * pi / 4 * Math.Sqrt(deltax * 1000 / v[m]) * Math.Sqrt(lyam_rl[i, m] / (ro_rl[i, m] * cp_rl[i, m])), 1 / 0.57);
                        
                        m++;
                    }
                    m--;
                    vol_s[i] = vol_s[i] + Math.Pow((1000 * (Math.Pow(lyam_sp[i, m], 1.8) + Math.Pow(lyam_sp[i - 1, m], 1.8)) * (pi / 4) *
                        Math.Pow(((D[m + 1] - 2 * sig[m + 1]) / 1000), 1.57) * Math.Pow((x_s[i] - x[m]) * 1000, 0.57)), (1 / 0.57));

                    //vol_dop = vol_dop + Math.Pow(4 * 1.645 * Math.Pow((D[m] - 2 * sig[m]) / 1000, 2) * Math.Sqrt(deltax * 1000 / v[m]) * Math.Sqrt(lyam_rl[i, m] / (ro_rl[i, m] * cp_rl[i, m])), 1 / 0.57);
                    //MessageBox.Show(vol_dop.ToString(), "Внимание", MessageBoxButtons.OK);
                    vol_s[i] = Math.Pow(vol_s[i], 0.57);

                    //Переписать учет первоначальной смеси
                    /*
                    if (x_s[i] <= 50)
                        vol_s[i] = vol_s[i] * 1.2;
                    else
                        if (x_s[i] > 50 && x_s[i] <= 100)
                        vol_s[i] = vol_s[i] * 1.15;
                    else
                        if (x_s[i] > 100 && x_s[i] <= 200)
                        vol_s[i] = vol_s[i] * 1.1;
                    else
                        if (x_s[i] > 200 && x_s[i] <= 300)
                        vol_s[i] = vol_s[i] * 1.05;
                    else
                        if (x_s[i] > 300 && x_s[i] <= 500)
                        vol_s[i] = vol_s[i] * 1.01;
                    else
                        if (x_s[i] > 500)
                        vol_s[i] = vol_s[i] * 1.006;
                    */
                    

                    dataGridView10.Rows[dataGridView9.Rows.Count - 2 - i].Cells[1].Value = vol_s[i];
                    vol_s[i] = vol_s[i] + 100 * Math.Pow((D[0] / 500), 2);
                    dataGridView10.Rows[dataGridView9.Rows.Count - 2 - i].Cells[18].Value = vol_s[i];
                    //Считаем коориднату начала смеси
                    m1 = m;
                    s_smes[i] = x_s[i];
                    prom = vol_s[i] / 2;
                    prom = prom - pi / 4 * (x_s[i] - x[m1]) * 1000 * Math.Pow(((D[m1 + 1] - 2 * sig[m1 + 1]) / 1000), 2);
                    s_smes[i] = s_smes[i] - (x_s[i] - x[m1]);
                    MessageBox.Show(m1.ToString(), "Внимание", MessageBoxButtons.OK);
                    while (m1 > 0 && prom - pi / 4 * deltax * 1000 * Math.Pow(((D[m1] - 2 * sig[m1]) / 1000), 2) > 0)
                    {
                        prom = prom - pi / 4 * deltax * 1000 * Math.Pow(((D[m1] - 2 * sig[m1]) / 1000), 2);
                        s_smes[i] = s_smes[i] - deltax;
                        m1--;
                    }
                    s_smes[i] = s_smes[i] - prom * 4 / pi * 1 / 1000 * 1 / Math.Pow(((D[m1] - 2 * sig[m1]) / 1000), 2);
                    dataGridView10.Rows[dataGridView9.Rows.Count - 2 - i].Cells[2].Value = s_smes[i];
                   
                    //Считаем коориднату хвоста смеси
                    m1 = m;
                    m1++;
                    f_smes[i] = x_s[i];
                    prom = vol_s[i] / 2;
                    prom = prom - pi / 4 * (x[m1] - x_s[i]) * 1000 * Math.Pow(((D[m1 + 1] - 2 * sig[m1 + 1]) / 1000), 2);
                    f_smes[i] = f_smes[i] + (x[m1] - x_s[i]);

                    while (prom - pi / 4 * deltax * 1000 * Math.Pow(((D[m1] - 2 * sig[m1]) / 1000), 2) > 0)
                    {
                        prom = prom - pi / 4 * deltax * 1000 * Math.Pow(((D[m1] - 2 * sig[m1]) / 1000), 2);
                        f_smes[i] = f_smes[i] + deltax;
                        m1++;
                    }
                    f_smes[i] = f_smes[i] + prom * 4 / pi * 1 / 1000 * 1 / Math.Pow(((D[m1] - 2 * sig[m1]) / 1000), 2);
                    dataGridView10.Rows[dataGridView9.Rows.Count - 2 - i].Cells[3].Value = f_smes[i];
                    //Считаем напоры и давления в начале и хвосте смеси
                    for (int j = 0; j <= dataGridView7.Rows.Count - 2; j++)
                        if (
                            Convert.ToDouble(dataGridView7.Rows[j].Cells[0].Value) <= s_smes[i] &&
                            Convert.ToDouble(dataGridView7.Rows[j + 1].Cells[0].Value) > s_smes[i]
                           )
                        {
                            x1 = Convert.ToDouble(dataGridView7.Rows[j].Cells[0].Value);
                            x2 = Convert.ToDouble(dataGridView7.Rows[j + 1].Cells[0].Value);
                            z1 = Convert.ToDouble(dataGridView7.Rows[j].Cells[1].Value);
                            z2 = Convert.ToDouble(dataGridView7.Rows[j + 1].Cells[1].Value);
                            T1 = Convert.ToDouble(dataGridView7.Rows[j].Cells[4].Value);
                            T2 = Convert.ToDouble(dataGridView7.Rows[j + 1].Cells[4].Value);
                            ro1 = Convert.ToDouble(dataGridView7.Rows[j].Cells[5].Value);
                            ro2 = Convert.ToDouble(dataGridView7.Rows[j + 1].Cells[5].Value);
                            nu1 = Convert.ToDouble(dataGridView7.Rows[j].Cells[6].Value);
                            nu2 = Convert.ToDouble(dataGridView7.Rows[j + 1].Cells[6].Value);
                            h1 = Convert.ToDouble(dataGridView7.Rows[j].Cells[12].Value);
                            h2 = Convert.ToDouble(dataGridView7.Rows[j + 1].Cells[12].Value);
                            p1 = Convert.ToDouble(dataGridView7.Rows[j].Cells[13].Value);
                            p2 = Convert.ToDouble(dataGridView7.Rows[j + 1].Cells[13].Value);
                            lyam1 = Convert.ToDouble(dataGridView7.Rows[j].Cells[10].Value);
                            lyam2 = Convert.ToDouble(dataGridView7.Rows[j + 1].Cells[10].Value);

                            z3 = z1 + (z2 - z1) / (x2 - x1) * (s_smes[i] - x1);
                            T3 = T1 + (T2 - T1) / (x2 - x1) * (s_smes[i] - x1);
                            ro3 = ro1 + (ro2 - ro1) / (x2 - x1) * (s_smes[i] - x1);
                            nu3 = nu1 + (nu2 - nu1) / (x2 - x1) * (s_smes[i] - x1);
                            h3 = h1 + (h2 - h1) / (x2 - x1) * (s_smes[i] - x1);
                            p3 = p1 + (p2 - p1) / (x2 - x1) * (s_smes[i] - x1);
                            lyam3 = lyam1 + (lyam2 - lyam1) / (x2 - x1) * (s_smes[i] - x1);

                            dataGridView10.Rows[dataGridView9.Rows.Count - 2 - i].Cells[4].Value = z3;
                            dataGridView10.Rows[dataGridView9.Rows.Count - 2 - i].Cells[6].Value = T3;
                            dataGridView10.Rows[dataGridView9.Rows.Count - 2 - i].Cells[8].Value = ro3;
                            dataGridView10.Rows[dataGridView9.Rows.Count - 2 - i].Cells[10].Value = nu3;
                            dataGridView10.Rows[dataGridView9.Rows.Count - 2 - i].Cells[12].Value = h3;
                            dataGridView10.Rows[dataGridView9.Rows.Count - 2 - i].Cells[14].Value = p3;
                            dataGridView10.Rows[dataGridView9.Rows.Count - 2 - i].Cells[16].Value = lyam3;
                        }
                        else
                        if (
                            Convert.ToDouble(dataGridView7.Rows[j].Cells[0].Value) <= f_smes[i] &&
                            Convert.ToDouble(dataGridView7.Rows[j + 1].Cells[0].Value) > f_smes[i]
                            )
                        {
                            x1 = Convert.ToDouble(dataGridView7.Rows[j].Cells[0].Value);
                            x2 = Convert.ToDouble(dataGridView7.Rows[j + 1].Cells[0].Value);
                            z1 = Convert.ToDouble(dataGridView7.Rows[j].Cells[1].Value);
                            z2 = Convert.ToDouble(dataGridView7.Rows[j + 1].Cells[1].Value);
                            T1 = Convert.ToDouble(dataGridView7.Rows[j].Cells[4].Value);
                            T2 = Convert.ToDouble(dataGridView7.Rows[j + 1].Cells[4].Value);
                            ro1 = Convert.ToDouble(dataGridView7.Rows[j].Cells[5].Value);
                            ro2 = Convert.ToDouble(dataGridView7.Rows[j + 1].Cells[5].Value);
                            nu1 = Convert.ToDouble(dataGridView7.Rows[j].Cells[6].Value);
                            nu2 = Convert.ToDouble(dataGridView7.Rows[j + 1].Cells[6].Value);
                            h1 = Convert.ToDouble(dataGridView7.Rows[j].Cells[12].Value);
                            h2 = Convert.ToDouble(dataGridView7.Rows[j + 1].Cells[12].Value);
                            p1 = Convert.ToDouble(dataGridView7.Rows[j].Cells[13].Value);
                            p2 = Convert.ToDouble(dataGridView7.Rows[j + 1].Cells[13].Value);
                            lyam1 = Convert.ToDouble(dataGridView7.Rows[j].Cells[10].Value);
                            lyam2 = Convert.ToDouble(dataGridView7.Rows[j + 1].Cells[10].Value);


                            z3 = z1 + (z2 - z1) / (x2 - x1) * (f_smes[i] - x1);
                            T3 = T1 + (T2 - T1) / (x2 - x1) * (f_smes[i] - x1);
                            ro3 = ro1 + (ro2 - ro1) / (x2 - x1) * (f_smes[i] - x1);
                            nu3 = nu1 + (nu2 - nu1) / (x2 - x1) * (f_smes[i] - x1);
                            h3 = h1 + (h2 - h1) / (x2 - x1) * (f_smes[i] - x1);
                            p3 = p1 + (p2 - p1) / (x2 - x1) * (f_smes[i] - x1);
                            lyam3 = lyam1 + (lyam2 - lyam1) / (x2 - x1) * (f_smes[i] - x1);

                            dataGridView10.Rows[dataGridView9.Rows.Count - 2 - i].Cells[5].Value = z3;
                            dataGridView10.Rows[dataGridView9.Rows.Count - 2 - i].Cells[7].Value = T3;
                            dataGridView10.Rows[dataGridView9.Rows.Count - 2 - i].Cells[9].Value = ro3;
                            dataGridView10.Rows[dataGridView9.Rows.Count - 2 - i].Cells[11].Value = nu3;
                            dataGridView10.Rows[dataGridView9.Rows.Count - 2 - i].Cells[13].Value = h3;
                            dataGridView10.Rows[dataGridView9.Rows.Count - 2 - i].Cells[15].Value = p3;
                            dataGridView10.Rows[dataGridView9.Rows.Count - 2 - i].Cells[17].Value = lyam3;
                        }
                }
            }
        }
        //Полный расчет смеси
        private void button30_Click(object sender, EventArgs e)
        {
            textBox50.Text = fuels[i_i] + "+" + fuels[i_i - 1];
            button33.Enabled = true;
            button32.Enabled = true;
            double D_sq = 0, sig_sq = 0, eps_sq = 0, V_sq = 0, t_sq = 0, K_sq = 0, c_sq = 0, 
                ro_sq = 0, nu_sq = 0, Re_sq = 0, i_sq = 0, H_sq = 0, P_sq = 0;
            dataGridView11.Rows.Clear();


            double delta_x = (f_smes[i_i] - s_smes[i_i]) / n_sq;
            double delta_z = (Convert.ToDouble(dataGridView10.Rows[dataGridView10.Rows.Count - 1 - i_i].Cells[5].Value) - Convert.ToDouble(dataGridView10.Rows[dataGridView10.Rows.Count - 1 - i_i].Cells[4].Value)) / n_sq;
            double delta_T = (Convert.ToDouble(dataGridView10.Rows[dataGridView10.Rows.Count - 1 - i_i].Cells[7].Value) - Convert.ToDouble(dataGridView10.Rows[dataGridView10.Rows.Count - 1 - i_i].Cells[6].Value)) / n_sq;

            double x_sq = s_smes[i_i];
            double z_sq = Convert.ToDouble(dataGridView10.Rows[dataGridView10.Rows.Count - 1 - i_i].Cells[4].Value);
            double T_sq = Convert.ToDouble(dataGridView10.Rows[dataGridView10.Rows.Count - 1 - i_i].Cells[6].Value);
            double lyam_sq = Convert.ToDouble(dataGridView10.Rows[dataGridView10.Rows.Count - 1 - i_i].Cells[16].Value);

            H_sq = Convert.ToDouble(dataGridView10.Rows[dataGridView10.Rows.Count - 1 - i_i].Cells[12].Value);
            P_sq = Convert.ToDouble(dataGridView10.Rows[dataGridView10.Rows.Count - 1 - i_i].Cells[14].Value);

            while (x_sq < f_smes[i_i])
            {
                t_sq = 0;
                for (int j = 0; j <= n - 2; j++)
                    if (
                        x[j] <= x_sq &&
                        x[j + 1] > x_sq
                       )
                    {
                        D_sq = D[j];
                        sig_sq = sig[j];
                        eps_sq = eps_sp[j];
                        V_sq = v[j];
                        t_sq = t_sq + (x_sq - x[j]) / v[j] * 1000;
                    }
                    else if (x[j] <= x_sq)
                        t_sq = t_sq + deltax * 1000 / v[j];
                K_sq = 1.785 * Math.Sqrt(lyam_sq) * (D_sq - 2 * sig_sq) / 1000 * V_sq;
                c_sq = 0.5 * MathNet.Numerics.SpecialFunctions.Erfc( 
                               (x_sq - x_f[i_i - 1]) * 1000 / 
                               (2 * Math.Sqrt(K_sq * t_sq))
                              );
                ro_sq = c_sq * Convert.ToDouble(dataGridView10.Rows[dataGridView10.Rows.Count - 1 - i_i].Cells[8].Value) +
                    (1 - c_sq) * Convert.ToDouble(dataGridView10.Rows[dataGridView10.Rows.Count - 1 - i_i].Cells[9].Value);
                nu_sq = c_sq * Convert.ToDouble(dataGridView10.Rows[dataGridView10.Rows.Count - 1 - i_i].Cells[10].Value) +
                    (1 - c_sq) * Convert.ToDouble(dataGridView10.Rows[dataGridView10.Rows.Count - 1 - i_i].Cells[11].Value);
                Re_sq = V_sq * (D_sq - 2 * sig_sq) / (nu_sq * 1000);

                if (Re_sq < 10 / eps_sq)
                    lyam_sq = 0.3164 / Math.Pow(Re_sq, 0.25);
                else
                        if (Re_sq >= 10 / eps_sq && Re_sq < 500 / eps_sq)
                    lyam_sq = 0.11 * Math.Pow((eps_sq + 68 / Re_sq), 0.25);
                else
                        if (Re_sq >= 500 / eps_sq)
                    lyam_sq = 0.11 * Math.Pow(eps_sq, 0.25);

                i_sq = lyam_sq * 8 * Math.Pow(1000 / (D_sq - 2 * sig_sq), 5) * Math.Pow(Q, 2) / (Math.Pow(pi, 2) * g);

                dataGridView11.Rows.Add(x_sq, z_sq, D_sq, sig_sq, T_sq, eps_sq, V_sq , t_sq, 
                    K_sq, c_sq, 1 - c_sq, ro_sq, nu_sq, Re_sq, lyam_sq, i_sq, H_sq , P_sq);
                x_sq = x_sq + delta_x;
                z_sq = z_sq + delta_z;
                T_sq = T_sq + delta_T;
                P_sq = P_sq - lyam_sq * delta_x * Math.Pow(V_sq, 2) * ro_sq / 2 * 1 / (D_sq - 2 * sig_sq);
                H_sq = P_sq * Math.Pow(10,6) / (ro_sq * g) + Math.Pow(V_sq, 2) / (2 * g) + z_sq;
            }
            
        }
        //Переключение расчитываемой смеси
        private void button33_Click(object sender, EventArgs e)
        {
            if (i_i > 1)
            {
                i_i--;
                textBox50.Text = fuels[i_i] + "+" + fuels[i_i - 1];
            }
                
        }
        private void button32_Click(object sender, EventArgs e)
        {
            if (i_i < dataGridView10.Rows.Count - 1)
            {
                i_i++;
                textBox50.Text = fuels[i_i] + "+" + fuels[i_i - 1];
            }
                
        }
        //График P-H смеси
        private void button31_Click(object sender, EventArgs e)
        {
            double x_rl = 1;
            chart6.Series.Clear();

            chart6.Series.Add("sqx_H");
            chart6.Series["sqx_H"].YAxisType = AxisType.Primary;
            chart6.Series["sqx_H"].ChartType = SeriesChartType.Spline;
            chart6.Series["sqx_H"].Color = T_c;
            chart6.Series["sqx_H"].BorderWidth = 2;

            chart6.Series.Add("sqx_P");
            chart6.Series["sqx_P"].YAxisType = AxisType.Secondary;
            chart6.Series["sqx_P"].ChartType = SeriesChartType.Spline;
            chart6.Series["sqx_P"].Color = T_c;
            chart6.Series["sqx_P"].BorderWidth = 2;

            foreach (var series in chart6.Series)
            {
                series.Points.Clear();
            }

            var chart = chart6.ChartAreas[0];

            chart6.Legends[0].Enabled = true;
            chart6.Legends[0].Font = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart6.ChartAreas[0].AxisY2.Title = "P, Мпа";
            chart6.ChartAreas[0].AxisY2.TitleFont = new System.Drawing.Font("Arial", 10f, FontStyle.Bold);
            chart6.ChartAreas[0].AxisY2.MajorGrid.LineColor = cl_g;
            chart6.ChartAreas[0].AxisY2.MinorGrid.LineColor = cl_g;
            chart6.ChartAreas[0].AxisY2.LineColor = cl_g;
            chart6.ChartAreas[0].AxisY2.LabelStyle.ForeColor = Color.White;
            chart6.ChartAreas[0].AxisY2.MajorTickMark.LineColor = cl_g;
            chart6.ChartAreas[0].AxisY2.MinorTickMark.LineColor = cl_g;
            chart6.ChartAreas[0].AxisY2.LabelStyle.Format = "0.00";
            chart6.ChartAreas[0].AxisY2.LabelStyle.Font = new System.Drawing.Font("Arial", 9f, FontStyle.Bold);

            chart6.ChartAreas[0].AxisY.Title = "H, м";
            chart6.ChartAreas[0].AxisY.TitleFont = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart6.ChartAreas[0].AxisY.MajorGrid.LineColor = cl_g;
            chart6.ChartAreas[0].AxisY.MinorGrid.LineColor = cl_g;
            chart6.ChartAreas[0].AxisY.LineColor = cl_g;
            chart6.ChartAreas[0].AxisY.LabelStyle.ForeColor = Color.White;
            chart6.ChartAreas[0].AxisY.MajorTickMark.LineColor = cl_g;
            chart6.ChartAreas[0].AxisY.MinorTickMark.LineColor = cl_g;
            chart6.ChartAreas[0].AxisY.LabelStyle.Format = "0.0";
            chart6.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Arial", 9f, FontStyle.Bold);

            chart6.ChartAreas[0].AxisX.Title = "x, км";
            chart6.ChartAreas[0].AxisX.TitleFont = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart6.ChartAreas[0].AxisX.MajorGrid.LineColor = cl_g;
            chart6.ChartAreas[0].AxisX.MinorGrid.LineColor = cl_g;
            chart6.ChartAreas[0].AxisX.LineColor = cl_g;
            chart6.ChartAreas[0].AxisX.LabelStyle.ForeColor = Color.White;
            chart6.ChartAreas[0].AxisX.MajorTickMark.LineColor = cl_g;
            chart6.ChartAreas[0].AxisX.MinorTickMark.LineColor = cl_g;
            chart6.ChartAreas[0].AxisX.LabelStyle.Format = "0.00";
            chart6.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Arial", 9f, FontStyle.Bold);

            chart.AxisX.Minimum = Convert.ToDouble(dataGridView11.Rows[0].Cells[0].Value);
            chart.AxisX.Maximum = Convert.ToDouble(dataGridView11.Rows[dataGridView11.Rows.Count - 2].Cells[0].Value);
            chart6.ChartAreas[0].AxisX.Interval = (chart.AxisX.Maximum - chart.AxisX.Minimum) / 15;

            chart.AxisY.Minimum = Convert.ToDouble(dataGridView11.Rows[dataGridView11.Rows.Count - 2].Cells[16].Value); 
            chart.AxisY.Maximum = Convert.ToDouble(dataGridView11.Rows[0].Cells[16].Value);
            chart6.ChartAreas[0].AxisY.Interval = (chart.AxisY.Maximum - chart.AxisY.Minimum) / 15;


            chart.AxisY2.Minimum = Convert.ToDouble(dataGridView11.Rows[dataGridView11.Rows.Count - 2].Cells[17].Value); 
            chart.AxisY2.Maximum = Convert.ToDouble(dataGridView11.Rows[0].Cells[17].Value);
            chart6.ChartAreas[0].AxisY2.Interval = (chart.AxisY2.Maximum - chart.AxisY2.Minimum) / 15;

            chart6.Series["sqx_H"].Color = Color.FromArgb(64, rnd.Next(128, 255), 64);
            chart6.Series["sqx_P"].Color = Color.FromArgb(rnd.Next(128, 255), 64, 64);

            for (int j = 0; j < dataGridView11.Rows.Count - 1; j++)
            {
                x_rl = Convert.ToDouble(dataGridView11.Rows[j].Cells[0].Value);
                chart6.Series["sqx_H"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView11.Rows[j].Cells[16].Value));
                chart6.Series["sqx_P"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView11.Rows[j].Cells[17].Value));
            }
        }
        //График ρ(T(x)) смеси
        private void button34_Click(object sender, EventArgs e)
        {
            double x_rl = 1;
            chart7.Series.Clear();

            chart7.Series.Add("sqx_ρ");
            chart7.Series["sqx_ρ"].YAxisType = AxisType.Primary;
            chart7.Series["sqx_ρ"].ChartType = SeriesChartType.Spline;
            chart7.Series["sqx_ρ"].Color = T_c;
            chart7.Series["sqx_ρ"].BorderWidth = 2;

            chart7.Series.Add("sqx_T");
            chart7.Series["sqx_T"].YAxisType = AxisType.Secondary;
            chart7.Series["sqx_T"].ChartType = SeriesChartType.Spline;
            chart7.Series["sqx_T"].Color = T_c;
            chart7.Series["sqx_T"].BorderWidth = 2;

            foreach (var series in chart7.Series)
            {
                series.Points.Clear();
            }

            var chart = chart7.ChartAreas[0];

            chart7.Legends[0].Enabled = true;
            chart7.Legends[0].Font = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart7.ChartAreas[0].AxisY2.Title = "ρ, кг/м^3";
            chart7.ChartAreas[0].AxisY2.TitleFont = new System.Drawing.Font("Arial", 10f, FontStyle.Bold);
            chart7.ChartAreas[0].AxisY2.MajorGrid.LineColor = cl_g;
            chart7.ChartAreas[0].AxisY2.MinorGrid.LineColor = cl_g;
            chart7.ChartAreas[0].AxisY2.LineColor = cl_g;
            chart7.ChartAreas[0].AxisY2.LabelStyle.ForeColor = Color.White;
            chart7.ChartAreas[0].AxisY2.MajorTickMark.LineColor = cl_g;
            chart7.ChartAreas[0].AxisY2.MinorTickMark.LineColor = cl_g;
            chart7.ChartAreas[0].AxisY2.LabelStyle.Format = "0.00";
            chart7.ChartAreas[0].AxisY2.LabelStyle.Font = new System.Drawing.Font("Arial", 9f, FontStyle.Bold);

            chart7.ChartAreas[0].AxisY.Title = "T, K";
            chart7.ChartAreas[0].AxisY.TitleFont = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart7.ChartAreas[0].AxisY.MajorGrid.LineColor = cl_g;
            chart7.ChartAreas[0].AxisY.MinorGrid.LineColor = cl_g;
            chart7.ChartAreas[0].AxisY.LineColor = cl_g;
            chart7.ChartAreas[0].AxisY.LabelStyle.ForeColor = Color.White;
            chart7.ChartAreas[0].AxisY.MajorTickMark.LineColor = cl_g;
            chart7.ChartAreas[0].AxisY.MinorTickMark.LineColor = cl_g;
            chart7.ChartAreas[0].AxisY.LabelStyle.Format = "0.0";
            chart7.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Arial", 9f, FontStyle.Bold);
            
            chart7.ChartAreas[0].AxisX.Title = "x, км";
            chart7.ChartAreas[0].AxisX.TitleFont = new System.Drawing.Font("Arial Narrow", 10f, FontStyle.Bold);
            chart7.ChartAreas[0].AxisX.MajorGrid.LineColor = cl_g;
            chart7.ChartAreas[0].AxisX.MinorGrid.LineColor = cl_g;
            chart7.ChartAreas[0].AxisX.LineColor = cl_g;
            chart7.ChartAreas[0].AxisX.LabelStyle.ForeColor = Color.White;
            chart7.ChartAreas[0].AxisX.MajorTickMark.LineColor = cl_g;
            chart7.ChartAreas[0].AxisX.MinorTickMark.LineColor = cl_g;
            chart7.ChartAreas[0].AxisX.LabelStyle.Format = "0.00";
            chart7.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Arial", 9f, FontStyle.Bold);

            chart.AxisX.Minimum = Convert.ToDouble(dataGridView11.Rows[0].Cells[0].Value);
            chart.AxisX.Maximum = Convert.ToDouble(dataGridView11.Rows[dataGridView11.Rows.Count - 2].Cells[0].Value);
            //chart7.ChartAreas[0].AxisX.Interval = (chart.AxisX.Maximum - chart.AxisX.Minimum) / 15;

            if (Convert.ToDouble(dataGridView11.Rows[0].Cells[11].Value) > Convert.ToDouble(dataGridView11.Rows[dataGridView11.Rows.Count - 2].Cells[11].Value))
            {
                chart.AxisY.Minimum = Convert.ToDouble(dataGridView11.Rows[dataGridView11.Rows.Count - 2].Cells[11].Value);
                chart.AxisY.Maximum = Convert.ToDouble(dataGridView11.Rows[0].Cells[11].Value);
            }
            else
            {
                chart.AxisY.Minimum = Convert.ToDouble(dataGridView11.Rows[0].Cells[11].Value);
                chart.AxisY.Maximum = Convert.ToDouble(dataGridView11.Rows[dataGridView11.Rows.Count - 2].Cells[11].Value);
            }
            
            //chart7.ChartAreas[0].AxisY.Interval = (chart.AxisY.Maximum - chart.AxisY.Minimum) / 15;

            if (Convert.ToDouble(dataGridView11.Rows[0].Cells[4].Value) > Convert.ToDouble(dataGridView11.Rows[dataGridView11.Rows.Count - 2].Cells[4].Value))
            {
                chart.AxisY2.Minimum = Convert.ToDouble(dataGridView11.Rows[dataGridView11.Rows.Count - 2].Cells[4].Value);
                chart.AxisY2.Maximum = Convert.ToDouble(dataGridView11.Rows[0].Cells[4].Value);
            }
            else
            {
                chart.AxisY2.Minimum = Convert.ToDouble(dataGridView11.Rows[0].Cells[4].Value);
                chart.AxisY2.Maximum = Convert.ToDouble(dataGridView11.Rows[dataGridView11.Rows.Count - 2].Cells[4].Value);
            }
            //chart7.ChartAreas[0].AxisY2.Interval = (chart.AxisY2.Maximum - chart.AxisY2.Minimum) / 15;

            chart7.Series["sqx_ρ"].Color = Color.FromArgb(64, rnd.Next(128, 255), 64);
            chart7.Series["sqx_T"].Color = Color.FromArgb(rnd.Next(128, 255), 64, 64);

            for (int j = 0; j < dataGridView11.Rows.Count - 1; j++)
            {
                x_rl = Convert.ToDouble(dataGridView11.Rows[j].Cells[0].Value);
                chart7.Series["sqx_ρ"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView11.Rows[j].Cells[11].Value));
                chart7.Series["sqx_T"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView11.Rows[j].Cells[4].Value));
            }
        }

        //Переменные для цикла
        System.Threading.Timer timer;
        int num = 1;
        private void button35_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(key_p.ToString(), "Чтение файла", MessageBoxButtons.OK, MessageBoxIcon.Information);
            double k = 1000;
            int i = 0;
            //Устанавливаем метод обратного вызова таймера
            //TimerCallback tm = new TimerCallback(Count);
            //Коэффициент ускорения (так как в жинзи нефть и нефтепродукты текут очень медленно относительно всей длины участка трубопровода)

            if (pumping == false)
            {
                pumping = true;
                //timer = new System.Threading.Timer(tm, num, 0, 2000);
                pictureBox1.BackColor = Color.Green;
                //Возвращаем первоначальную массу
                //for (i = 0; i < dataGridView6.Rows.Count - 1; i++)
                //    mass[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[4].Value);
            }
            else
            {
                pumping = false;
                //timer.Change(Timeout.Infinite, Timeout.Infinite);
                //timer = null;
                pictureBox1.BackColor = Color.Red;
            }

            //Возвращаем первоначальную массу
            for (i = 0; i <= dataGridView6.Rows.Count - 2; i++)
                mass[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[4].Value);
            i = 0;

            //Определяем первую закачиваемую в трубопровод жидкость
            //for (i = 0; i < dataGridView6.Rows.Count - 1; i++)
            //    if (x_f[i] == m_x[dataGridView5.Rows.Count - 2] && mass[i] > 0)
            //{
            //    MessageBox.Show(mass[i].ToString(), "Чтение файла", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    break;
            //}
            if (mass[key_p] - Q * ro[key_p] / 1000 * k * num >= 0)
            {
                //MessageBox.Show(key_p.ToString(), "Чтение файла", MessageBoxButtons.OK, MessageBoxIcon.Information);
                mass[key_p] = mass[key_p] - Q * ro[key_p] / 1000 * k * num;
            }
            else
            {
                mass[key_p] = 0;
                x_s[key_p] = -1000;
                x_f[key_p] = -1000;
                num = 0;
                if (key_p > 0)
                    key_p--;
                else
                    key_p = dataGridView6.Rows.Count - 2;
            }
            num++;


                // for (i = 0; i < dataGridView6.Rows.Count - 1; i++)
                //    MessageBox.Show(mass[i].ToString(), "Чтение файла", MessageBoxButtons.OK, MessageBoxIcon.Information);


            //Перерасчет положения 

            x_f[key_p] = m_x[dataGridView5.Rows.Count - 2];
            x_s[key_p] = x_f[key_p];

            //i = dataGridView6.Rows.Count - 2;
            i = key_p;
            int j = n - 2;

            for (j = n - 2; j >= 0; j--)
                if (mass[i] - ro_rl[i, j] * (deltax) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2) >= 0)
                {
                    if (x_s[i] - deltax >= 0)
                    {
                        mass[i] = mass[i] - ro_rl[i, j] * (deltax) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2);
                        x_s[i] = x_s[i] - deltax;
                    }
                    else
                    {
                        x_s[i] = 0;
                        mass[i] = mass[i] - ro_rl[i, j] * (x_s[i]) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2);
                    }
                }
                else
                {
                    if (x_s[i] -  (mass[i]) / (ro_rl[i, j]  * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2)) >= 0)
                        x_s[i] = x_s[i] - (deltax) * (mass[i]) / (ro_rl[i, j] * (deltax) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2));
                    if (i > 0)
                    {
                        i--;
                        x_f[i] = x_s[i + 1];
                        x_s[i] = x_f[i] - (deltax) * (1 - mass[i + 1] / (ro_rl[i + 1, j] * (deltax) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2)));
                        mass[i + 1] = 0;
                        mass[i] = mass[i] - ro_rl[i, j] * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2) * ((deltax) * (1 - mass[i + 1] / (ro_rl[i + 1, j] * (deltax) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2))));
                    }
                    else
                    {
                        i = dataGridView6.Rows.Count - 2;
                        x_f[i] = x_s[0];
                        x_s[i] = x_f[i] - (deltax) * (1 - mass[0] / (ro_rl[0, j] * (deltax) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2)));
                        mass[0] = 0;
                        mass[i] = mass[i] - ro_rl[i, j] * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2) * ((deltax) * (1 - mass[0] / (ro_rl[0, j] * (deltax) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2))));
                    }
                }
            dataGridView9.Rows.Clear();
            for (i = dataGridView6.Rows.Count - 2; i >= 0; i--)
            {
                //mass[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[4].Value);
                dataGridView9.Rows.Add(fuels[i], mass[i], vol[i],
                x_s[i], x_f[i], 0, 0, 0, 0);
            }


            
            //Перасчет в новых координатах сечений
            int kol = dataGridView6.Rows.Count - 2;
            dataGridView7.Rows.Clear();
            double H = 0, P = 0;
            //Определяем последнюю закачиваемую в трубопровод жидкость
            i = 0;
            for (i = 0; i < dataGridView6.Rows.Count - 1; i++)
                if (x_s[i] == 0 && mass[i] > 0)
                    break;
            int l = i;
            for (j = 0; j < n; j++)
            {
                if (x_s[l] <= x[j] && x_f[l] >= x[j])
                {
                    if (j == 0)
                    {
                        H = Convert.ToDouble(textBox14.Text);
                    }
                    else
                    {
                        //MessageBox.Show( (dataGridView7.Rows[j - 1].Cells[10].Value).ToString(), "Чтение файла", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        H = H - i_sp[l, j - 1] * deltax * 1000;
                    }
                    P = (H - Math.Pow(v[j], 2) / (2 * g) - z[j]) * ro_rl[l, j] * g;

                    if (x[j] == x_s[l] && kol -l >= 0)
                    {
                        
                        dataGridView9.Rows[kol - l].Cells[5].Value = H.ToString();
                        dataGridView9.Rows[kol - l].Cells[7].Value = (P * Math.Pow(10, -6)).ToString();
                    }
                    dataGridView7.Rows.Add(x[j].ToString("0.0000"), z[j], D[j], sig[j], Tk[l, j],
                        ro_rl[l, j], nu_rl[l, j], v[j], eps_sp[j], Re[l, j], lyam_sp[l, j], i_sp[l, j], H, (P * Math.Pow(10, -6)));
                }
                else
                {
                    if (j < n - 1)
                        l++;
                    if (kol + 1 - l >= 0 && x_s[l] <= x[j] && x_f[l - 1] < x[j] && j != n - 1 )
                    {
                        H = H - i_sp[l, j - 1] * (x[j] - x_f[l - 1]) * 1000;
                        P = (H - Math.Pow(v[j], 2) / (2 * g) - z[j]) * ro_rl[l - 1, j] * g;

                        dataGridView9.Rows[kol + 1 - l].Cells[6].Value = H.ToString();
                        dataGridView9.Rows[kol + 1 - l].Cells[8].Value = (P * Math.Pow(10, -6)).ToString();

                        double zs = z[j - 1] + (z[j] - z[j - 1]) / (x[j] - x[j - 1]) * (x_f[l - 1] - x[j - 1]);
                        dataGridView7.Rows.Add(x_f[l - 1].ToString("0.0000"), zs, D[j], sig[j], Tk[l, j],
                            ro_rl[l - 1, j], nu_rl[l - 1, j], v[j], eps_sp[j], Re[l - 1, j], lyam_sp[l - 1, j], i_sp[l - 1, j], H, (P * Math.Pow(10, -6)));

                        P = P;
                        H = P / (ro_rl[l, j] * g) + Math.Pow(v[j], 2) / (2 * g) + z[j];

                        dataGridView9.Rows[kol - l].Cells[5].Value = H.ToString();
                        dataGridView9.Rows[kol - l].Cells[7].Value = (P * Math.Pow(10, -6)).ToString();

                        dataGridView7.Rows.Add(x_s[l].ToString("0.0000"), zs, D[j], sig[j], Tk[l, j],
                            ro_rl[l, j], nu_rl[l, j], v[j], eps_sp[j], Re[l, j], lyam_sp[l, j], i_sp[l, j], H, (P * Math.Pow(10, -6)));
                    }

                    if (j <= n - 1)
                    {
                        //MessageBox.Show(l.ToString(), "Чтение файла", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        H = H - i_sp[l, j - 1] * (x[j] - x_s[l]) * 1000;
                        P = (H - Math.Pow(v[j], 2) / (2 * g) - z[j]) * ro_rl[l, j] * g;
                    }
                    else
                    {
                        H = H - i_sp[l, j - 1] * (deltax) * 1000;
                        P = (H - Math.Pow(v[j], 2) / (2 * g) - z[j]) * ro_rl[l, j] * g;
                        dataGridView9.Rows[kol - l].Cells[6].Value = H.ToString();
                        dataGridView9.Rows[kol - l].Cells[8].Value = (P * Math.Pow(10, -6)).ToString();
                    }
                    dataGridView7.Rows.Add(x[j].ToString("0.0000"), z[j], D[j], sig[j], Tk[l, j],
                        ro_rl[l, j], nu_rl[l, j], v[j], eps_sp[j], Re[l, j], lyam_sp[l, j], i_sp[l, j], H, (P * Math.Pow(10, -6)));
                }
            }


            //Перерисовка графика
            double x_rl = 0;
            //Определяем последнюю закачиваемую в трубопровод жидкость
            i = 0;
            for (i = 0; i < dataGridView6.Rows.Count - 1; i++)
                if (x_s[i] == 0 && mass[i] > 0)
                    break;
            l = i;
            //MessageBox.Show(l.ToString(), "Чтение файла", MessageBoxButtons.OK, MessageBoxIcon.Information);

            foreach (var series in chart1.Series)
            {
                series.Points.Clear();
            }

            var chart = chart1.ChartAreas[0];

            for (j = 0; j < dataGridView7.Rows.Count - 1; j++)
            {
                x_rl = Convert.ToDouble(dataGridView7.Rows[j].Cells[0].Value);
                if (x_s[l] <= x_rl && x_f[l] >= x_rl)
                {
                    chart1.Series[fuels[l] + "_H"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[12].Value));
                    chart1.Series[fuels[l] + "_P"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[13].Value));
                    chart1.Series["z"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[1].Value));
                }
                else
                {
                    chart1.Series[fuels[l] + "_H"].Points.AddXY(x_f[l], Convert.ToDouble(dataGridView7.Rows[j].Cells[12].Value));
                    chart1.Series[fuels[l] + "_P"].Points.AddXY(x_f[l], Convert.ToDouble(dataGridView7.Rows[j].Cells[13].Value));
                    chart1.Series["z"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[1].Value));

                    if (j != n - 1)
                    {
                        if (l < dataGridView6.Rows.Count - 2)
                            l++;
                        else
                            l = 0;
                        chart1.Series[fuels[l] + "_H"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[12].Value));
                        chart1.Series[fuels[l] + "_P"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[13].Value));
                        chart1.Series["z"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[1].Value));
                    }
                }
            }
            

                //Thread.Sleep(1000);

            /*
            //Обработка временного события
            void Count(object obj)
            {
                int i = 0;
                //Определяем первую закачиваемую в трубопровод жидкость
                for (i = 0; i < dataGridView6.Rows.Count - 1; i++)
                    if (x_f[i] == m_x[dataGridView5.Rows.Count - 2] && mass[i] > 0)
                        break;
                mass[i] = mass[i] - Q * ro[i] / 1000 * k;


                //Перерасчет положения 
                x_f[dataGridView6.Rows.Count - 2] = m_x[dataGridView5.Rows.Count - 2];
                x_s[dataGridView6.Rows.Count - 2] = x_f[dataGridView6.Rows.Count - 2];

                i = dataGridView6.Rows.Count - 2;
                int j = n - 2;

                for (j = n - 2; j >= 0; j--)
                    if (mass[i] - ro_rl[i, j] * (deltax) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2) >= 0)
                    {
                        if (x_s[i] - deltax >= 0)
                        {
                            mass[i] = mass[i] - ro_rl[i, j] * (deltax) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2);
                            x_s[i] = x_s[i] - deltax;
                        }
                        else
                        {
                            x_s[i] = 0;
                            mass[i] = mass[i] - ro_rl[i, j] * (x_s[i]) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2);
                        }
                    }
                    else
                    {
                        if (x_s[i] - (deltax) * (mass[i]) / (ro_rl[i, j] * (deltax) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2)) >= 0)
                            x_s[i] = x_s[i] - (deltax) * (mass[i]) / (ro_rl[i, j] * (deltax) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2));
                        if (i > 0)
                        {
                            i--;
                            x_f[i] = x_s[i + 1];
                            x_s[i] = x_f[i] - (deltax) * (1 - mass[i + 1] / (ro_rl[i + 1, j] * (deltax) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2)));
                            mass[i + 1] = 0;
                            mass[i] = mass[i] - ro_rl[i, j] * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2) * ((deltax) * (1 - mass[i + 1] / (ro_rl[i + 1, j] * (deltax) * pi / 4 * Math.Pow((D[j] - 2 * sig[j]) / 1000, 2))));
                        }
                    }
                dataGridView9.Invoke
                    (
                        (ThreadStart)delegate ()
                        {
                            dataGridView9.Rows.Clear();
                            for (i = dataGridView6.Rows.Count - 2; i >= 0; i--)
                            {
                                //mass[i] = Convert.ToDouble(dataGridView6.Rows[i].Cells[4].Value);
                                dataGridView9.Rows.Add(fuels[i], mass[i], vol[i],
                                x_s[i], x_f[i], 0, 0, 0, 0);
                            }
                        }
                    );
                //Перасчет в новых координатах сечений
                int kol = dataGridView6.Rows.Count - 2;
                double H = 0, P = 0;
                //Определяем последнюю закачиваемую в трубопровод жидкость
                for (i = 0; i < dataGridView6.Rows.Count - 1; i++)
                    if (x_s[i] == 0 && mass[i] > 0)
                        break;
                int l = i;

                dataGridView7.Invoke
                    (
                        (ThreadStart)delegate ()
                        {
                            dataGridView7.Rows.Clear();
                        }
                    );

                for (j = 0; j < n; j++)
                {
                    if (x_s[l] <= x[j] && x_f[l] >= x[j])
                    {
                        if (j == 0)
                        {
                            H = Convert.ToDouble(textBox14.Text);

                        }
                        else
                        {
                            //MessageBox.Show( (dataGridView7.Rows[j - 1].Cells[10].Value).ToString(), "Чтение файла", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            H = H - i_sp[l, j - 1] * deltax * 1000;
                        }
                        P = (H - Math.Pow(v[j], 2) / (2 * g) - z[j]) * ro_rl[l, j] * g;

                        if (x[j] == x_s[l])
                        {
                            dataGridView9.Invoke
                            (
                                (ThreadStart)delegate ()
                                {
                                    MessageBox.Show(kol.ToString() + " " + l.ToString(), "Данные", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    dataGridView9.Rows[kol - l].Cells[5].Value = H.ToString();
                                    dataGridView9.Rows[kol - l].Cells[7].Value = (P * Math.Pow(10, -6)).ToString();
                                }
                            );
                        }
                        dataGridView7.Invoke
                        (
                            (ThreadStart)delegate ()
                            {
                                dataGridView7.Rows.Add(x[j].ToString("0.0000"), z[j], D[j], sig[j], Tk[l, j],
                                    ro_rl[l, j], nu_rl[l, j], v[j], eps_sp[j], Re[l, j], lyam_sp[l, j], i_sp[l, j], H, (P * Math.Pow(10, -6)));
                            }
                        );
                    }
                    else
                    {
                        if (j != n - 1)
                            l++;
                        if (x_s[l] <= x[j] && x_f[l - 1] < x[j] && j != n - 1)
                        {
                            H = H - i_sp[l, j - 1] * (x[j] - x_f[l - 1]) * 1000;
                            P = (H - Math.Pow(v[j], 2) / (2 * g) - z[j]) * ro_rl[l - 1, j] * g;
                            dataGridView9.Invoke
                            (
                                (ThreadStart)delegate ()
                                {
                                    dataGridView9.Rows[kol + 1 - l].Cells[6].Value = H.ToString();
                                    dataGridView9.Rows[kol + 1 - l].Cells[8].Value = (P * Math.Pow(10, -6)).ToString();
                                }
                            );

                            double zs = z[j - 1] + (z[j] - z[j - 1]) / (x[j] - x[j - 1]) * (x_f[l - 1] - x[j - 1]);
                            dataGridView7.Invoke
                            (
                                (ThreadStart)delegate ()
                                {
                                    dataGridView7.Rows.Add(x_f[l - 1].ToString("0.0000"), zs, D[j], sig[j], Tk[l, j],
                                        ro_rl[l - 1, j], nu_rl[l - 1, j], v[j], eps_sp[j], Re[l - 1, j], lyam_sp[l - 1, j], i_sp[l - 1, j], H, (P * Math.Pow(10, -6)));
                                }
                            );

                            P = P;
                            H = P / (ro_rl[l, j] * g) + Math.Pow(v[j], 2) / (2 * g) + z[j];
                            dataGridView9.Invoke
                            (
                                (ThreadStart)delegate ()
                                {
                                    dataGridView9.Rows[kol - l].Cells[5].Value = H.ToString();
                                    dataGridView9.Rows[kol - l].Cells[7].Value = (P * Math.Pow(10, -6)).ToString();
                                }
                            );
                            dataGridView7.Invoke
                            (
                                (ThreadStart)delegate ()
                                {
                                    dataGridView7.Rows.Add(x_s[l].ToString("0.0000"), zs, D[j], sig[j], Tk[l, j],
                                        ro_rl[l, j], nu_rl[l, j], v[j], eps_sp[j], Re[l, j], lyam_sp[l, j], i_sp[l, j], H, (P * Math.Pow(10, -6)));
                                }
                            );
                        }

                        if (j != n - 1)
                        {
                            H = H - i_sp[l, j - 1] * (x[j] - x_s[l]) * 1000;
                            P = (H - Math.Pow(v[j], 2) / (2 * g) - z[j]) * ro_rl[l, j] * g;
                        }
                        else
                        {
                            H = H - i_sp[l, j - 1] * (deltax) * 1000;
                            P = (H - Math.Pow(v[j], 2) / (2 * g) - z[j]) * ro_rl[l, j] * g;
                            dataGridView9.Invoke
                            (
                                (ThreadStart)delegate ()
                                {
                                    dataGridView9.Rows[kol - l].Cells[6].Value = H.ToString();
                                    dataGridView9.Rows[kol - l].Cells[8].Value = (P * Math.Pow(10, -6)).ToString();
                                }
                            );
                        }
                        dataGridView7.Invoke
                            (
                                (ThreadStart)delegate ()
                                {
                                    dataGridView7.Rows.Add(x[j].ToString("0.0000"), z[j], D[j], sig[j], Tk[l, j],
                                        ro_rl[l, j], nu_rl[l, j], v[j], eps_sp[j], Re[l, j], lyam_sp[l, j], i_sp[l, j], H, (P * Math.Pow(10, -6)));
                                }
                            );
                    }
                }
                //Перерисовка графика
                double x_rl = 0;
                //Определяем последнюю закачиваемую в трубопровод жидкость
                i = 0;
                for (i = 0; i < dataGridView6.Rows.Count - 1; i++)
                    if (x_s[i] == 0 && mass[i] > 0)
                        break;
                l = i;
                chart1.Invoke
                    (
                        (ThreadStart)delegate ()
                        {
                            foreach (var series in chart1.Series)
                            {
                                series.Points.Clear();
                            }

                            var chart = chart1.ChartAreas[0];

                            for (j = 0; j < dataGridView7.Rows.Count - 1; j++)
                            {
                                x_rl = Convert.ToDouble(dataGridView7.Rows[j].Cells[0].Value);
                                if (x_s[l] <= x_rl && x_f[l] >= x_rl)
                                {
                                    chart1.Series[fuels[l] + "_H"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[12].Value));
                                    chart1.Series[fuels[l] + "_P"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[13].Value));
                                    chart1.Series["z"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[1].Value));
                                }
                                else
                                {
                                    chart1.Series[fuels[l] + "_H"].Points.AddXY(x_f[l], Convert.ToDouble(dataGridView7.Rows[j].Cells[12].Value));
                                    chart1.Series[fuels[l] + "_P"].Points.AddXY(x_f[l], Convert.ToDouble(dataGridView7.Rows[j].Cells[13].Value));
                                    chart1.Series["z"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[1].Value));

                                    if (j != n - 1)
                                    {
                                        l++;
                                        chart1.Series[fuels[l] + "_H"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[12].Value));
                                        chart1.Series[fuels[l] + "_P"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[13].Value));
                                        chart1.Series["z"].Points.AddXY(x_rl, Convert.ToDouble(dataGridView7.Rows[j].Cells[1].Value));
                                    }
                                }
                            }
                        }
                    );
            }
            */
        }
    }
}
