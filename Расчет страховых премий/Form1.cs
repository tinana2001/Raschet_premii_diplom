
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
//using System.Windows.Forms.DataVisualization.Charting;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using Application = System.Windows.Forms.Application;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using System.Runtime.InteropServices;
namespace Расчет_страховых_премий
{
    public partial class Form1 : Form
    {
        public int age;//возраст человека
        public int w = 101; //предельный возраст
        double[] probty = new double[100];
        public int N; //количество челов
        public double razor;
        public double razor2;
        public int n; //срок договора
        public double premia; //премия
        public double premia2; //премия2
        public Form1()
        {
            InitializeComponent();
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            var excelBook = excelApp.Workbooks.Open(@"E:\ВКР\Башкирия_2021М_U.xlsx");
            var excelSheet = (Excel.Worksheet)excelBook.Sheets[1];
            var lastrowR = excelSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            var lastrowC = excelSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

            dataGridView1.Columns.Add("AGE", "Возраст x");
            dataGridView1.Columns.Add("N", "Количество доживших до возраста x");


            for (int j = 0; j <= lastrowR; j++)
            {
                dataGridView1.Rows.Add();
            }

            for (int x = 0; x <= 1; x++)
            {
                for (int y = 0; y <= 100; y++)
                {
                    dataGridView1.Rows[y].Cells[x].Value = excelSheet.Cells[y + 1, x + 1].Value.ToString();
                }
            }

            Excel.Range excelRange = excelSheet.Range["A1:B2"];
            Marshal.ReleaseComObject(excelRange);
            excelBook.Close();
            excelApp.Quit();

            if (excelSheet != null) Marshal.ReleaseComObject(excelSheet);
            if (excelApp != null) excelApp.Quit();


            Marshal.ReleaseComObject(excelSheet);
            Marshal.ReleaseComObject(excelBook);
            Marshal.ReleaseComObject(excelApp);
            GC.Collect();
        }

        public void probability(int age)
        {
            //w - age - 1
            for (int i = 0; i < w - age - 1; i++)
            {
                probty[i] = (double)(Int32.Parse(dataGridView1[1, age + i].Value.ToString()) - Int32.Parse(dataGridView1[1, age + i + 1].Value.ToString())) / Int32.Parse(dataGridView1[1, age].Value.ToString()); //(l(x+k)-l(x+k+1))/l(x)
            }
        }

        void Calculate()
        {
            textBox_calc.Text = "1";
        }

        private void textBox_age_TextChanged(object sender, EventArgs e)
        {
            if (int.TryParse(textBox_age.Text, out int result))
            {
                age = Int32.Parse(textBox_age.Text);
            }
            else
            {
                MessageBox.Show("Ошибка ввода данных. Пожалуйста, введите целое число.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox_age.Text = "1";// Очищаем поле ввода, чтобы пользователь мог повторно ввести значение  
            }
            if (double.Parse(textBox_age.Text) >= 100)
            {
                MessageBox.Show("Ошибка ввода данных. Слишком большой возраст страхователя.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox_age.Text = "1";
            }
        }

        private void textBox_years_TextChanged(object sender, EventArgs e)
        {
            if (int.TryParse(textBox_years.Text, out int result))
            {
                n = Int32.Parse(textBox_years.Text);
            }
            else
            {
                MessageBox.Show("Ошибка ввода данных. Пожалуйста, введите целое число.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox_years.Text = "1";
            }
        }

        private void textBox_insuredCount_TextChanged(object sender, EventArgs e)
        {
            if (int.TryParse(textBox_insuredCount.Text, out int result))
            {
                N = Int32.Parse(textBox_insuredCount.Text);
            }
            else
            {
                MessageBox.Show("Ошибка ввода данных. Пожалуйста, введите целое число.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox_insuredCount.Text = "1";

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {

            if (textBox1.Text == "" || textBox_age.Text == "" || textBox_years.Text == "" || textBox_insuredCount.Text == "" || textBox_premia.Text == "")
            {
                MessageBox.Show("Ошибка ввода данных. Пожалуйста, заполните недостающие данные.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox_premia.Text = "0";
                textBox1.Text = "0,01";
                textBox_age.Text = "1";
                textBox_years.Text = "1";
                textBox_insuredCount.Text = "1";
                textBox2.Text = "40";
                textBox3.Text = "2000";
                textBox4.Text = "300";
                textBox6.Text = "40000";
            }
            int age_v = Convert.ToInt32(textBox_age.Text);
            int years_v = Convert.ToInt32(textBox_years.Text);

            if ((age_v + years_v) >= 110)
            {
                MessageBox.Show("Ошибка ввода данных. Слишком большой срок страхования", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            if (textBox_age.Text == "0")
            {
                MessageBox.Show("Ошибка ввода данных. Пожалуйста, введите возраст больше или равный 1 году.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
            }
            
            double a = double.Parse(textBox1.Text);
            Calculate();
            while (double.Parse(textBox_calc.Text) > a)
            {
                button1_Click(sender, e);
               /* if (textBox_calc.Text == "0")
                {
                    textBox_calc.Text = "0.01";
                }*/
            }
        }

        private void textBox_premia_TextChanged(object sender, EventArgs e)
        {
            if (double.TryParse(textBox_premia.Text, out double result))
            {
                string s = textBox_premia.Text;
                premia = Convert.ToDouble(s);
            }
            else
            {
                MessageBox.Show("Ошибка ввода данных. Пожалуйста, введите размер премии целым числом", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox_premia.Text = "0";
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (double.TryParse(textBox1.Text, out double result))
            {
                string l = textBox1.Text;
                razor = Convert.ToDouble(l);
            }
            else
            {
                MessageBox.Show("Ошибка ввода данных. Пожалуйста, введите десятичное  число", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox1.Text = "0,01";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int k = 1000;

            double premia2 = double.Parse(textBox_premia.Text);
            double Smin = 0;
            double Smax = 100000;
            double a = double.Parse(textBox1.Text);
            double Sleft = Smin;
            double Sright = Smax;
            Random rnd = new Random();
            probability(age); //здесь мы высчитываем q, оно одинаково для всех экспериментов
            int[] K = new int[N]; // массив остатков лет для каждого человека
            int[] death = new int[5000]; //массив, считать количество людей кто умрет в 1 год, 2, 3...
            int ll = int.Parse(textBox_insuredCount.Text);
            int[] live = new int[5000]; //количество живых
                                        //double akz = double.Parse(label8.Text);
                                        //double obl = double.Parse(label9.Text);
                                        // double doh = akz + obl;
            double premia3 = double.Parse(textBox_premia.Text);
            double ro = 0; //счетчик разорений
            double ro2 = 0;


            /****************************************************************************/
            /******************************************************************************/
            int kk = 100;
            double[] obll = new double[kk];
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;
            var excelBook = excelApp.Workbooks.Open(@"E:\ВКР\obl11.xlsx");
            var excelSheet = (Excel.Worksheet)excelBook.Sheets[1];
            var lastrowR = excelSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            var lastrowC = excelSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;


            double[,] array = new double[lastrowR, lastrowC];

            for (int i = 1; i <= lastrowR; i++)
            {
                for (int j = 1; j <= lastrowC; j++)
                {
                    double.TryParse((string)(excelSheet.Cells[i, j] as Excel.Range).Value, out array[i - 1, j - 1]);
                }
            }

            // Вычисление относительных частот
            //int k = 10; // количество интервалов
            double roundedNum;
            roundedNum = (1 + 3.322 * Math.Log10(lastrowR)); //получаем приблизительное количество интервалов по формуле Стерджиса
            int kl = (int)Math.Round(roundedNum);

            if (roundedNum < kl) { kl = kl + 1; }
            double[] xi = new double[kl];
            double[] wi = new double[kl];
            double min = array[0, 0];
            double max = array[0, 0];

            for (int i = 0; i < lastrowR; i++)
            {
                if (array[i, 0] < min) min = array[i, 0];
                if (array[i, 0] > max) max = array[i, 0];
            }

            double hh = (max - min) / kl; //Длина каждого частичного интервала
            int h = (int)Math.Round(hh);
            if (h < hh) { h = h + 1; }
            double[] midpoints = new double[h];
            for (int i = 0; i < kl; i++)
            {
                xi[i] = min + (i + 1) * h;

            }

            for (int i = 0; i < lastrowR; i++)
            {
                for (int j = 0; j < kl; j++)
                {
                    if ((array[i, 0] >= min + j * h) && (array[i, 0] < min + (j + 1) * h))
                    {
                        wi[j]++; //увеличиваем счетчик попадания в интервал

                        break;
                    }
                }
            }

            //double sum = 0;
            for (int i = 0; i < kl; i++)
            {
                wi[i] /= lastrowR;
                dataGridView2.Rows.Add(xi[i], wi[i]);
                // sum += wi[i];
            }
            double[] wii = new double[k];
            //рандом 
            double x = 0;

            Random rndd = new Random();
            //цикл по отрезкам 
            //если z<= p1, то z=x1 (вероятность в 1 интервале) 
            //иначе смотрим z<= p1+p2, если да, то z=x2 (вероятность во 2 интервале)
            for (int i = 0; i < kl; i++) { wii[i] = wi[i]; }

            for (int p = 0; p < kk; p++)
            {
                double z = rndd.NextDouble();
                for (int i = 0; i < kl; i++)
                {

                    if (z <= wii[i])
                    {
                        x = wi[i]; //искомое значение для акций
                        break;
                    }
                    else
                    {
                        wii[i + 1] = wii[i + 1] + wii[i];
                    }

                }

                obll[p] = 0.3 * x;
            }

            /*************************************************************************************/
            /**********************************************************************************/






            /****************************************************************************/
            /******************************************************************************/

            double[] akzz = new double[kk];
            excelApp.Visible = false;
            var excelBook2 = excelApp.Workbooks.Open(@"E:\ВКР\akz111.xlsx");
            var excelSheet2 = (Excel.Worksheet)excelBook2.Sheets[1];
            var lastrowR2 = excelSheet2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            var lastrowC2 = excelSheet2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;


            double[,] array2 = new double[lastrowR2, lastrowC2];

            for (int i = 1; i <= lastrowR2; i++)
            {
                for (int j = 1; j <= lastrowC2; j++)
                {
                    //array[i - 1, j - 1] = (double)(usedRange.Cells[i, j] as Excel.Range).Value;
                    double.TryParse((string)(excelSheet2.Cells[i, j] as Excel.Range).Value, out array2[i - 1, j - 1]);
                }
            }

            // Вычисление относительных частот
            //int k = 10; // количество интервалов
            double roundedNum2;
            roundedNum2 = (1 + 3.322 * Math.Log10(lastrowR2)); //получаем приблизительное количество интервалов по формуле Стерджиса
            int kl2 = (int)Math.Round(roundedNum2);

            if (roundedNum2 < kl2) { kl2 = kl2 + 1; }
            double[] xi2 = new double[kl];
            double[] wi2 = new double[kl];
            double min2 = array2[0, 0];
            double max2 = array2[0, 0];

            for (int i = 0; i < lastrowR2; i++)
            {
                if (array2[i, 0] < min2) min2 = array2[i, 0];
                if (array2[i, 0] > max2) max2 = array2[i, 0];
            }

            double hh2 = (max2 - min2) / kl2; //Длина каждого частичного интервала
            int h2 = (int)Math.Round(hh2);
            if (h2 < hh2) { h2 = h2 + 1; }
            double[] midpoints2 = new double[h2];
            for (int i = 0; i < kl2; i++)
            {
                xi2[i] = min2 + (i + 1) * h2;

            }

            for (int i = 0; i < lastrowR2; i++)
            {
                for (int j = 0; j < kl2; j++)
                {
                    if ((array2[i, 0] >= min2 + j * h2) && (array2[i, 0] < min2 + (j + 1) * h2))
                    {
                        wi2[j]++; //увеличиваем счетчик попадания в интервал

                        break;
                    }
                }
            }

            double sum2 = 0;
            for (int i = 0; i < kl2; i++)
            {
                wi2[i] /= lastrowR2;
                dataGridView3.Rows.Add(xi2[i], wi2[i]);
                sum2 += wi2[i];
            }
            double[] wii2 = new double[k];
            //рандом 
            double x2 = 0;

            Random rndd2 = new Random();
            //цикл по отрезкам 
            //если z<= p1, то z=x1 (вероятность в 1 интервале) 
            //иначе смотрим z<= p1+p2, если да, то z=x2 (вероятность во 2 интервале)
            for (int i = 0; i < kl2; i++) { wii2[i] = wi2[i]; }

            for (int p = 0; p < kk; p++)
            {
                double z2 = rndd2.NextDouble();
                for (int i = 0; i < kl2; i++)
                {

                    if (z2 <= wii2[i])
                    {
                        x2 = wi2[i]; //искомое значение для акций
                        break;
                    }
                    else
                    {
                        wii2[i + 1] = wii2[i + 1] + wii2[i];
                    }

                }

                akzz[p] = 0.4 * x2;
            }

            /*************************************************************************************/
            /**********************************************************************************/







            //эксперименты от 1 до 1000
            for (int expt = 1; expt <= k; ++expt)
            {   //считаем остатки жизней K(x)



                for (int j = 0; j < N; ++j)
                {
                    double z = rnd.NextDouble(); //рандомное число от 0 до 1
                    double s = probty[0]; int l = 0;
                    while (l < w - age - 1)
                    {
                        if (z < s) { K[j] = l; break; } //сравниваем с суммой вероятностей q,
                                                        //если все ок, то выходим
                        else
                        {
                            l = l + 1;                        //иначе добавляем следующее q в сумму и потом будем сравнивать
                            s = s + probty[l];
                        }
                    }
                }

                //считаем количество умерших в каждом году
                for (int i = 0; i < N; ++i)
                {
                    //if (K[i] > n) { ++death[n]; }
                    for (int j = 0; j < n; ++j)
                    {

                        if (K[i] == j) //сравниваем количество остатков лет каждого с годом договора,
                                       //если совпадает, то в массиве, в котором считаются умершие, увеличиваем число
                        {
                            ++death[i];
                        }
                        else { ++live[i]; }
                    }

                }





                // int S = 100000;
                int S = Convert.ToInt32(textBox8.Text);
                int b = Convert.ToInt32(textBox9.Text);
                //double S = 350000;
                premia = S / N;
                double u = N * premia;
                double d1 = Convert.ToDouble(textBox2.Text);
                int d2 = Convert.ToInt32(textBox3.Text);
                int d3 = Convert.ToInt32(textBox4.Text);
                int d4 = Convert.ToInt32(textBox6.Text);
                //int incom = d1 + d2 + d3 + d4;
                //вычисляем капитал
                for (int j = 0; j < n; ++j)
                {
                    //double incom = d1 *live[j]*premia2 + d2 + d3*death[j]*S + d4;
                    double incom = d1 * live[j] * premia + d2 + d3 * death[j] + d4;
                    S = S + b;
                    //u = u * (1 + 0.08) - S * death[j]; //u=u+live[j]*premia2+доход-S * death[j]-расходы
                    // u = u + live[j] * premia2 + doh - S * death[j];
                    //u = u - live[j] * premia2 - incom + (obll[j] + akzz[j] + (1 - (0.3 + 0.4))) + S * death[j] ;
                    u = u + live[j] * premia2 + (obll[j] + akzz[j] + (1 - (0.3 + 0.4))) * 8.291 - S * death[j] - incom;
                    // u2= u2 - live[j] * premia2 - (obll[j] + akzz[j] + (1 - (0.3 + 0.4))) + incom;
                    //u = u - n * premia2 - obll[j] + akzz[j] + (0.2 * 0.3) + 0.1 + incom;
                    if (u <= 0) { ++ro; break; }
                    // if (u2 >= 0) { ++ro2; break; }
                }
                for (int i = 0; i < N; K[i++] = 0) { } //очищаем массивы, а то они накапливаются при каждом эксперименте
                for (int i = 0; i < n; death[i++] = 0) { }
                for (int i = 0; i < n; live[i++] = 0) { }
            
                double R = (double)ro / k;
                // double S1 = (Smin + Smax) / 2;



                textBox_calc.Text = R.ToString();
                

                if (R > a) // если вероятность больше a, отсекаем правую часть отрезка
                    {
                        Smax = (Smin + Smax) / 2;
                        premia2 = premia2 + 10;
                        textBox_premia.Text = premia2.ToString();
                        textBox5.Text = premia2.ToString();

                    }
                    else // иначе отсекаем левую часть отрезка
                    {
                        Smin = (Smin + Smax) / 2;
                   
                    }
                
            }

           
            Excel.Range excelRange = excelSheet.Range["A1:B2"];
            Marshal.ReleaseComObject(excelRange);
            excelBook.Close();
            excelApp.Quit();

            if (excelSheet != null) Marshal.ReleaseComObject(excelSheet);
            if (excelSheet2 != null) Marshal.ReleaseComObject(excelSheet2);
            if (excelApp != null) excelApp.Quit();


            Marshal.ReleaseComObject(excelSheet);
            Marshal.ReleaseComObject(excelSheet2);
            Marshal.ReleaseComObject(excelBook);
            Marshal.ReleaseComObject(excelBook2);
            Marshal.ReleaseComObject(excelApp);
            GC.Collect();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }










        private void button3_Click_1(object sender, EventArgs e)
        {
            Random pic = new Random();
            double pic_v = pic.NextDouble() * (0.09 - 0.01) + 0.01;
            textBox1.Text = pic_v.ToString();

            Random age = new Random();
            int age_v = age.Next(10, 100);
            textBox_age.Text = age_v.ToString();

            Random years = new Random();
            int years_v = years.Next(0, 90);
            textBox_years.Text = years_v.ToString();

            Random insuredCount = new Random();
            int insuredCount_v = insuredCount.Next(1, 1000);
            textBox_insuredCount.Text = insuredCount_v.ToString();

            textBox_premia.Text = "0";

            if (age_v + years_v > 100)
            {
                Random years1 = new Random();
                int years1_v = years1.Next(0, 25);
                textBox_years.Text = years1_v.ToString();

                Random age1 = new Random();
                int age1_v = age1.Next(10, 70);
                textBox_age.Text = age1_v.ToString();
            }


            Random doh1 = new Random();
            int doh1_v = doh1.Next(0, 100);
            textBox2.Text = doh1_v.ToString();

            Random doh2 = new Random();
            int doh2_v = doh2.Next(0, 5000);
            textBox3.Text = doh2_v.ToString();

            Random doh3 = new Random();
            int doh3_v = doh3.Next(0, 600);
            textBox4.Text = doh3_v.ToString();

            Random doh4 = new Random();
            int doh4_v = doh4.Next(0, 100000);
            textBox6.Text = doh4_v.ToString();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Загрузка данных из Excel
                Excel.Application excel = new Excel.Application();
                Excel.Workbook workbook = excel.Workbooks.Open(openFileDialog.FileName);
                Excel.Worksheet worksheet = workbook.Sheets[1];
                Excel.Range usedRange = worksheet.UsedRange;

                // Вывод данных в DataGridView
                int rowsCount = usedRange.Rows.Count;
                int columnsCount = usedRange.Columns.Count;
                double[,] array = new double[rowsCount, columnsCount];

                for (int i = 1; i <= rowsCount; i++)
                {
                    for (int j = 1; j <= columnsCount; j++)
                    {
                        //array[i - 1, j - 1] = (double)(usedRange.Cells[i, j] as Excel.Range).Value;
                        double.TryParse((string)(usedRange.Cells[i, j] as Excel.Range).Value, out array[i - 1, j - 1]);
                    }
                }

                // Вычисление относительных частот
                //int k = 10; // количество интервалов
                double roundedNum;
                roundedNum = (1 + 3.322 * Math.Log10(rowsCount)); //получаем приблизительное количество интервалов по формуле Стерджиса
                int k = (int)Math.Round(roundedNum);
                if (roundedNum < k) { k = k + 1; }
                double[] xi = new double[k];
                double[] wi = new double[k];
                double min = array[0, 0];
                double max = array[0, 0];

                for (int i = 0; i < rowsCount; i++)
                {
                    if (array[i, 0] < min) min = array[i, 0];
                    if (array[i, 0] > max) max = array[i, 0];
                }

                double hh = (max - min) / k; //Длина каждого частичного интервала
                int h = (int)Math.Round(hh);
                if (h < hh) { h = h + 1; }
                double[] midpoints = new double[h];
                for (int i = 0; i < k; i++)
                {
                    xi[i] = min + (i + 1) * h;

                }

                for (int i = 0; i < rowsCount; i++)
                {
                    for (int j = 0; j < k; j++)
                    {
                        if ((array[i, 0] >= min + j * h) && (array[i, 0] < min + (j + 1) * h))
                        {
                            wi[j]++; //увеличиваем счетчик попадания в интервал

                            break;
                        }
                    }
                }

                double sum = 0;
                for (int i = 0; i < k; i++)
                {
                    wi[i] /= rowsCount;
                    dataGridView2.Rows.Add(xi[i], wi[i]);
                    sum += wi[i];
                }
                double[] wii = new double[k];
                //рандом 
                double x = 0;
                Random rnd = new Random();
                //цикл по отрезкам 
                //если z<= p1, то z=x1 (вероятность в 1 интервале) 
                //иначе смотрим z<= p1+p2, если да, то z=x2 (вероятность во 2 интервале)
                for (int i = 0; i < k; i++) { wii[i] = wi[i]; }


                double z = rnd.NextDouble();
                for (int i = 0; i < k; i++)
                {

                    if (z <= wii[i])
                    {
                        x = wi[i]; //искомое значение для акций
                        break;
                    }
                    else
                    {
                        wii[i + 1] = wii[i + 1] + wii[i];
                    }

                }
                double lumd = 0.2;
                double akz = lumd * x;
                // Закрытие файла Excel
                workbook.Close();
                excel.Quit();
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }
    }
}