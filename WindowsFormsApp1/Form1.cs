using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Timers;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.IO;

namespace WindowsFormsApp1
{
    
    public partial class Form1 : Form
    {
        Random rnd = new Random();
        int marriage = 0;
        Block[] arr_block = new Block[11];
        int complete = 0;


        //-----------------------------------------

        public class Block
        {
            public int time_max;

            public Block(int value)
            {
                time_max = value;
            }
            public int time_curr { get; set; } = 0;
            public bool work { get; set; } = false;
            public int value { get; set; } = 0;
            public int value1 { get; set; } = 0;
            public int value_comp { get; set; } = 0;

        }
        public Form1()
        {
            InitializeComponent();

            label10.Text = "0";
            label11.Text = "0";
            label12.Text = "0";
            label13.Text = "0";
            label14.Text = "0";
            label15.Text = "0";
            label16.Text = "0";
            label17.Text = "0";

            label21.Text = "0";

            label20.Text = marriage.ToString();


        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void check_foto(Block[] arr_block, int i)
        {
            arr_block[i].work = false;
            foto(arr_block, i);
        }

        private void process(Block[] arr_block, int i)
        {
            if (arr_block[i].work == true)
            {
                arr_block[i].time_curr++;

                if (arr_block[i].time_curr == arr_block[i].time_max)
                {
                    arr_block[i].value_comp++;
                    arr_block[i].time_curr = 0;
                    switch (i)
                    {
                        case 0:
                            {
                                arr_block[i].value--;
                                arr_block[i + 1].value++;
                                arr_block[i + 1].work = true;
                                foto(arr_block, i + 1);

                                if (arr_block[i].value == 0)
                                {
                                    check_foto(arr_block, i);
                                }
                                break;
                            }

                        case 1:
                            {
                                arr_block[i].value--;

                                arr_block[i + 1].value += 40;
                                arr_block[i + 2].value += 40;
                                arr_block[i + 3].value += 40;

                                arr_block[i + 1].work = true;
                                arr_block[i + 2].work = true;
                                arr_block[i + 3].work = true;

                                foto(arr_block, i + 1);
                                foto(arr_block, i + 2);
                                foto(arr_block, i + 3);

                                if (arr_block[i].value == 0)
                                {
                                    check_foto(arr_block, i);
                                }

                                
                                break;
                            }

                        case 2:
                            {
                                arr_block[i].value--;
                                arr_block[i].time_max = rnd.Next(8, 12);
                                arr_block[5].value++;
                                arr_block[5].work = true;
                                foto(arr_block, 5);

                                if (arr_block[i].value == 0)
                                {
                                    check_foto(arr_block, i);
                                }
                                break;
                            }
                        case 3:
                            {
                                arr_block[i].value--;
                                arr_block[i].time_max = rnd.Next(7, 10);
                                arr_block[6].value++;
                                arr_block[6].work = true;
                                foto(arr_block, 6);

                                if (arr_block[i].value == 0)
                                {
                                    check_foto(arr_block, i);
                                }
                                break;
                            }
                        case 4:
                            {
                                arr_block[i].value--;
                                arr_block[i].time_max = rnd.Next(8, 12);
                                arr_block[8].value1++;

                                if (arr_block[8].value1 > 0 && arr_block[8].value > 0)
                                {
                                    arr_block[8].work = true;
                                    foto(arr_block, 8);
                                }

                                if (arr_block[i].value == 0)
                                {
                                    check_foto(arr_block, i);
                                }

                                break;
                            }
                        case 5:
                            {
                                arr_block[i].time_max = rnd.Next(8, 12);
                                arr_block[i].value--;
                                arr_block[7].value++;

                                if (arr_block[7].value1 > 0 && arr_block[7].value > 0)
                                {
                                    arr_block[7].work = true;
                                    foto(arr_block, 7);
                                }

                                if (arr_block[i].value == 0)
                                {
                                    check_foto(arr_block, i);
                                }

                                break;
                            }
                        case 6:
                            {
                                arr_block[i].time_max = rnd.Next(7, 10);
                                if (arr_block[i].value_comp == 1)
                                {
                                    int q = 0;
                                    q++;
                                }
                                arr_block[i].value--;
                                arr_block[7].value1++;

                                if (arr_block[7].value1 > 0 && arr_block[7].value > 0)
                                {
                                    arr_block[7].work = true;
                                    foto(arr_block, 7);
                                }



                                if (arr_block[i].value == 0)
                                {
                                    check_foto(arr_block, i);
                                }

                                break;
                            }
                        case 7:
                            {
                                arr_block[i].time_max = rnd.Next(8, 12);
                                arr_block[i].value1--;
                                arr_block[i].value--;
                                arr_block[8].value++;

                                if (arr_block[8].value1 > 0 && arr_block[8].value > 0)
                                {
                                    arr_block[8].work = true;
                                    foto(arr_block, 8);
                                }

                                if (arr_block[i].value == 0 || arr_block[i].value1 == 0)
                                {
                                    check_foto(arr_block, i);
                                }

                                break;
                            }
                        case 8:
                            {
                                if (rnd.Next(1, 100) != 50)
                                {
                                    arr_block[i + 1].value++;
                                    arr_block[i + 1].work = true;
                                    foto(arr_block, i + 1);
                                }
                                else
                                {
                                    marriage++;
                                    label20.Text = marriage.ToString();
                                }

                                arr_block[i].time_max = rnd.Next(6, 7);
                                arr_block[i].value1--;
                                arr_block[i].value--;

                                if (arr_block[i].value == 0 || arr_block[i].value1 == 0)
                                {
                                    check_foto(arr_block, i);
                                }

                                break;
                            }
                        case 9:
                            {
                                arr_block[i].time_max = rnd.Next(4, 6);
                                arr_block[i].value--;

                                arr_block[i + 1].value++;

                                //foto(arr_block, i + 1);

                                if (arr_block[i].value == 0)
                                {
                                    check_foto(arr_block, i);
                                }

                                break;
                            }
                        case 10:
                            {
                                check_foto(arr_block, i);
                                complete++;

                                break;
                            }
                    }
                }
            }
        }

        private void foto(Block[] arr_block, int i)
        {
            switch (i)
            {
                case 0:
                    {
                        if (arr_block[i].work == true)
                        {
                            pictureBox1.Image = Properties.Resources.подвоз_комп_1;
                        }
                        else
                            pictureBox1.Image = Properties.Resources.подвоз_комп;
                        break;
                    }
                case 1:
                    {
                        if (arr_block[i].work == true)
                        {
                            pictureBox2.Image = Properties.Resources.Сорт_комп_1;
                        }
                        else
                            pictureBox2.Image = Properties.Resources.Сортировка;
                        break;
                    }
                case 2:
                    {
                        if (arr_block[i].work == true)
                        {
                            pictureBox3.Image = Properties.Resources.запаивание_1;
                        }
                        else
                            pictureBox3.Image = Properties.Resources.запаивание;
                        break;
                    }
                case 3:
                    {
                        if (arr_block[i].work == true)
                        {
                            pictureBox4.Image = Properties.Resources.сборка_кора_1;
                        }
                        else
                            pictureBox4.Image = Properties.Resources.Сборка_корп;
                        break;
                    }
                case 4:
                    {
                        if (arr_block[i].work == true)
                        {
                            pictureBox5.Image = Properties.Resources.Обработка_фиг_1;
                        }
                        else
                            pictureBox5.Image = Properties.Resources.Обработка_фиг;
                        break;
                    }
                case 5:
                    {
                        if (arr_block[i].work == true)
                        {
                            pictureBox6.Image = Properties.Resources.Уст_ПО_1;
                        }
                        else
                            pictureBox6.Image = Properties.Resources.Установка_ПО;
                        break;
                    }
                case 6:
                    {
                        if (arr_block[i].work == true)
                        {
                            pictureBox7.Image = Properties.Resources.Обработка_корп_1;
                        }
                        else
                            pictureBox7.Image = Properties.Resources.обработка_корп;
                        break;
                    }
                case 7:
                    {
                        if (arr_block[i].work == true)
                        {
                            pictureBox8.Image = Properties.Resources.установка_верхней_1;
                        }
                        else
                            pictureBox8.Image = Properties.Resources.установка_крышки;
                        break;
                    }
                case 8:
                    {
                        if (arr_block[i].work == true)
                        {
                            pictureBox9.Image = Properties.Resources.тестирование_1;
                        }
                        else
                            pictureBox9.Image = Properties.Resources.Тестирование;
                        break;
                    }
                case 9:
                    {
                        if (arr_block[i].work == true)
                        {
                            pictureBox10.Image = Properties.Resources.Конеч_упак_1;
                        }
                        else
                            pictureBox10.Image = Properties.Resources.Конеч_упаковка_доски;
                        break;
                    }
                case 10:
                    {
                        if (arr_block[i].work == true)
                        {
                            pictureBox11.Image = Properties.Resources.вывоз_1;
                        }
                        else
                            pictureBox11.Image = Properties.Resources.вывоз_на_склад;
                        break;
                    }
            }
        }

        private async void button1_Click(object sender, EventArgs e)
        {

            button1.Enabled = false;
            int time = 0;


            arr_block[0] = new Block(rnd.Next(15, 18));
            arr_block[1] = new Block(rnd.Next(10, 15));
            arr_block[2] = new Block(rnd.Next(8, 12));
            arr_block[3] = new Block(rnd.Next(7, 11));
            arr_block[4] = new Block(rnd.Next(8, 12));
            arr_block[5] = new Block(rnd.Next(8, 12));
            arr_block[6] = new Block(rnd.Next(6, 10));
            arr_block[7] = new Block(rnd.Next(8, 12));
            arr_block[8] = new Block(rnd.Next(5, 7));
            arr_block[9] = new Block(rnd.Next(4, 6));
            arr_block[10] = new Block(rnd.Next(5, 7));

            arr_block[0].value = 1;
            arr_block[0].work = true;
            pictureBox1.Image = Properties.Resources.подвоз_комп_1;

            while (time < 460)
            {
                Vremya(time);
                for (int i = 0; i <= 10; i++)
                {
                    process(arr_block, i);

                    await Task.Delay(1);
                }

                label10.Text = arr_block[2].value_comp.ToString();
                label11.Text = arr_block[3].value_comp.ToString();
                label12.Text = arr_block[4].value_comp.ToString();
                label13.Text = arr_block[5].value_comp.ToString();
                label14.Text = arr_block[6].value_comp.ToString();
                label15.Text = arr_block[7].value_comp.ToString();
                label16.Text = arr_block[8].value_comp.ToString();
                label17.Text = arr_block[9].value_comp.ToString();

                label21.Text = complete.ToString();

                time++;


            }

            pictureBox1.Image = Properties.Resources.подвоз_комп;

            pictureBox2.Image = Properties.Resources.Сортировка;

            pictureBox3.Image = Properties.Resources.запаивание;

            pictureBox4.Image = Properties.Resources.Сборка_корп;

            pictureBox5.Image = Properties.Resources.Обработка_фиг;

            pictureBox6.Image = Properties.Resources.Установка_ПО;

            pictureBox7.Image = Properties.Resources.обработка_корп;

            pictureBox8.Image = Properties.Resources.установка_крышки;

            pictureBox9.Image = Properties.Resources.Тестирование;

            pictureBox10.Image = Properties.Resources.Конеч_упаковка_доски;

            pictureBox11.Image = Properties.Resources.вывоз_1;

            
            while (time < 480)
            {
                await Task.Delay(100);
                Vremya(time);

                if (time == arr_block[10].time_max)
                {
                    pictureBox11.Image = Properties.Resources.вывоз_1;
                    break;
                }

                if(time == 479)
                {
                    pictureBox11.Image = Properties.Resources.вывоз_на_склад;
                }

                time++;
            }

            label21.Text = complete.ToString("1");


            try
            {
                using (ExcelHelper helper = new ExcelHelper())
                {
                    if (helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, "Info.xlsx")))
                    {
                        helper.Set(column: "A", row: 1, data: "Затрачено времени, мин");
                        //var val = helper.Get(column: "A", row: 6);
                        helper.Set(column: "A", row: 2, data: time);
                        helper.Set(column: "B", row: 1, data: "Использовано наборов комплектующих");
                        helper.Set(column: "B", row: 2, data: 40);
                        helper.Set(column: "C", row: 1, data: "Произведено продукции всего");
                        helper.Set(column: "C", row: 2, data: arr_block[9].value_comp);
                        helper.Set(column: "D", row: 1, data: "Произведено продукции 1 типа");
                        helper.Set(column: "D", row: 2, data: Math.Round((arr_block[9].value_comp*0.7),0));
                        helper.Set(column: "E", row: 1, data: "Произведено продукции 2 типа");
                        helper.Set(column: "E", row: 2, data: Math.Round((arr_block[9].value_comp * 0.3), 0));
                        helper.Set(column: "F", row: 1, data: "Брак всего");
                        helper.Set(column: "F", row: 2, data: marriage);
                        helper.Set(column: "G", row: 1, data: "Брак 1 типа");
                        helper.Set(column: "G", row: 2, data: Math.Round((marriage * 0.7), 0));
                        helper.Set(column: "H", row: 1, data: "Брак 2 типа");
                        helper.Set(column: "H", row: 2, data: Math.Round((marriage * 0.3), 0));

                        helper.Save();
                    }
                }

                Console.Read();
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }


        }
        private void Vremya(int time)
        {
            int min, hour;
            string noll = "";
            string nol = "";
            min = time % 60;
            hour = (time / 60)+8;

            if (min < 10)
                noll = "0";
            if (hour < 10)
                nol = "0";

            string clock_sring = $"   {nol}{hour}:{noll}{min}";

            label9.Text = clock_sring;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            arr_block[0].value = 1;
            arr_block[0].work = true;
            pictureBox1.Image = Properties.Resources.подвоз_комп_1;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            arr_block[10].work = true;
            foto(arr_block, 10);

            arr_block[10].time_max = rnd.Next(18, 20);

            arr_block[10].value -= 40;

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        public void OpenDoc()
        {
            
        }
    }

    class ExcelHelper : IDisposable
    {
        private Excel.Application _excel;
        private Excel.Workbook _workbook;
        private string _filePath;

        public ExcelHelper()
        {
            _excel = new Excel.Application();
        }

        internal bool Open(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    _workbook = _excel.Workbooks.Open(filePath);
                }
                else
                {
                    _workbook = _excel.Workbooks.Add();
                    _filePath = filePath;
                }

                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }

        internal void Save()
        {
            if (!string.IsNullOrEmpty(_filePath))
            {
                _workbook.SaveAs(_filePath);
                _filePath = null;
            }
            else
            {
                _workbook.Save();
            }
        }

        internal bool Set(string column, int row, object data)
        {
            try
            {
                // var val = ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column].Value2;

                ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column] = data;
                _excel.ActiveSheet.Columns.ColumnWidth = 35;
                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }

        internal object Get(string column, int row)
        {
            try
            {
                _excel.ActiveSheet.Columns.ColumnWidth = 35;
                return ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column].Value2;

            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return null;
        }

        public void Dispose()
        {
            try
            {
                _workbook.Close();
                _excel.Quit();
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}
