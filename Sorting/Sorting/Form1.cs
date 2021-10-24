using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Exc = Microsoft.Office.Interop.Excel;

namespace Sorting
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public void Excel()
        {
            try
            {
                string str;
                int rCnt;
                int cCnt;

                OpenFileDialog opf = new OpenFileDialog();
                opf.Filter = "Файл Excel|*.XLSX;*.XLS";
                opf.ShowDialog();
                System.Data.DataTable tb = new System.Data.DataTable();
                string filename = opf.FileName;

                Exc.Application ExcelApp = new Exc.Application();
                Exc._Workbook ExcelWorkBook;
                Exc.Worksheet ExcelWorkSheet;
                Exc.Range ExcelRange;

                ExcelWorkBook = ExcelApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Exc.XlPlatform.xlWindows, "\t", false,
                    false, 0, true, 1, 0);
                ExcelWorkSheet = (Exc.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

                ExcelRange = ExcelWorkSheet.UsedRange;
                for (rCnt = 1; rCnt <= ExcelRange.Rows.Count; rCnt++)
                {
                    dataGridView1.Rows.Add(1);
                    for (cCnt = 1; cCnt <= 2; cCnt++)
                    {
                        str = (string)(ExcelRange.Cells[rCnt, cCnt] as Exc.Range).Text;
                        dataGridView1.Rows[rCnt - 1].Cells[cCnt - 1].Value = str;
                    }
                }
                ExcelWorkBook.Close(true, null, null);
                ExcelApp.Quit();

                releaseObject(ExcelWorkSheet);
                releaseObject(ExcelWorkBook);
                releaseObject(ExcelApp);
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Невозможно очистить " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }

        }

        static void Swap(ref int e1, ref int e2)
        {
            var temp = e1;
            e1 = e2;
            e2 = temp;
        }



        public void Bubblesort()
        {
            string[] x;

            int rows = dataGridView1.Rows.Count;

            x = new string[dataGridView1.RowCount - 1];

            for (int i = 0; i < dataGridView1.RowCount - 1; ++i)
            {
                x[i] = dataGridView1[0, i].Value.ToString();

            }

            int[] array = x.Select(i => int.Parse(i)).ToArray();

            var len = array.Length;
            for (var i = 1; i < len; i++)
            {
                for (var j = 0; j < len - i; j++)
                {
                    if (array[j] > array[j + 1])
                    {
                        Swap(ref array[j], ref array[j + 1]);
                    }
                }
            }



            int N = dataGridView1.RowCount;

            for (int i = 0; i < dataGridView1.RowCount - 1; ++i)
            {
                dataGridView2.Rows.Add();
                dataGridView2[0, i].Value = array[i];
            }
        }

        public void Vstavka()
        {
            string[] x;

            int rows = dataGridView1.Rows.Count;

            x = new string[dataGridView1.RowCount - 1];

            for (int i = 0; i < dataGridView1.RowCount - 1; ++i)
            {
                x[i] = dataGridView1[0, i].Value.ToString();

            }

            int[] array = x.Select(i => int.Parse(i)).ToArray();

            for (var i = 1; i < array.Length; i++)
            {
                var key = array[i];
                var j = i;
                while ((j > 1) && (array[j - 1] > key))
                {
                    Swap(ref array[j - 1], ref array[j]);
                    j--;
                }

                array[j] = key;
            }

            int N = dataGridView1.RowCount;

            for (int i = 0; i < dataGridView1.RowCount - 1; ++i)
            {
                dataGridView3.Rows.Add();
                dataGridView3[0, i].Value = array[i];
            }
        }

        public void Shake()
        {
            string[] x;

            int rows = dataGridView1.Rows.Count;

            x = new string[dataGridView1.RowCount - 1];

            for (int i = 0; i < dataGridView1.RowCount - 1; ++i)
            {
                x[i] = dataGridView1[0, i].Value.ToString();

            }

            int[] array = x.Select(i => int.Parse(i)).ToArray();

            for (var i = 0; i < array.Length / 2; i++)
            {
                var swapFlag = false;
                //проход слева направо
                for (var j = i; j < array.Length - i - 1; j++)
                {
                    if (array[j] > array[j + 1])
                    {
                        Swap(ref array[j], ref array[j + 1]);
                        swapFlag = true;
                    }
                }

                //проход справа налево
                for (var j = array.Length - 2 - i; j > i; j--)
                {
                    if (array[j - 1] > array[j])
                    {
                        Swap(ref array[j - 1], ref array[j]);
                        swapFlag = true;
                    }
                }

                //если обменов не было выходим
                if (!swapFlag)
                {
                    break;
                }
            }

            int N = dataGridView1.RowCount;

            for (int i = 0; i < dataGridView1.RowCount - 1; ++i)
            {
                dataGridView4.Rows.Add();
                dataGridView4[0, i].Value = array[i];
            }
        }

        static int Partition(int[] array, int minIndex, int maxIndex)
        {
            var pivot = minIndex - 1;
            for (var i = minIndex; i < maxIndex; i++)
            {
                if (array[i] < array[maxIndex])
                {
                    pivot++;
                    Swap(ref array[pivot], ref array[i]);
                }
            }

            pivot++;
            Swap(ref array[pivot], ref array[maxIndex]);
            return pivot;
        }

        static int[] QuickSort(int[] array, int minIndex, int maxIndex)
        {
            if (minIndex >= maxIndex)
            {
                return array;
            }

            var pivotIndex = Partition(array, minIndex, maxIndex);
            QuickSort(array, minIndex, pivotIndex - 1);
            QuickSort(array, pivotIndex + 1, maxIndex);

            return array;
        }

        public void Fast()
        {
            string[] x;

            int rows = dataGridView1.Rows.Count;

            x = new string[dataGridView1.RowCount - 1];

            for (int i = 0; i < dataGridView1.RowCount - 1; ++i)
            {
                x[i] = dataGridView1[0, i].Value.ToString();

            }

            int[] array = x.Select(i => int.Parse(i)).ToArray();

            QuickSort(array, 0, array.Length - 1);

            int N = dataGridView1.RowCount;

            for (int i = 0; i < dataGridView1.RowCount - 1; ++i)
            {
                dataGridView5.Rows.Add();
                dataGridView5[0, i].Value = array[i];
            }
        }

        //метод для проверки упорядоченности массива
        static bool IsSorted(int[] a)
        {
            for (int i = 0; i < a.Length - 1; i++)
            {
                if (a[i] > a[i + 1])
                    return false;
            }

            return true;
        }

        //перемешивание элементов массива
        static int[] RandomPermutation(int[] a)
        {
            Random random = new Random();
            var n = a.Length;
            while (n > 1)
            {
                n--;
                var i = random.Next(n + 1);
                var temp = a[i];
                a[i] = a[n];
                a[n] = temp;
            }

            return a;
        }

        //случайная сортировка
        static int[] BogoSorting(int[] a)
        {
            while (!IsSorted(a))
            {
                a = RandomPermutation(a);
            }

            return a;
        }

        public void BOGOSORT()
        {
            string[] x;

            int rows = dataGridView1.Rows.Count;

            x = new string[dataGridView1.RowCount - 1];

            for (int i = 0; i < dataGridView1.RowCount - 1; ++i)
            {
                x[i] = dataGridView1[0, i].Value.ToString();

            }

            int[] array = x.Select(i => int.Parse(i)).ToArray();

            BogoSorting(array);

            int N = dataGridView1.RowCount;

            for (int i = 0; i < dataGridView1.RowCount - 1; ++i)
            {
                dataGridView6.Rows.Add();
                dataGridView6[0, i].Value = array[i];
            }
        }

        private void menuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                Bubblesort();
            }
            if (checkBox2.Checked == true)
            {
                Vstavka();
            }
            if (checkBox3.Checked == true)
            {
                Shake();
            }
            if (checkBox4.Checked == true)
            {
                Fast();
            }
            if (checkBox5.Checked == true)
            {
                BOGOSORT();
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void чистимToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            dataGridView3.Rows.Clear();
            dataGridView4.Rows.Clear();
            dataGridView5.Rows.Clear();
            dataGridView6.Rows.Clear();
        }

        private void выбратьExcelДокументToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Excel();
        }
    }
}
