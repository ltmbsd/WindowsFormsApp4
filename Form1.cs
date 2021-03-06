using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using LiveCharts;
using LiveCharts.Wpf;
using System.Drawing;
using Microsoft.Office.Interop.Word;

namespace WindowsFormsApp4
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog qq = new OpenFileDialog();
            qq.Filter = "csv|*.csv";
            if (qq.ShowDialog() == DialogResult.OK)
            {
                using (StreamReader reader = new StreamReader(qq.FileName, System.Text.Encoding.Default))
                {
                    textBox4.Text = "";
                    dataGridView1.Rows.Clear();
                    dataGridView1.Columns.Clear();
                    int max = 0, count = 0;
                    double summ = 0;
                    string[] line1 = File.ReadAllLines(qq.FileName);
                    for (int i = 0; i < line1.Length; i++)
                    {
                        string[] line = line1[i].Split(';');
                        if (line.Length > max)
                            max = line.Length;
                    }
                    for (int i = 0; i < max - 1; i++)
                        dataGridView1.Columns.Add("", "");
                    for (int i = 1; i < line1.Length; i++)
                    {
                        dataGridView1.Rows.Add();
                    }

                    for (int i = 0; i < line1.Length; i++)
                    {
                        string[] line = line1[i].Split(';');
                        for (int j = 0; j < line.Length; j++)
                        {
                            try
                            {
                                dataGridView1.Rows[i].Cells[j].Value = Convert.ToDouble(line[j].Replace(".", ","));
                                summ += Convert.ToDouble(dataGridView1.Rows[i].Cells[j].Value);
                                count++;
                            }
                            catch { }
                        }

                    }
                    if (count > 0)
                    {
                        Genm(count, summ);
                        TableToArray();
                    }
                }
            }
        }

        private void Genm(int count, double summ)
        {
            textBox1.Text = Convert.ToString(Math.Round((summ / count), 3));
        }

        private void TableToArray()
        {
            int count = 0, k = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                        count++;
                }
            double[] nums = new double[count];
            int[] repeats = new int[count];
            for (int i = 0; i < count; i++)
                repeats[i]++;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                    {
                        nums[k] = Convert.ToDouble(dataGridView1.Rows[i].Cells[j].Value);
                        k++;
                    }
                }
            DropRepeats(nums, repeats);
        }

        public void DropRepeats(double[] nums, int[] repeats)
        {
            int count = nums.Count();
            for (int i = 0; i < repeats.Count(); i++)
                repeats[i]--;
            for (int i = 0; i < count; i++)
            {
                for (int j = i + 1; j < count; j++)
                {
                    if (nums[i] == nums[j])
                    {
                        for (int k = j; k < count - 1; k++)
                        {
                            nums[k] = nums[k + 1];
                            repeats[k] = repeats[k + 1];
                        }
                        count--;
                    }
                }
            }
            double[] numsfixed = new double[count];
            int[] repeatsfixed = new int[count];
            for (int i = 0; i < count; i++)
            {
                numsfixed[i] = nums[i];
                repeatsfixed[i] = repeats[i];
            }
            Maxmin(numsfixed);
            IntervalsCount(nums.Length, numsfixed);
        }

        private void Maxmin(double[] numsfixed)
        {
            Array.Sort(numsfixed);
            textBox5.Text = Convert.ToString(numsfixed[numsfixed.Count() - 1]);
            textBox6.Text = Convert.ToString(numsfixed[0]);
        }

        private void IntervalsCount(int count, double[] numsfixed)
        {
            double inter;
            inter = 1 + 3.322 * Math.Log10(count);
            if ((inter % 1) == 0)
                textBox7.Text = Convert.ToString(Math.Round(inter, 0));
            else textBox7.Text = Convert.ToString(Math.Round(inter, 0) + 1);
            Interlenght(numsfixed);
        }

        private void Interlenght(double[] numsfixed)
        {
            textBox8.Text = Convert.ToString(Math.Round((Convert.ToDouble(textBox5.Text) - Convert.ToDouble(textBox6.Text)) / Convert.ToDouble(textBox7.Text), 2));
            Intervals();
        }

        private void Intervals()
        {
            int k = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                    {
                        k++;
                    }
                }
            }
            double[] allnumssorted = new double[k];
            k = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                    {
                        allnumssorted[k] = Convert.ToDouble(dataGridView1.Rows[i].Cells[j].Value);
                        k++;
                    }
                }
            }
            Array.Sort(allnumssorted);
            richTextBox1.Text = "";
            double min = Convert.ToDouble(textBox6.Text);
            double h = Convert.ToDouble(textBox8.Text);
            string[] intervals = new string[Convert.ToInt32(textBox7.Text)];
            int[] frequency = new int[Convert.ToInt32(textBox7.Text)];
            double[] intmin = new double[Convert.ToInt32(textBox7.Text)];
            int count;
            for (int i = 0; i < Convert.ToInt32(textBox7.Text); i++)
            {
                intmin[i] = min;
                count = 0;
                richTextBox1.Text += (i + 1) + " интервал(От " + min + " до " + Math.Round(min + h, 2) + "): ";
                intervals[i] = "От " + min + " до " + Math.Round(min + h, 2);
                for (int j = 0; j < allnumssorted.Count(); j++)
                {
                    if (i == 0)
                    {
                        if ((allnumssorted[j] >= min) && (allnumssorted[j] < (min + h)))
                        {
                            richTextBox1.Text += allnumssorted[j] + "; ";
                            count++;
                        }
                    }
                    else
                    {
                        if ((allnumssorted[j] > min) && (allnumssorted[j] <= (min + h)))
                        {
                            richTextBox1.Text += allnumssorted[j] + "; ";
                            count++;
                        }
                        else if ((i == Convert.ToInt32(textBox7.Text) - 1) && (j == allnumssorted.Length - 1))
                        {
                            richTextBox1.Text += allnumssorted[j] + "; ";
                            count++;
                        }
                    }
                }
                min = Math.Round(min + h, 2);
                richTextBox1.Text += "\nЧастота: " + count +
                                     "\nЧастость: " + Math.Round((double)count / allnumssorted.Length * 100, 2) +
                                     "%\n\n";
                frequency[i] = count;
            }
            Moda(frequency, intmin, h);
            IntervalsGraphic(intervals, frequency, allnumssorted);
        }

        private void Moda(int[] frequency, double[] intmin, double h)
        {
            int i = 0;
            while (frequency[i] != frequency.Max())
                i++;
            if ((i > 0) && (i < frequency.Length - 1))
            {
                double result = Math.Round(intmin[i] + ((frequency[i] - frequency[i - 1]) / (frequency[i] - frequency[i - 1] + (frequency[i] - frequency[i + 1])) * h), 2);
                textBox2.Text = Convert.ToString(result);
            }
            else if (i == 0)
            {
                double result = Math.Round(intmin[i] + ((frequency[i]) / (frequency[i] + (frequency[i] - frequency[i + 1])) * h), 2);
                textBox2.Text = Convert.ToString(result);
            }
            else
            {
                double result = Math.Round(intmin[i] + ((frequency[i] - frequency[i - 1]) / (frequency[i] - frequency[i - 1] + (frequency[i])) * h), 2);
                textBox2.Text = Convert.ToString(result);
            }
            Mediana(frequency, intmin, h, i);
            MidleLineDeviation(frequency, intmin, h);
        }

        private void Mediana(int[] frequency, double[] intmin, double h, int i)
        {
            if (i != 0)
            {
                double result = intmin[i] + (0.5 * frequency.Sum() - frequency[i - 1]) / frequency[i] * h;
                textBox3.Text = Convert.ToString(Math.Round(result, 2));
            }
            else
            {
                double result = intmin[i] + (0.5 * frequency.Sum()) / frequency[i] * h;
                textBox3.Text = Convert.ToString(Math.Round(result, 2));
            }
        }

        private void MidleLineDeviation(int[] frequency, double[] intmin, double h)
        {
            double[] numerator = new double[frequency.Length];
            double sum = 0;
            for (int i = 0; i < numerator.Length; i++)
                numerator[i] = Math.Abs((2 * intmin[i] + h) / 2 - Convert.ToDouble(textBox1.Text));
            for (int i = 0; i < numerator.Length; i++)
            {
                sum += numerator[i] * frequency[i];
            }
            textBox9.Text = Convert.ToString(Math.Round(sum / numerator.Length, 2));
            Range();
            Dispersion(frequency, intmin, h);
        }

        private void IntervalsGraphic(string[] intervals, int[] frequency, double[] allnamssorted)
        {
            LiveCharts.SeriesCollection series = new LiveCharts.SeriesCollection();
            ChartValues<int> frq = new ChartValues<int>();
            List<string> intrvls = new List<string>();
            for (int i = 0; i < intervals.Count(); i++)
            {
                intrvls.Add(intervals[i]);
                frq.Add(frequency[i]);
            }
            cartesianChart2.AxisX.Clear();
            cartesianChart2.AxisY.Clear();
            cartesianChart2.AxisX.Add(new LiveCharts.Wpf.Axis()
            {
                Title = "Интервалы",
                Labels = intrvls
            });

            ColumnSeries values = new ColumnSeries();
            values.Title = "Частота";
            values.Values = frq;
            series.Add(values);

            cartesianChart2.Series = series;

            LiveCharts.SeriesCollection series1 = new LiveCharts.SeriesCollection();
            cartesianChart1.AxisX.Clear();
            cartesianChart1.AxisY.Clear();
            cartesianChart1.AxisX.Add(new LiveCharts.Wpf.Axis()
            {
                Title = "Интервалы",
                Labels = intrvls
            });
            LineSeries values1 = new LineSeries();
            values1.Title = "Частота";
            values1.Values = frq;
            series1.Add(values1);

            cartesianChart1.Series = series1;
        }

        private void Range()
        {
            textBox10.Text = Convert.ToString(Convert.ToDouble(textBox5.Text) - Convert.ToDouble(textBox6.Text));
        }
        private void Dispersion(int[] frequency, double[] intmin, double h)
        {
            double sum = 0;
            for (int i = 0; i < frequency.Length; i++)
            {
                sum += Math.Pow((2 * intmin[i] + h) / 2 - Convert.ToDouble(textBox1.Text), 2) * frequency[i];
            }
            textBox11.Text = Convert.ToString(Math.Round(sum / frequency.Length, 2));
            MidleSquereDiviation();
            CoefAssim(frequency, intmin, h);
            Ekscess(frequency, intmin, h);
        }
        private void MidleSquereDiviation()
        {
            textBox12.Text = Convert.ToString(Math.Round(Math.Pow(Convert.ToDouble(textBox11.Text), 0.5), 2));
            CoefficientVariation();
        }

        private void CoefficientVariation()
        {
            textBox13.Text = Convert.ToString(Math.Round(Convert.ToDouble(textBox12.Text) / Convert.ToDouble(textBox1.Text), 2));
        }
        private void CoefAssim(int[] frequency, double[] intmin, double h)
        {
            double sum = 0;
            for (int i = 0; i < frequency.Length; i++)
            {
                sum += Math.Pow((2 * intmin[i] + h / 2) - Convert.ToDouble(textBox1.Text), 3) * frequency[i] / frequency.Sum();
            }
            sum = sum / Math.Pow(Convert.ToDouble(textBox12.Text), 3);
            textBox14.Text = Convert.ToString(Math.Round(sum, 2));
            AssimText();
            SushAssim(frequency.Length);
        }


        private void AssimText()
        {
            label19.Visible = true;
            if (Convert.ToDouble(textBox14.Text) > 0)
                label19.Text = "Асимметрия правосторонняя";
            else if (Convert.ToDouble(textBox14.Text) < 0)
                label19.Text = "Асимметрия левосторонняя";
            else
                label19.Text = "График симметричен";
        }
        private void SushAssim(double count)
        {
            textBox15.Text = Convert.ToString(Math.Round(Math.Abs(Convert.ToDouble(textBox14.Text)) / Math.Sqrt(6 * (count - 1) / ((count + 1) * (count + 3))), 2));
            SushAssimText();
        }

        private void SushAssimText()
        {
            label20.Visible = true;
            if (Convert.ToDouble(textBox15.Text) > 3)
                label20.Text = "Асимметрия существенна";
            else
                label20.Text = "Асимметрия несущественна";
        }

        private void Ekscess(int[] frequency, double[] intmin, double h)
        {
            int k = 0, l = 0, j = 0;
            double sum = 0;
            for (int i = 0; i < frequency.Length; i++)
            {
                if (k >= l)
                {
                    l += frequency[j];
                    j++;
                }
                sum += Math.Pow((2 * intmin[i] + h) / 2 - Convert.ToDouble(textBox1.Text), 4) * frequency[i] / frequency.Sum();
                k++;
            }
            sum = sum / Math.Pow(Convert.ToDouble(textBox12.Text), 4) - 3;
            textBox17.Text = Convert.ToString(Math.Round(sum, 2));
            EkscessText();
            SushEkscess(frequency.Length);
        }

        private void EkscessText()
        {
            label21.Visible = true;
            if (Convert.ToDouble(textBox17.Text) > 0)
                label21.Text = "Распределение более островершинное";
            else if (Convert.ToDouble(textBox17.Text) < 0)
                label21.Text = "Распределение более плосковершинное";
            else label21.Text = "Распределение нормальное";
        }

        private void SushEkscess(double count)
        {
            double o;
            o = Math.Sqrt((24 * count * (count - 2) * (count - 3)) / (Math.Pow(count - 1, 2) * (count + 3) * (count + 5)));
            textBox16.Text = Convert.ToString(Math.Round(Math.Abs(Convert.ToDouble(textBox17.Text)) / o, 2));
            SushEkscessText();
        }

        private void SushEkscessText()
        {
            label22.Visible = true;
            if (Convert.ToDouble(textBox16.Text) > 3)
                label22.Text = "Эксцесс существенный";
            else label22.Text = "Эксцесс несущественный";
        }


        private void button2_Click(object sender, EventArgs e)
        {
            int k = 0;
            double sum = 0;
            DataGridViewSelectedCellCollection qq = dataGridView1.SelectedCells;
            double[] nums = new double[qq.Count];
            for (int i = 0; i < qq.Count; i++)
            {
                if (qq[i].Value != null)
                {
                    sum += Convert.ToDouble(qq[i].Value);
                    k++;
                }
            }
            if (!(double.IsNaN(sum / k)))
                textBox4.Text = Convert.ToString(Math.Round((sum / k), 3));
            else textBox4.Text = "";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog qq = new SaveFileDialog();
            qq.Filter = "docx|*.docx";
            if (qq.ShowDialog() == DialogResult.OK)
            {
                Bitmap bmp1 = new Bitmap(cartesianChart2.Width, cartesianChart2.Height);
                cartesianChart2.DrawToBitmap(bmp1, new System.Drawing.Rectangle(0, 0, bmp1.Width, bmp1.Height));
                Bitmap bmp2 = new Bitmap(cartesianChart1.Width, cartesianChart1.Height);
                cartesianChart1.DrawToBitmap(bmp2, new System.Drawing.Rectangle(0, 0, bmp2.Width, bmp2.Height));

                Microsoft.Office.Interop.Word._Application oWord = new Microsoft.Office.Interop.Word.Application();

                var oDoc = oWord.Documents.Add();

                //Insert a paragraph at the beginning of the document.
                oDoc.Content.Paragraphs.Add();
                oDoc.Content.Paragraphs.Add();
                oDoc.Content.Paragraphs.Add();
                oDoc.Content.Paragraphs.Add();

                var cb = System.Windows.Forms.Clipboard.GetDataObject();

                System.Windows.Forms.Clipboard.SetImage(bmp1);
                oDoc.Paragraphs[1].Range.Paste();

                System.Windows.Forms.Clipboard.SetImage(bmp2);
                oDoc.Paragraphs[2].Range.Paste();

                System.Windows.Forms.Clipboard.SetDataObject(cb);

                oDoc.Paragraphs[3].Range.Text = "Генеральная средняя: " + textBox1.Text + "\n" +
                                                "Мода: " + textBox2.Text + "\n" +
                                                "Медиана: " + textBox3.Text + "\n" +
                                                "Максимальное: " + textBox5.Text + "\n" +
                                                "Минимальное: " + textBox6.Text + "\n" +
                                                "Количество интервалов: " + textBox7.Text + "\n" +
                                                "Длина интервала: " + textBox8.Text + "\n" +
                                                "Размах: " + textBox10.Text + "\n" +
                                                "Среднее линейное отклонение: " + textBox9.Text + "\n" +
                                                "Дисперсия: " + textBox11.Text + "\n" +
                                                "Среднее квадратическое отклонение: " + textBox12.Text + "\n" +
                                                "Коэффицент вариаций: " + textBox13.Text + "\n" +
                                                "Коэффицент асимметрии: " + textBox14.Text + "\n" +
                                                "Существенность асимметрии: " + textBox15.Text + "\n" +
                                                "Эксцесс: " + textBox17.Text + "\n" +
                                                "Существенность эксцесса: " + textBox16.Text + "\n\n" +
                                                label19.Text + "\n" +
                                                label20.Text + "\n" +
                                                label21.Text + "\n" +
                                                label22.Text + "\n";

                oDoc.SaveAs2(qq.FileName);

                oWord.Quit();
                MessageBox.Show("Данные успешно сохранены");
            }
        }

        private void построениеВариационногоРядаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            построениеВариационногоРядаToolStripMenuItem.Checked = true;
            рассчетПоказателейЦентраРаспределенияИСтруктурныхХарактеристикВариационногоРядаToolStripMenuItem.Checked = false;
            расчетПоказателейВариацииToolStripMenuItem.Checked = false;
            анализНаСимметриюИЭксцессToolStripMenuItem.Checked = false;
            panel1.Visible = true;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
        }

        private void рассчетПоказателейЦентраРаспределенияИСтруктурныхХарактеристикВариационногоРядаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            построениеВариационногоРядаToolStripMenuItem.Checked = false;
            рассчетПоказателейЦентраРаспределенияИСтруктурныхХарактеристикВариационногоРядаToolStripMenuItem.Checked = true;
            расчетПоказателейВариацииToolStripMenuItem.Checked = false;
            анализНаСимметриюИЭксцессToolStripMenuItem.Checked = false;
            panel1.Visible = false;
            panel2.Visible = true;
            panel3.Visible = false;
            panel4.Visible = false;
        }

        private void расчетПоказателейВариацииToolStripMenuItem_Click(object sender, EventArgs e)
        {
            построениеВариационногоРядаToolStripMenuItem.Checked = false;
            рассчетПоказателейЦентраРаспределенияИСтруктурныхХарактеристикВариационногоРядаToolStripMenuItem.Checked = false;
            расчетПоказателейВариацииToolStripMenuItem.Checked = true;
            анализНаСимметриюИЭксцессToolStripMenuItem.Checked = false;
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = true;
            panel4.Visible = false;
        }

        private void анализНаСимметриюИЭксцессToolStripMenuItem_Click(object sender, EventArgs e)
        {
            построениеВариационногоРядаToolStripMenuItem.Checked = false;
            рассчетПоказателейЦентраРаспределенияИСтруктурныхХарактеристикВариационногоРядаToolStripMenuItem.Checked = false;
            расчетПоказателейВариацииToolStripMenuItem.Checked = false;
            анализНаСимметриюИЭксцессToolStripMenuItem.Checked = true;
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = true;
        }
    }
}
