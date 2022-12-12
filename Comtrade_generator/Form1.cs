using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Comtrade_generator
{
    public partial class Form1 : Form
    {
        double[] amps;
        double[] phases;
        double[,] data;
        int[,] dataInt;
        double[] As;
        double[] Bs;
        string[] names;
        string[] unites;
        ListBox[] listbox = new ListBox[21];
        NumericUpDown[] nuds = new NumericUpDown[21];
        public Form1()
        {
            InitializeComponent();
            cbxFreq.SelectedIndex = 2;
            startForm();
        }

        private void startForm()
        {
            string[] labelText = { "Ia(RMS)", "Ia(ph)", "Ib(RMS)", "Ib(ph)", "Ic(RMS)", "Ic(ph)", "In(RMS)", "In(ph)",
                "Ua(RMS)", "Ua(ph)", "Ub(RMS)", "Ub(ph)", "Uc(RMS)", "Uc(ph)", "Ur(RMS)", "Ur(ph)", "Us(RMS)", "Us(ph)", "Ut(RMS)", "Ut(ph)", "duration" };

            for (int i = 0; i < listbox.Length; i++)
            {
                listbox[i] = new ListBox(); ;
                listbox[i].Height = 100;
                listbox[i].Width = 50;
                listbox[i].Location = new Point(20 + i * listbox[i].Width, 300);
                Label label = new Label();
                label.Width = listbox[i].Width;
                label.TextAlign = ContentAlignment.MiddleCenter;
                label.Text = labelText[i];
                label.Location = new Point(20 + i * listbox[i].Width, 270);
                this.Controls.Add(label);
                this.Controls.Add(listbox[i]);
            }
            nuds[0] = nudIaAmp; nuds[1] = nudIaAngle;
            nuds[2] = nudIbAmp; nuds[3] = nudIbAngle;
            nuds[4] = nudIcAmp; nuds[5] = nudIcAngle;
            nuds[6] = nudInAmp; nuds[7] = nudInAngle;
            nuds[8] = nudUaAmp; nuds[9] = nudUaAngle;
            nuds[10] = nudUbAmp; nuds[11] = nudUbAngle;
            nuds[12] = nudUcAmp; nuds[13] = nudUcAngle;
            nuds[14] = nudUrAmp; nuds[15] = nudUrAngle;
            nuds[16] = nudUsAmp; nuds[17] = nudUsAngle;
            nuds[18] = nudUtAmp; nuds[19] = nudUtAngle;
            nuds[20] = comtradeLength;
        }

        private void setSimmetric()
        {
            for (int i = 0; i < nuds.Length; i++)
            {
                nuds[i].Enabled = false;
            }
            nuds[0].Enabled = true; nuds[1].Enabled = true;
            nuds[8].Enabled = true; nuds[9].Enabled = true;
            nuds[14].Enabled = true; nuds[15].Enabled = true;
            nuds[20].Enabled = true;
        }

        private void setAsimmertic()
        {
            for (int i = 0; i < nuds.Length; i++)
                nuds[i].Enabled = true;
        }


        private void addInterval()
        {
            if (!cbSimmetric.Checked)
            {
                for (int i = 0; i < listbox.Length; i++)
                    listbox[i].Items.Add(nuds[i].Value);
            }
            else
            {
                for(int i = 0; i < 3; i++)
                {
                    listbox[2 * i].Items.Add(nuds[0].Value);
                    listbox[8 + 2 * i].Items.Add(nuds[8].Value);
                    listbox[14 + 2 * i].Items.Add(nuds[14].Value);
                    listbox[1 + 2 * i].Items.Add((nuds[1].Value + 240 * i) % 360);
                    listbox[9 + 2 * i].Items.Add((nuds[9].Value + 240 * i) % 360);
                    listbox[15 + 2 * i].Items.Add((nuds[15].Value + 240 * i) % 360);
                }
                listbox[6].Items.Add(0);
                listbox[7].Items.Add(0);
                listbox[20].Items.Add(comtradeLength.Value);
            }
            
        }

        private void setAmps(int index)
        {
            amps = new double[10]
            {
                Convert.ToDouble(listbox[0].Items[index]) * Math.Sqrt(2),
                Convert.ToDouble(listbox[2].Items[index]) * Math.Sqrt(2),
                Convert.ToDouble(listbox[4].Items[index]) * Math.Sqrt(2),
                Convert.ToDouble(listbox[6].Items[index]) * Math.Sqrt(2),
                Convert.ToDouble(listbox[8].Items[index]) * Math.Sqrt(2),
                Convert.ToDouble(listbox[10].Items[index]) * Math.Sqrt(2),
                Convert.ToDouble(listbox[12].Items[index]) * Math.Sqrt(2),
                Convert.ToDouble(listbox[14].Items[index]) * Math.Sqrt(2),
                Convert.ToDouble(listbox[16].Items[index]) * Math.Sqrt(2),
                Convert.ToDouble(listbox[18].Items[index]) * Math.Sqrt(2)
            };
        }

        private void setPhases(int index)
        {
            phases = new double[10]
            {
                Convert.ToDouble(listbox[1].Items[index]),
                Convert.ToDouble(listbox[3].Items[index]),
                Convert.ToDouble(listbox[5].Items[index]),
                Convert.ToDouble(listbox[7].Items[index]),
                Convert.ToDouble(listbox[9].Items[index]),
                Convert.ToDouble(listbox[11].Items[index]),
                Convert.ToDouble(listbox[13].Items[index]),
                Convert.ToDouble(listbox[15].Items[index]),
                Convert.ToDouble(listbox[17].Items[index]),
                Convert.ToDouble(listbox[19].Items[index])
            };
        }

        private void setStartValues()
        {
            names = new string[10] { "Ia", "Ib", "Ic", "I0", "Ua", "Ub", "Uc", "Ur", "Us", "Ut" };

            unites = new string[10] { "A", "A", "A", "A", "V", "V", "V", "V", "V", "V" };
            int frequency = 50;
            int an_max = 65535;
            int an_min = 0;
            int samples = 0;
            foreach (object listBoxItem in listbox[20].Items)
                samples += Convert.ToInt32(Convert.ToDouble(listBoxItem.ToString()) * Convert.ToInt32(cbxFreq.Text));
            Console.WriteLine(samples);
            data = new double[10, samples];
            dataInt = new int[10, samples];
            As = new double[10];
            Bs = new double[10];
        }

        private void generateDataInt()
        {
            for (int i = 0; i < 10; i++)
            {
                As[i] = 2 * amps[i] / 65535 + 1e-8;
                Bs[i] = - amps[i];
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    if (amps[i] == 0)
                        dataInt[i, j] = 0;
                    else
                        dataInt[i,j] = (int)((data[i,j] - Bs[i])/As[i]);
                }
            }
        }
        private void writeToFile(string fileName)
        {
            using (StreamWriter writetext = new StreamWriter(fileName))
            {
                writetext.WriteLine("TEL,0,1999");
                writetext.WriteLine("10,10A,0D");
                for (int i = 0; i < 10; i++)
                {
                    writetext.WriteLine("" + (i+1) + "," + names[i] + "," + ",," + unites[i] + "," +
                        (As[i]).ToString("E", CultureInfo.GetCultureInfo("en-GB")) + "," + 
                        (Bs[i]).ToString("E", CultureInfo.GetCultureInfo("en-GB")) + ",0,0,65535,5.0,5.0,P");
                }
                writetext.WriteLine("50");
                writetext.WriteLine("1");
                writetext.WriteLine("" + Convert.ToInt32(cbxFreq.Text) + "," +
                    Convert.ToInt32(cbxFreq.Text) * Convert.ToDouble(comtradeLength.Value));
                writetext.WriteLine("19/03/2019,10:22:30.89000");
                writetext.WriteLine("19/03/2019,10:22:31.89000");
                writetext.WriteLine("ASCII");
                writetext.WriteLine("1");
            }

            using (StreamWriter writetext = new StreamWriter(Path.GetDirectoryName(fileName) + "/" + Path.GetFileNameWithoutExtension(fileName) + ".dat"))
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    writetext.Write("" + (j + 1) + "," + (1000000 / Convert.ToInt32(cbxFreq.Text) * (j + 1)));
                    for (int i = 0; i < 10; i++)
                        writetext.Write("," + dataInt[i, j]);
                    writetext.WriteLine("");
                }
            }
        }

        private void generateData()
        {
            int freq = Convert.ToInt32(cbxFreq.Text);
            int counter = 0;
            int sampleCounter = 0;
            foreach (var listBoxItem in listbox[0].Items)
            {
                setAmps(counter);
                setPhases(counter);
                Console.WriteLine("iteration");
                for (int i = 0; i < 10; i++)
                {
                    for (int j = 0; j < freq * Convert.ToDouble(listbox[20].Items[counter]); j++)
                    {
                        data[i, sampleCounter + j] = amps[i] * Math.Sin(2 * Math.PI * j * 50 / freq + 2 * Math.PI * phases[i] / 360);
                    }
                }
                sampleCounter += freq * Convert.ToInt32(listbox[20].Items[counter]);
                counter++;
            }
        }

        private void writeComtrade()
        {
            setStartValues();

            generateData();
            generateDataInt();
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Comtrade|*.cfg";
            sfd.RestoreDirectory = true;
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                writeToFile(sfd.FileName);
                System.Diagnostics.Process.Start(sfd.FileName);
            }
        }

        private void cbSimmetric_CheckedChanged(object sender, EventArgs e)
        {
            if (cbSimmetric.Checked)
                setSimmetric();
            else
                setAsimmertic();
        }
        private void btnWriteComtrade_Click(object sender, EventArgs e)
        {
            if (listbox[0].Items.Count == 0)
                MessageBox.Show("Empty data", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
                writeComtrade();
        }

        private void btnAddInterval_Click(object sender, EventArgs e)
        {
            addInterval();
        }

        private void btnClearAll_Click(object sender, EventArgs e)
        {
            foreach (ListBox lb in listbox)
                lb.Items.Clear();
        }
    }
}
