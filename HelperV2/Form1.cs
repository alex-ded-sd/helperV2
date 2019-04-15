using DuplicatesFinder.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HelperV2
{
    using System.IO;

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog1.FileName;
                ExcelWorker worker = new ExcelWorker();
                List<string> sheetsNames = worker.GetWorkSheets(filePath);
                listBox1.Items.AddRange(sheetsNames.ToArray());
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            object item = listBox1.SelectedItem;
            if (item == null)
            {
                MessageBox.Show("Ты же забыла файл прочитать");
            }
            else
            {
             
                if (string.IsNullOrEmpty(openFileDialog1.FileName))
                {
                    if(openFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            string filePath = openFileDialog1.FileName;
                            ExcelWorker worker = new ExcelWorker();
                            List<Entity> allRecords = worker.ReadDataFrom(filePath, item.ToString());
                            UniqueRecordsFinder finder = new UniqueRecordsFinder();
                            List<Entity> uniqueRecords = finder.GetUniqueRecords(allRecords);
                            string folder = Path.GetDirectoryName(filePath);
                            worker.WriteData(Path.Combine(folder, "обработан.xlsx"), uniqueRecords);
                            MessageBox.Show("Сохранено в файл обработан.xlsx возле прогаммы");
                        }
                        catch (Exception exception)
                        {
                            MessageBox.Show(exception.Message);
                        }
                    }
                }
                else
                {
                    try
                    {
                        string filePath = openFileDialog1.FileName;
                        ExcelWorker worker = new ExcelWorker();
                        List<Entity> allRecords = worker.ReadDataFrom(filePath, item.ToString());
                        UniqueRecordsFinder finder = new UniqueRecordsFinder();
                        List<Entity> uniqueRecords = finder.GetUniqueRecords(allRecords);
                        string folder = Path.GetDirectoryName(filePath);
                        worker.WriteData(Path.Combine(folder, "обработан.xlsx"), uniqueRecords);
                        MessageBox.Show("Сохранено в файл обработан.xlsx возле исходного файла");
                    }
                    catch(Exception exception)
                    {
                        MessageBox.Show(exception.Message);
                    }
                }

            
            }
        }
    }
}
