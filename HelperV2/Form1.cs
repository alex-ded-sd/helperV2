using DuplicatesFinder.Core;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace HelperV2
{
	public partial class Form1 : Form
    {
	    private ExcelWorker _worker;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog1.FileName;
                try {
	                if (_worker == null) {
		                _worker = new ExcelWorker(filePath);
	                }
	                List<string> sheetsNames = _worker.GetWorkSheets();
	                listBox1.Items.AddRange(sheetsNames.ToArray());
                }
                catch (NullReferenceException exception) {
	                MessageBox.Show(exception.Message);
                }
                catch (Exception ex) {
	                MessageBox.Show($"Что пошло не так. Ошибка\n{ex.Message}");
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            object item = listBox1.SelectedItem;
            if (item != null)
            {
	            try {
					List<Entity> allRecords = _worker.ReadDataFrom(item.ToString());
					UniqueRecordsFinder finder = new UniqueRecordsFinder();
					List<Entity> uniqueRecords = finder.GetUniqueRecords(allRecords);
					_worker.SaveAndShowAsExcelFile(uniqueRecords);
					listBox1.Items.Clear();
					_worker = null;
	            }
	            catch (NullReferenceException exception) {
		            MessageBox.Show(exception.Message);
	            }
	            catch (Exception exception) {
		            MessageBox.Show($"Проблемка........ отправь хозяину скриншот с ошибкой\n{exception.Message}");
	            }
            }
            else
            {
	            MessageBox.Show("Ты не выбрала книгу ексельки, в которой треки");
            }
        }

		private void button3_Click(object sender, EventArgs e) {
			if (_worker != null) {
				MessageBox.Show("Ты уже загрузила ексельку");
				return;
			}
			if (openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				string filePath = openFileDialog1.FileName;
				try {
					_worker = new ExcelWorker(filePath);
				}
				catch (NullReferenceException exception) {
					MessageBox.Show(exception.Message);
				}
				catch (Exception ex) {
					MessageBox.Show($"Что пошло не так. Ошибка\n{ex.Message}");
				}
			}
		}
	}
}
