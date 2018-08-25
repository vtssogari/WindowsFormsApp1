using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        Int16[] sizes = new Int16[20];

        private void Form1_Load(object sender, EventArgs e)
        {
            for (int i = 0; i < sizes.Length; i++)
            {
                sizes[i] = 30;
            }
            try
            {
                string[] args = Environment.GetCommandLineArgs();
                string appPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                if (args.Length > 2)
                {
                    string wordFile = args[1];
                    string excelFile = args[2];
                    if(args.Length > 3)
                    {
                        string columnSizeList = args[3];
                        string[] list = columnSizeList.Split(',');
                        for(int i = 0; i < sizes.Length; i++)
                        {
                            if(i < list.Length)
                            {
                                sizes[i] = Int16.Parse( list[i]);
                            }
                            else
                            {
                                sizes[i] = 30;
                            }
                        }
                    }
                    export(wordFile, excelFile, sizes);
                }
            }
            catch (Exception ex)
            {
                textBox1.AppendText("Error: " + ex.Message);
            }
        }

        private void export(string wordFile, string excelFile, Int16[] sizes)
        {
            if (File.Exists(excelFile))
            {
                File.Delete(excelFile);
            }
            WordAutomation word = new WordAutomation(textBox1);
            bool isWordOpened = word.open(wordFile);
            if (isWordOpened)
            {
                ExcelAutomation excel = new ExcelAutomation(textBox1);
                bool isExported = excel.export(excelFile, sizes);
                word.close();
                textBox1.AppendText("Word document table contents exported to excel file:" + excelFile);
                this.Close();
            }
            else
            {
                word.close();
            }
        }

        private void btnOpenWordFile_Click(object sender, EventArgs e)
        {
            this.openFileDialogWord.ShowDialog();
        }

        private void openFileDialogWord_FileOk(object sender, CancelEventArgs e)
        {
            this.txtWordFile.Text = this.openFileDialogWord.FileName;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            this.saveFileDialog1.ShowDialog();
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            if (!txtWordFile.Text.Equals(""))
            {
                try
                {
                    export(txtWordFile.Text, this.saveFileDialog1.FileName, sizes);
                }catch(Exception ex)
                {
                    this.textBox1.AppendText(ex.Message);
                }
            }
        }
    }
}
