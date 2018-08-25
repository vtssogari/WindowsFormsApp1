using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    class ExcelAutomation
    {
        TextBox textBox1 = null;
        Int16[] columnSizes = new Int16[20];
        List<string> columnList = new List<string>();

        public ExcelAutomation(TextBox textBox)
        {
            this.textBox1 = textBox;
            this.columnList.Add("A");
            this.columnList.Add("B");

            this.columnList.Add("C");
            this.columnList.Add("D");
            this.columnList.Add("E");
            this.columnList.Add("F");
            this.columnList.Add("G");

            this.columnList.Add("H");
            this.columnList.Add("I");
            this.columnList.Add("J");
            this.columnList.Add("K");
            this.columnList.Add("L");

            this.columnList.Add("M");
            this.columnList.Add("N");
            this.columnList.Add("O");
            this.columnList.Add("P");
            this.columnList.Add("Q");

            this.columnList.Add("R");
            this.columnList.Add("S");
            this.columnList.Add("T");
        }

        public bool export(string excelFile, Int16[] sizes)
        {
            if(columnSizes != null)
            {
                this.columnSizes = sizes;
            }
            else
            {
                for(int i =0; i < 20; i++)
                {
                    this.columnSizes[i] = 30;
                }
            }
            Microsoft.Office.Interop.Excel._Application objExcelApp = null;
            try
            {
                objExcelApp = new Microsoft.Office.Interop.Excel.Application();
                objExcelApp.Visible = false;
                Microsoft.Office.Interop.Excel._Workbook workbook = objExcelApp.Workbooks.Add(1);
                Microsoft.Office.Interop.Excel._Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
                if (worksheet == null)
                {
                    textBox1.AppendText("Worksheet could not be created. Check that your office installation and project references are correct.");
                    objExcelApp.Quit();
                    return false;
                }

                worksheet.Range["A1"].Activate();
                // resize the column
                for(int i = 0; i < this.columnList.Count; i++)
                {
                    worksheet.Range[this.columnList[i]+"1"].ColumnWidth = this.columnSizes[i];
                }
                
                worksheet.Paste();
                // Save the excel file
                workbook.SaveAs(excelFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault);
                objExcelApp.Workbooks.Close();
                objExcelApp.Quit();
            }
            catch (Exception ex)
            {
                textBox1.AppendText(ex.Message);
                if(objExcelApp != null)
                {
                    objExcelApp.Quit();
                }
                return false;
            }
            return true;
        }
    }
}
