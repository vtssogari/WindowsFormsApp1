using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    class WordAutomation
    {
        TextBox textBox1 = null;
        Microsoft.Office.Interop.Word._Application objWordApp;
        public WordAutomation(TextBox textBox)
        {
            this.textBox1 = textBox;
        }
        
        public bool open(string wordFile)
        {
            this.objWordApp = new Microsoft.Office.Interop.Word.Application();
            if (this.objWordApp == null)
            {
                textBox1.AppendText("Error: Word could not be started. Check that your office installation and project references are correct.");
                return false;
            }
            this.objWordApp.Visible = false;
            try
            {
                Microsoft.Office.Interop.Word.Document objDoc = this.objWordApp.Documents.Open(wordFile);
                if (objDoc.Tables.Count == 0)
                {
                    textBox1.AppendText("Error: This document contains no tables");
                    return false;
                }
                Microsoft.Office.Interop.Word.Table tbl = objDoc.Tables[1];
                tbl.Range.Copy();
            }
            catch (Exception ex)
            {
                textBox1.AppendText("Error: " + ex.Message);
                return false;
            }
            return true;
        }
        
        public void close()
        {
            this.objWordApp.Documents.Close();
            this.objWordApp.Application.Quit();
        }
    }
}
