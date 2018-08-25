namespace WindowsFormsApp1
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtWordFile = new System.Windows.Forms.TextBox();
            this.openFileDialogWord = new System.Windows.Forms.OpenFileDialog();
            this.btnOpenWordFile = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox1.Location = new System.Drawing.Point(55, 106);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(1170, 448);
            this.textBox1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(50, 41);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(104, 25);
            this.label1.TabIndex = 1;
            this.label1.Text = "Word File";
            // 
            // txtWordFile
            // 
            this.txtWordFile.Location = new System.Drawing.Point(179, 41);
            this.txtWordFile.Name = "txtWordFile";
            this.txtWordFile.Size = new System.Drawing.Size(632, 31);
            this.txtWordFile.TabIndex = 3;
            // 
            // openFileDialogWord
            // 
            this.openFileDialogWord.Filter = "files|*.docx";
            this.openFileDialogWord.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialogWord_FileOk);
            // 
            // btnOpenWordFile
            // 
            this.btnOpenWordFile.Location = new System.Drawing.Point(817, 37);
            this.btnOpenWordFile.Name = "btnOpenWordFile";
            this.btnOpenWordFile.Size = new System.Drawing.Size(134, 49);
            this.btnOpenWordFile.TabIndex = 5;
            this.btnOpenWordFile.Text = "Browse";
            this.btnOpenWordFile.UseVisualStyleBackColor = true;
            this.btnOpenWordFile.Click += new System.EventHandler(this.btnOpenWordFile_Click);
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(1007, 37);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(208, 41);
            this.btnExport.TabIndex = 7;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.Filter = "file|*.xlsx";
            this.saveFileDialog1.OverwritePrompt = false;
            this.saveFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog1_FileOk);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1270, 620);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.btnOpenWordFile);
            this.Controls.Add(this.txtWordFile);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox1);
            this.Name = "Form1";
            this.Text = "Export Excel";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtWordFile;
        private System.Windows.Forms.OpenFileDialog openFileDialogWord;
        private System.Windows.Forms.Button btnOpenWordFile;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
    }
}

