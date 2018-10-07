namespace TestPDF
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
            progressBar1 = new System.Windows.Forms.ProgressBar();
            percentlabel = new System.Windows.Forms.Label();


            this.button1 = new System.Windows.Forms.Button();
            this.pdfFolderPath = new System.Windows.Forms.RichTextBox();
            this.Generate = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.excelPath = new System.Windows.Forms.RichTextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.excel = new System.Windows.Forms.Button();
            
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(657, 19);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(36, 30);
            this.button1.TabIndex = 0;
            this.button1.Text = "...";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // pdfFolderPath
            // 
            this.pdfFolderPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.pdfFolderPath.Location = new System.Drawing.Point(141, 21);
            this.pdfFolderPath.Name = "pdfFolderPath";
            this.pdfFolderPath.ReadOnly = true;
            this.pdfFolderPath.Size = new System.Drawing.Size(510, 28);
            this.pdfFolderPath.TabIndex = 1;
            this.pdfFolderPath.Text = "Select PDF Folder Path";
            // 
            // progressBar1
            // 
            progressBar1.Location = new System.Drawing.Point(139, 114);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new System.Drawing.Size(512, 23);
            progressBar1.TabIndex = 2;
            // 
            // Generate
            // 
            this.Generate.Location = new System.Drawing.Point(527, 153);
            this.Generate.Name = "Generate";
            this.Generate.Size = new System.Drawing.Size(124, 32);
            this.Generate.TabIndex = 0;
            this.Generate.Text = "Generate";
            this.Generate.UseVisualStyleBackColor = true;
            this.Generate.Click += new System.EventHandler(this.Generate_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(41, 29);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(88, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "PDF Folder Path:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(81, 119);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(51, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Progress:";
            // 
            // excelPath
            // 
            this.excelPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.excelPath.Location = new System.Drawing.Point(141, 67);
            this.excelPath.Name = "excelPath";
            this.excelPath.ReadOnly = true;
            this.excelPath.Size = new System.Drawing.Size(510, 28);
            this.excelPath.TabIndex = 1;
            this.excelPath.Text = "Select SpreadSheet File";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(41, 75);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(80, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "Excel File Path:";
            // 
            // excel
            // 
            this.excel.Location = new System.Drawing.Point(657, 67);
            this.excel.Name = "excel";
            this.excel.Size = new System.Drawing.Size(36, 30);
            this.excel.TabIndex = 4;
            this.excel.Text = "...";
            this.excel.UseVisualStyleBackColor = true;
            this.excel.Click += new System.EventHandler(this.excel_Click);
            // 
            // percentlabel
            // 
            percentlabel.AutoSize = true;
            percentlabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            percentlabel.Location = new System.Drawing.Point(658, 117);
            percentlabel.Name = "percentlabel";
            percentlabel.Size = new System.Drawing.Size(28, 17);
            percentlabel.TabIndex = 5;
            percentlabel.Text = "0%";
            // 
            // Form1

            this.Controls.Add(progressBar1);
            this.Controls.Add(percentlabel);
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(724, 197);
            
            this.Controls.Add(this.excel);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            
            this.Controls.Add(this.excelPath);
            this.Controls.Add(this.pdfFolderPath);
            this.Controls.Add(this.Generate);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "UI";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.RichTextBox pdfFolderPath;
        private System.Windows.Forms.Button Generate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.RichTextBox excelPath;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button excel;

        public static System.Windows.Forms.Label percentlabel;
        public static System.Windows.Forms.ProgressBar progressBar1;
    }
}

