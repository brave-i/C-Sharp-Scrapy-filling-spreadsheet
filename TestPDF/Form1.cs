using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.IO;
using System.Threading; // Required for this example


namespace TestPDF
{
    public partial class Form1 : Form
    {
        public static string pdfurl;
        public static string[] pdfFullPaths;
        public static string[] pdfFileLists;

        public static string excelurl = null;

        public static int progressvalue = 0;
        private bool isProcessRunning = false;


        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            string defaultstring = "Select PDF Folder Path";

            FolderBrowserDialog fbd = new FolderBrowserDialog();
            //fbd.Description = "Custom Description"; //not mandatory

            if (fbd.ShowDialog() == DialogResult.OK)
                pdfurl = fbd.SelectedPath;
            else
                pdfurl = defaultstring;//string.Empty;

            pdfFolderPath.Clear();
            pdfFolderPath.SelectedText = pdfurl;
            //Console.WriteLine(pdfurl);

            GetPDFFileLists(pdfurl, "*.pdf");

        }

        public static void GetPDFFileLists(string pdfurl, string extension)
        {
            //string pdfurl2 = @"D:\Project 08.22~09.22\William-50usd\transfer\Files\Diagrams\";
            //string extension = "*.pdf";
            //string[] filePaths = Directory.GetFiles(pdfurl, extension);

            //pdfurl = pdfurl + @"\";

            try
            {
                pdfFullPaths = Directory.GetFiles(pdfurl, extension);

            }
            catch(Exception e)
            {
                Console.WriteLine(e.StackTrace);
                return;
            }
            

            int nLen = pdfFullPaths.Length;
            pdfFileLists = new string[nLen];

            for (int k = 0; k < nLen; k++)
            {

                //Console.WriteLine(Path.GetFileName(pdfFullPaths[k]));      //file name
                //Console.WriteLine(pdfFullPaths[k]);                        //full name

                pdfFileLists[k] = Path.GetFileName(pdfFullPaths[k]);

            }
        }

        private void excel_Click(object sender, EventArgs e)
        {
            //string path = "";
            string defaultText = "Select SpreadSheet File";

            OpenFileDialog ofd = new OpenFileDialog();

            ofd.CheckFileExists = true;
            ofd.CheckPathExists = true;

            ofd.DefaultExt = "xlsx";
            ofd.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            ofd.FilterIndex = 1;

            //ofd.RestoreDirectory = true;
            //ofd.ReadOnlyChecked = true;
            //ofd.ShowReadOnly = true;

            excelPath.Clear();

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                excelurl = ofd.FileName;
                //Console.WriteLine(excelurl);                        //excel path

                excelPath.SelectedText = excelurl;
               

            }

            if (excelurl == null) excelPath.SelectedText = defaultText;
            else
            {
                excelPath.Clear();
                excelPath.SelectedText = excelurl;
            }

        }

        private void Generate_Click(object sender, EventArgs e)
        {
            
            if (excelurl == null)
            {
                MessageBox.Show("Please Select Spreadsheet file for Importing & Exporting!");
                return;
            }

            if (isProcessRunning)
            {
                MessageBox.Show("A process is already running.");
                return;
            }

            // Initialize the thread that will handle the background process
            Thread backgroundThread = new Thread(
                new ThreadStart(() =>
            {
                // Set the flag that indicates if a process is currently running
                isProcessRunning = true;

                /**********************************************************************/
                //1 . scrapy all informatn from pdf url

                Dictionary<int, ManagePDF.PDFLineInfo> dictionary = ManagePDF.ScrapyDataFromPDFiles(pdfFullPaths);

                //Console.WriteLine("++++++++++++++" + progressvalue +"++++++++++++++++");

                int secondcurrentpercent = 0;
                int diccount = 0;

                //2. update & add this infomation to excel file.
                ManageExcel.SpreadSheetOepnAndLoad(excelurl);

                foreach (KeyValuePair<int, ManagePDF.PDFLineInfo> item in dictionary)
                {
                    //if (diccount == 100) break;
                    //Console.WriteLine(item.Value.PartID, item.Value.PartName, item.Value.PartKey, item.Value.Quantity, item.Value.ProductName);
                    secondcurrentpercent = (int)(60 * diccount / dictionary.Count);

                    /*if(item.Value.ProductName.Length < 3)
                    {
                        MessageBox.Show(item.Value.ProductName);
                    }*/

                    ManageExcel.UpdateSpreadSheet(item.Value.PartID, item.Value.PartName, item.Value.PartKey, item.Value.Quantity, item.Value.ProductName);

                    //progressvalue += secondcurrentpercent;

                    progressBar1.BeginInvoke(new Action(() => Form1.progressBar1.Value = progressvalue+ secondcurrentpercent));
                    percentlabel.Text = (progressvalue + secondcurrentpercent).ToString() + "%";

                    diccount++;
                }

                progressvalue = progressvalue + secondcurrentpercent;

                ManageExcel.loopWholeSheet();         // backfill to excel file.

                progressvalue = 100;
                progressBar1.BeginInvoke(new Action(() => Form1.progressBar1.Value = progressvalue));
                percentlabel.Text = (progressvalue).ToString() + "%";

                ManageExcel.SpreadSheetSaveAndClose(excelurl);
                /**********************************************************************/

                MessageBox.Show("Processing completed!");

                // Reset the progress bar's value if it is still valid to do so
                //if (progressBar1.InvokeRequired)
                //    progressBar1.BeginInvoke(new Action(() => progressBar1.Value = 0));

                // Reset the flag that indicates if a process is currently running
                isProcessRunning = false;
                   }
                ));

                // Start the background process thread
                backgroundThread.Start();
            }
        }

    //}
}
