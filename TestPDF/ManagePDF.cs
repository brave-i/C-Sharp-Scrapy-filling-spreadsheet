using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Bytescout.PDFExtractor;
using System.Windows.Forms;

namespace TestPDF
{
    
    public class ManagePDF
    {

        public static int nTotalIndex = 0;
        public static bool findTitle = false;
        public static bool findheader = false;
        
        public struct PDFLineInfo
        {
            public string PartID;
            public string PartName;
            public string PartKey;
            public int Quantity;
            public string ProductName;

        };

        //public static Dictionary<int, PDFLineInfo> dictionary;


        public static string RemoveSpace(string spacestr)
        {
            return spacestr.Replace(" ", "");
        }

        public static string detectTitle(string line)
        {
            if (line.ToLower().Contains("description") == true)
                return "notable";

            /*if (line.ToLower().Contains("demo") == true)
            {
                int aaa = 5;
            }*/

            int nStartIndex = 0;

            string find = "part";
            string product = null;
            string removedspace = line.Replace(" ", "");

            var s2 = line.ToLower().Replace(find, "");

            int nrepeat = (line.Length - s2.Length) / find.Length;

            if ((nrepeat == 0) || (nrepeat == 1))
            {
                product = removedspace.ToLower();
                product = product.Replace("part", "");
                product = product.Replace("*demo*", "");
                product = product.Replace("ole", "");
            }

            else
            {
                do
                {

                    product = removedspace.Substring(nStartIndex * removedspace.Length / nrepeat, removedspace.Length / nrepeat).ToLower();
                    nStartIndex++;
                } while ((product.ToLower().Contains("*") == true) && (nStartIndex < nrepeat));


            }

            if (product.ToLower().Contains("*") == true)
            {
                product = product.Replace("o*le", "");
            }

            product = product.Replace("sole", "");
            product = product.Replace("sole", "");
            product = product.Replace("-", "");

            product = product.Replace("parts", "");
            product = product.Replace("part", "");

            product = product.Replace("lists", "");
            product = product.Replace("list", "");

            product = product.Replace("(", "");
            product = product.Replace(")", "");
            product = product.Replace("_", "");

            if(product == "")
                return "notable";

            return product;

        }


        public static Dictionary<int, PDFLineInfo> ScrapyDataFromPDFiles(string[] urllist)
        {
            PDFLineInfo temp = new PDFLineInfo();
            Dictionary<int, PDFLineInfo> dictionary = new Dictionary<int, PDFLineInfo>();
            //dictionary = new Dictionary<int, PDFLineInfo>();


            // Create Bytescout.PDFExtractor.TextExtractor instance
            TextExtractor extractor = new TextExtractor();
            extractor.RegistrationName = "demo";
            extractor.RegistrationKey = "demo";

            for(int nFileIndex = 0; nFileIndex < urllist.Length; nFileIndex++)
            {

                //string currentFileName = "sample2.pdf";
                string currentFileName = urllist[nFileIndex];
                string currentTitleName = "";

                // Load each PDF Document
                extractor.LoadDocumentFromFile(currentFileName);
                int pageCount = extractor.GetPageCount();


                //most of all case i = 0 but one case  i = 0
                int pdfDocumentType = -1;// 1: material type 2: spirit type 3: Empty type.

                /*if(currentFileName.Contains("R92(592112)_ExpViewPartList") == true)
                {
                    int zz = 5;
                }*/

                for (int i = 1; i < pageCount; i++)
                {
                    if (currentTitleName.Contains("notable") == true) break;

                    //if (extractor.Find(i, "Dyaco", false))
                    {
                        //extractor.SetExtractionArea(0, 0, 800, 2000);
                        string wholetext = extractor.GetTextFromPage(i);
                        //Console.WriteLine(wholetext);

                        string[] lines = wholetext.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                        //Console.WriteLine("Length ===================== >" + lines.Length);
                        //if line.getlen is not 4 alert!!

                        
                        //1 . Notify Header Strings
                        int j = 0;


                        while (j < lines.Length)
                        {
                            if ((lines[j].ToLower().Contains("part") == true) && (findTitle == false))
                            {
                                //Console.WriteLine("Title = = = = => " + detectTitle(lines[i]));
                                currentTitleName = detectTitle(lines[j]);
                                findTitle = true;

                                j++;
                                continue;
                            }
                            if (currentTitleName.Contains("notable") == true) break;

                            if (findTitle == false)
                            {
                                if (j > 2)
                                {
                                    currentTitleName = "notable";
                                    break;
                                }
                                j++;
                                continue;
                            }

                            var array = lines[j].Split(new string[] { "  " }, StringSplitOptions.RemoveEmptyEntries);

                            if ((findheader == false) && (findTitle == true))
                            {
                                if ((lines[j].ToLower().Contains("dyaco") == true) || (lines[j].ToLower().Contains("material") == true) || 
                                    (lines[j].ToLower().Contains("spirit") == true) || (lines[j].ToLower().Contains("no") == true) || 
                                    (lines[j].ToLower().Contains("part") == true) || (lines[j].ToLower().Contains("qty") == true))
                                {
                                    findheader = true;

                                    {
                                        if (lines[j].ToLower().Contains("material") == true) pdfDocumentType = 1;

                                        else if (lines[j].ToLower().Contains("spirit") == true) pdfDocumentType = 2;

                                        else
                                        {
                                            if (array.Length > 2)
                                            {
                                                if (array[2].ToLower().Contains("part") == true)
                                                    pdfDocumentType = 2;
                                            }
                                            pdfDocumentType = 3;
                                        }
                                    }

                                    j++;
                                    continue;
                                }

                                if (array.Length < 4)
                                {
                                    j++;
                                    continue;
                                }

                            }

                            if ((lines[j].Contains("(TRIAL VER. PDF Extractor SDK 8.4.1.2829.888331924)") == true) || ((lines[j].Contains("TRIAL VERSION EXPIRES 90 DAYS AFTER INSTALLATION") == true)))
                            {
                                j++;
                                continue;
                            }

                            if (pdfDocumentType == -1)
                            {
                                MessageBox.Show("Can not get pdf Type");
                            }

                            if (pdfDocumentType == 3)
                            {
                                if ((array.Length != 3) ||(array[0].Length > 5))
                                {
                                    j++;
                                    continue;
                                }
                                
                            }
                            else
                            {
                                if (array.Length != 4)
                                {
                                    j++;
                                    continue;
                                }
                            }
                            
                            

                            //Console.WriteLine("Document Type =====>" + pdfDocumentType);

                            //Console.WriteLine(RemoveSpace(array[0]));
                            //Console.WriteLine(RemoveSpace(array[1]));
                            //Console.WriteLine(RemoveSpace(array[2]));
                            //Console.WriteLine(RemoveSpace(array[3]));
                            /*if(array[0].Contains("57") == true)
                            {
                                int awe = 5;
                            }*/

                            switch (pdfDocumentType)
                            {
                                case 1:
                                    temp.PartID = RemoveSpace(array[1]);
                                    temp.PartName = RemoveSpace(array[2]);
                                    temp.PartKey = RemoveSpace(array[0]);
                                    temp.Quantity = Int32.Parse(RemoveSpace(array[3])); //no change

                                    break;

                                case 2:
                                    temp.PartID = RemoveSpace(array[2]);
                                    temp.PartName = RemoveSpace(array[1]);
                                    temp.PartKey = RemoveSpace(array[0]);
                                    temp.Quantity = Int32.Parse(RemoveSpace(array[3])); //no change

                                    break;

                                case 3:
                                    temp.PartID = "";                                   //empty
                                    temp.PartName = RemoveSpace(array[1]);
                                    temp.PartKey = RemoveSpace(array[0]);
                                    temp.Quantity = Int32.Parse(RemoveSpace(array[2])); //no change

                                    break;

                            }

                            temp.ProductName = currentTitleName;                 //no change

                            /*if(currentTitleName.Length <3)
                            {
                                int qqq = 5;
                            }*/

                            j++;

                            //2. Add values to PDFLineInfo
                            dictionary.Add(nTotalIndex, temp);

                            nTotalIndex++;

                        }

                    }
                }

                findheader = false;
                findTitle = false;
                currentTitleName = "";

                int currentpercent = (int)(20 * nFileIndex / urllist.Length);

                Console.WriteLine("*******" + nFileIndex + "*********"+currentpercent + "********");

                //updateing value
                Form1.progressvalue = currentpercent;
                Form1.progressBar1.BeginInvoke(new Action(() => Form1.progressBar1.Value = currentpercent));
                Form1.percentlabel.Text = currentpercent.ToString() + "%";

            }

            return dictionary;
        }
    }
}
