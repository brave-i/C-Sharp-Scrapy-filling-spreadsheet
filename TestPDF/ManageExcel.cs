using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Bytescout.Spreadsheet;
using System.Diagnostics;
using System.IO;

using Bytescout.Spreadsheet.BaseClasses;

using System.Data;
using System.Windows.Forms;
using System.Threading;

namespace TestPDF
{
    public class ManageExcel
    {
        public static int usedRows = 0;
        public static Spreadsheet document = null;
        public static Worksheet worksheet = null;
        //public static Worksheet pworksheet = null;
        public static int estimaterows = 0;
        public static int previousIndex = 0;
        public static int nstacked = 0;

        public static List<string> list = null;
        


        public static string findmatchingID(string arg)
        {
            
            // arg = categoryid

            string partID = null;

            string column_A = null;
            string column_C = "Catalog Number";
            int nIndex = 0;

            while (column_C != "")
            {

                column_C = worksheet.Cell(nIndex, 2).ValueAsString;

                if (arg.Contains(column_C) == true)
                {
                    column_A = worksheet.Cell(nIndex, 0).ValueAsString;

                    if (column_A == "")
                    {
                        nIndex++;
                        continue;
                    }

                    if (partID == null)
                        partID = column_A;

                    else
                    {
                        if (partID == "")
                            partID = column_A;
                        else
                        {
                            if (partID.Contains(column_A) == false)
                                partID = partID + " OR " + column_A;
                        }

                    }

                }
                nIndex++;
            }
            //estimaterows = nIndex;

            return partID;
        }

        public static bool loopWholeSheet()
        {
            string column_A = null;
            string column_C = null;
            string findID = null;

            int currentprogress = 0;

            estimaterows = getusedrowcount();

            for (int i = 0; i < estimaterows; i++)
            {
                currentprogress = (int)(20 * i / estimaterows);

                //updateing value
                
                Form1.progressBar1.BeginInvoke(new Action(() => Form1.progressBar1.Value = Form1.progressvalue + currentprogress));
                Form1.percentlabel.Text = (Form1.progressvalue + currentprogress).ToString() + "%";

                column_A = worksheet.Cell(i, 0).ValueAsString;

                if (column_A == "")
                {
                    column_C = worksheet.Cell(i, 2).ValueAsString;
                    findID = findmatchingID(column_C);
                    worksheet.Cell(i, 0).Value = findID;
                }

                //Console.WriteLine("*****************" + i + "**************************");
            }
            return true;
        }

        public static int getusedrowcount()
        {
            string column_C = "cateloge";
            int nusedRow = 0;

            while (column_C != "")
            {
                column_C = worksheet.Cell(nusedRow, 2).ValueAsString;
                nusedRow++;

            }
            return nusedRow;
        }


        public static string RemoveAllRegex(string expression)
        {
            string result = "";

            result = expression.Replace(".pdf", "");

            result = result.Replace(" ", "");
            result = result.Replace("_", "");
            result = result.Replace("-", "");
            result = result.Replace("(", "");
            result = result.Replace(")", "");
            result = result.Replace(")", "");

            return result;
        }

        public static bool MatchingRegex(string col_K, string pdfTitle)
        {
            string m_colk = RemoveEscapeQuoteLowerSpace(col_K);                                   
            string m_pdfTitle = pdfTitle;      

            bool b = m_colk.Contains(m_pdfTitle);

            return b;

        }
        public static bool IsValidNumber(string arg)
        {
            int pnumber = 0;

            string marg = arg.Replace(" ", "");

            if ((arg.Contains(".") == true) || (arg.Contains("&") == true) || (arg.Contains("~") == true))
                return false;

             try
            {
                pnumber = Int32.Parse(marg);
                return true;
            }
            catch
            {
                return false;
            }

            
        }
        
        public static bool MatchingIDFromExcel(string excelID, string pdfID)
        {
            
            excelID = excelID.Replace(" ", "");
            pdfID = pdfID.Replace(" ", "");// i believe.

            if (excelID.Contains(".1") == true)
                excelID = excelID.Replace(".1", "");
            
            if (excelID.Contains("?") == true)
                excelID = excelID.Replace("?", "");
            
            if (excelID.Contains("~") == true)
            {
                string[] words = excelID.Split('~');
                excelID = words[0];
            }
            if (pdfID.Contains("~") == true)
            {
                string[] words2 = pdfID.Split('~');
                pdfID = words2[0];
            }


            //if (excelID.Contains(pdfID) == true)
            if(excelID == pdfID)
                return true;

            return false;

            /*
            pdfID = pdfID.Replace(" ", "");
            if (IsValidNumber(pdfID) == false)
                return false;

            excelID = excelID.Replace(" ", "");
            if ((excelID.Contains(".") == true) || (excelID.Contains("&") == true) || (excelID.Contains("~") == true))
                return false;


            string m_exp1 = pdfID;
            string m_exp2 = excelID;

            //bool b = m_exp2.Contains(m_exp1);

            bool b = false;

            if (m_exp2 == m_exp1) b = true;
            

            return b;
            */
        }

        public static string RemoveEscapeQuoteLowerSpace(string colk)
        {
            string sfinal = colk.ToLower();

            sfinal = sfinal.Replace(" ", "");

            sfinal = sfinal.Replace("-", "");
            sfinal = sfinal.Replace("(", "");
            sfinal = sfinal.Replace(")", "");

            return sfinal;
        }


        public static int getRowIndex(Worksheet worksheet, string partID, string partName, string partKey, int Quanity, string productName)
        {

            string col_k = "start";
            string col_i = "";

            int i = 0;
            
            bool findproductName = false;


            while (col_k != "")
            {
                /*if(productName.Contains("r92592112") == true)
                {
                    int aaa = 5;
                }*/
                col_k = worksheet.Cell(i, 10).ValueAsString;    //column k

                if(MatchingRegex(col_k, productName) == true)
                {
                    findproductName = true;
                    col_i = worksheet.Cell(i, 8).ValueAsString;

                    if (MatchingIDFromExcel(col_i, partKey) == true)
                    {
                        
                        if (!list.Contains(i.ToString()))
                        {
                            list.Add(i.ToString());
                            return i;
                        }

                    }                  
                }

                usedRows = i;
                i++;
            }


            /*while (col_k != "")
            {
                col_k = worksheet.Cell(i, 10).ValueAsString;    //column j
                col_i = worksheet.Cell(i, 8).ValueAsString;

                //if (col_k == "") return -1;

                if (MatchingRegex(col_k , productName) == true)
                {
                    findproductName = true;
                
                    if (MatchingMatchingID(partKey, col_i) == true)
                        return i;                    
                }

                usedRows = i;
                i++;
            }*/

            if (findproductName == true)
            {
                //MessageBox.Show("PartID: " + partID + ", PartName: " + partName + ", PartKey: " + partKey + ", ProductName: " + productName);
               

                return -2;
            }
                

            else
                return -1;
        }

        public static bool SpreadSheetOepnAndLoad(string spreadsheeetFullPath)
        {
            // Create new Spreadsheet
            document = new Spreadsheet();
            document.LoadFromFile(spreadsheeetFullPath);

            // Get first worksheet
            worksheet = document.Workbook.Worksheets.ByName("Parts");
            //pworksheet = worksheet;

            list = new List<string>();

            return true;
        }

        public static bool SpreadSheetSaveAndClose(string spreadsheeetFullPath)
        {
            // Save document
            //MessageBox.Show(spreadsheeetFullPath);

            document.SaveAs(spreadsheeetFullPath);

            //Close document
            document.Dispose();
            document.Close();

            return true;
        }


        public static void UpdateSpreadSheet(string partID, string partName, string partKey, int Quanity, string productName)
        {
            // Create new Spreadsheet
            //Spreadsheet document = new Spreadsheet();

            //document.LoadFromFile("D:\\sample.xlsx");
            //document.LoadFromFile("D:\\sample.xlsx");

            // Read cell value
            //Console.WriteLine(worksheet.Cell(0, 0).ValueAsString);

            //string partID = "B020001-Z1";
            //string partName = "Link-2T";
            //string partKey = "10";
            //int Quanity = 1;
            //string productName = "UF80_(580886)_ExplodedView_PartsList.pdf";

            //bellow step is about 1

            int nIndex = getRowIndex(worksheet, partID, partName, partKey, Quanity, productName);

            //int kkk = 

            if (nIndex == -1)// no find. must add at last.
            {
                //int asd = 5;

                usedRows++;

                // Add new row
                worksheet.Rows.Insert(usedRows, 1);
                // Set values
                worksheet.Rows[usedRows][1].Value = partName;
                worksheet.Rows[usedRows][8].Value = partKey;
                worksheet.Rows[usedRows][9].Value = Quanity;
                worksheet.Rows[usedRows][10].Value = productName.Replace(".pdf","");
                
            }
            else if(nIndex == -2)
            {
                //int asd = 6;
            }
            else
            {

                /*if (nIndex == 518)
                {
                    Console.WriteLine("partID=>" + partID + "partNam=>=" + partName + "partKey=>" + partKey + "productName=>" + productName);
                    MessageBox.Show("partID: " + partID + "PartNam: " + partName + " PartKey=>" + partKey + " ProductName=>" + productName);
                }*/

                worksheet.Cell(nIndex, 0).Value = partID;

            }

        }
    }
}

