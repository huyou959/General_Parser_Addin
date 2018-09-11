using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace General_Parser
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender,RibbonControlEventArgs e)
        {
            Excel.Worksheet datasheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range MC = Globals.ThisAddIn.Application.Selection as Excel.Range;
            int r = 0;
            int c = 0;
            foreach (Excel.Range singleMC in MC)
            {
                if (singleMC.MergeCells)
                {
                    dynamic mergeAreaValue2 = singleMC.MergeArea.Value2;
                    object[,] vals = mergeAreaValue2 as object[,];

                    if (vals != null)
                    {
                        r = vals.GetLength(0);
                        c = vals.GetLength(1);
                    }
                    singleMC.UnMerge();
                    for (int i = singleMC.Row; i < singleMC.Row + r; i++)
                    {
                        for (int j = singleMC.Column; j < singleMC.Column + c; j++)
                        {
                            datasheet.Cells[i, j].Value = singleMC.Value;
                        }
                    }
                }
            }


        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Parse_Form Form1 = new Parse_Form(Globals.ThisAddIn.Application);
            Form1.TopMost = true;
            Form1.Show();
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveWorkbook.Close();
            Globals.ThisAddIn.Application.Quit();
        }

        private void IterateDirectory_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog fBD = new FolderBrowserDialog();
            if (fBD.ShowDialog()==DialogResult.OK)
            {
              String path=  fBD.SelectedPath;
                String[] files = System.IO.Directory.GetFiles(path);

                for (int i = 246+0; i < files.Length; i++)
                {

                    String thisFile = files[i];
                   Globals.ThisAddIn.Application.Workbooks.Open(thisFile);

                    //FileInfo finf = new FileInfo(thisFile);

                    Form_ISTSSelect fIT = new Form_ISTSSelect();
                    if (fIT.ShowDialog() == DialogResult.OK)
                    {
                        if (fIT.IsTimeseries)
                        {
                            //      System.IO.File.Copy(thisFile,)
                            Console.WriteLine(thisFile);
                            System.Diagnostics.Debug.WriteLine(thisFile);
                            System.IO.File.AppendAllText(path + "/timeseriesfiles.txt", thisFile + "\r\n");

                        }
                        else
                        {
                            System.IO.File.AppendAllText(path + "/Non_timeseriesfiles.txt", thisFile + "\r\n");
                        }
                    }


                    Globals.ThisAddIn.Application.ActiveWorkbook.Close(false);
                 //   break;
                }

            }
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {

            FormParsingInfo fPI = new FormParsingInfo();
            if (fPI.ShowDialog() == DialogResult.OK)
            {
                
               Excel.Range last = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastUsedRow = last.Row;
                int lastUsedColumn = last.Column;

                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Cells[lastUsedRow + 2, 1] = "ManualTableName";
                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Cells[lastUsedRow + 2, 2] = fPI.TableName;
                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Cells[lastUsedRow + 3, 1] = "ManualLink";

                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Cells[lastUsedRow + 3, 2] = fPI.Url;
                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Cells[lastUsedRow + 4, 1] = "ProcessingNotes";

                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Cells[lastUsedRow + 4, 2] = fPI.ProcessingNotes;
                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Cells[lastUsedRow + 5, 1] = "Metric#";

                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Cells[lastUsedRow + 5, 2] = fPI.MetricNumber;
            }
          //  OpenQA.Selenium.Chrome.ChromeDriver cD = new OpenQA.Selenium.Chrome.ChromeDriver();
         //   cD.Url = "http://www.cnn.com";
          //  MessageBox.Show("Hello World");
          //  ChromeURLGetter cURL = new ChromeURLGetter();
        //    IntPtr CHandle = ChromeURLGetter.GetChromeHandle();
        //    if (!CHandle.Equals(IntPtr.Zero))
 {
       //         string url = ChromeURLGetter.getChromeUrl(CHandle);
            }
        }
    }
}
