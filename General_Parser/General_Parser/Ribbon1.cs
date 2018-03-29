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
    }
}
