using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Text.RegularExpressions;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace General_Parser
{
    public partial class Parse_Form : Form
    {
        public Parse_Form()
        {
            InitializeComponent();
        }

        private readonly Microsoft.Office.Interop.Excel.Application _excel0;
        private readonly Microsoft.Office.Interop.Excel.Application _excel1;
        private readonly Microsoft.Office.Interop.Excel.Application _excel2;
        private readonly Microsoft.Office.Interop.Excel.Application _excel3;
        private readonly Microsoft.Office.Interop.Excel.Application _excel4;

        public Parse_Form(Microsoft.Office.Interop.Excel.Application excel)
        {
            _excel0 = excel;
            _excel1 = excel;
            _excel2 = excel;
            _excel3 = excel;
            _excel4 = excel;
            InitializeComponent();
            refedit1._Excel = _excel0;
            refedit2._Excel = _excel1;
            refedit3._Excel = _excel2;
            refedit4._Excel = _excel3;
            refedit5._Excel = _excel4;
        }

        private void propertyGrid1_PropertyValueChanged(object s, PropertyValueChangedEventArgs e)
        {
            refedit1.Invalidate();
            refedit1.Focus();
            refedit2.Invalidate();
            refedit2.Focus();
            refedit3.Invalidate();
            refedit3.Focus();
            refedit4.Invalidate();
            refedit4.Focus();
            refedit5.Invalidate();
            refedit5.Focus();
        }

        private string title_pos = "Empty"; // position of title, Example: A1:B2

        private string data_str = "Empty"; // date on row or column

        private string date_range = "Empty"; // position of dates

        private string mainheader_range = "Empty"; // position of mainheader

        private string subheader_range = "Empty"; // position of subheader

        private string subheader_pos = "Empty"; // subheader is next to date or mainheader

        private string value_range = "Empty"; // position of values

        private void refedit1_CellChanged(object sender, EventArgs e)
        {
            Excel.Worksheet datasheet = Globals.ThisAddIn.Application.ActiveSheet;
            if (date_range != "Empty")
            {
                Excel.Range stuff_old = datasheet.Range[date_range];
                //stuff_old.Interior.ColorIndex = 0;
            }
            date_range = null;
            date_range = refedit1.Text;
            Excel.Range stuff_new = datasheet.Range[date_range];
            //stuff_new.Interior.ColorIndex = 3;
        }

        private void refedit2_CellChanged(object sender, EventArgs e)
        {
            Excel.Worksheet datasheet = Globals.ThisAddIn.Application.ActiveSheet;
            if (mainheader_range != "Empty")
            {
                Excel.Range stuff_old = datasheet.Range[mainheader_range];
                //stuff_old.Interior.ColorIndex = 0;
            }
            mainheader_range = null;
            mainheader_range = refedit2.Text;
            Excel.Range stuff_new = datasheet.Range[mainheader_range];
            //stuff_new.Interior.ColorIndex = 4;
        }

        private void refedit3_CellChanged(object sender, EventArgs e)
        {
            subheader_range = null;
            subheader_range = refedit3.Text;
        }

        private void refedit4_CellChanged(object sender, EventArgs e)
        {
            value_range = null;
            value_range = refedit4.Text;
        }

        private void refedit5_CellChanged(object sender, EventArgs e)
        {
            title_pos = null;
            title_pos = refedit5.Text;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                data_str = "Row";
            }
            else
            {
                if (radioButton2.Checked)
                {
                    data_str = "Col";
                }
                else
                {
                    data_str = "Empty";
                }
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                data_str = "Row";
            }
            else
            {
                if (radioButton2.Checked)
                {
                    data_str = "Col";
                }
                else
                {
                    data_str = "Empty";
                }
            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked)
            {
                subheader_pos = "Mainheader";
            }
            else
            {
                if (radioButton4.Checked)
                {
                    subheader_pos = "Date";
                }
                else
                {
                    subheader_pos = "Empty";
                }
            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked)
            {
                subheader_pos = "Mainheader";
            }
            else
            {
                if (radioButton4.Checked)
                {
                    subheader_pos = "Date";
                }
                else
                {
                    subheader_pos = "Empty";
                }
            }
        }

        private String createNeum(String neum)
        {
            var modifiedNeum = Regex.Replace(neum, @"\n+", "");

            if (String.IsNullOrEmpty(modifiedNeum)) throw new ArgumentNullException("neum");

            var i = 0;
            var first = 0;
            var newNuem = "";
            while (i < modifiedNeum.Length)
            {
                foreach (var letter in modifiedNeum)
                {

                    if (letter == ' ' || letter == modifiedNeum[modifiedNeum.Length - 1])
                    {
                        char firstLetter = modifiedNeum[first];
                        if (Char.IsUpper(firstLetter))
                        {
                            newNuem = newNuem + firstLetter;
                        }
                        if (Char.IsLower(firstLetter))
                        {
                            newNuem = newNuem + Char.ToUpper(firstLetter);
                        }
                        first = i + 1;
                    }
                    i++;
                }
            }
            return Convert.ToString(newNuem);
        }


        private void button_directory_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = "C:\\Users";
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                textBox_directory.Text = dialog.FileName + "\\" + "Parameters.txt";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Worksheet datasheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Worksheet newworksheet;
            string title = datasheet.Range[title_pos].Cells[1, 1].Value;
            string nk = createNeum(title);
            string[] lines = { title,nk, data_str, date_range, mainheader_range, subheader_range, subheader_pos, value_range};
            //if (System.IO.File.Exists(@"D:\Parameters.txt") == false)
            //{
            //    System.IO.File.CreateText(@"D:\Parameters.txt");
            //}

            System.IO.File.WriteAllLines(textBox_directory.Text,lines); // directory could change
            System.Diagnostics.Process.Start(textBox_directory.Text);
            /*
            newworksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
            if (data_str == "Row")  // dates on the row
            {
                newworksheet.Cells[1, 1].Value = "Date";
                newworksheet.Cells[1, 2].Value = "Frequency";
                newworksheet.Cells[1, 3].Value = "Neumkey";
                newworksheet.Cells[1, 4].Value = "Value";
                int col_num = datasheet.Range[value_range].Columns.Count;
                int row_num = datasheet.Range[value_range].Rows.Count;
                int k = 2;
                for (int i = 2; i <= row_num + 1; i++)
                {
                    for (int j = 2; j <= col_num + 1; j++)
                    {
                        newworksheet.Cells[k, 4].Value = datasheet.Range[value_range].Cells[i - 1, j - 1].Value;
                        newworksheet.Cells[k, 1].Value = datasheet.Range[date_range].Cells[1, j - 1].Value;
                        newworksheet.Cells[k, 2].Value = "Annually";
                        if (subheader_range == "Empty")
                        {
                            newworksheet.Cells[k, 3].Value = nk + "_" + datasheet.Range[mainheader_range].Cells[i - 1, 1].Value;
                        }
                        else
                        {
                            if (subheader_pos == "Mainheader")
                            {
                                newworksheet.Cells[k, 3].Value = nk + "_" + datasheet.Range[mainheader_range].Cells[i - 1, 1].Value + "_" + datasheet.Range[subheader_range].Cells[j - 1, 1].Value;
                            }
                            else
                            {
                                newworksheet.Cells[k, 3].Value = nk + "_" + datasheet.Range[mainheader_range].Cells[i - 1, 1].Value + "_" + datasheet.Range[subheader_range].Cells[1, j - 1].Value;
                            }
                        }
                        
                        k++;
                    }
                }

            }

            else //dates on the column
            {
                newworksheet.Cells[1, 1].Value = "Date";
                newworksheet.Cells[1, 2].Value = "Frequency";
                newworksheet.Cells[1, 3].Value = "Neumkey";
                newworksheet.Cells[1, 4].Value = "Value";
                int col_num = datasheet.Range[value_range].Columns.Count;
                int row_num = datasheet.Range[value_range].Rows.Count;
                int k = 2;
                for (int i = 2; i <= row_num + 1; i++)
                {
                    for (int j = 2; j <= col_num + 1; j++)
                    {
                        newworksheet.Cells[k, 4].Value = datasheet.Range[value_range].Cells[i - 1, j - 1].Value;
                        newworksheet.Cells[k, 1].Value = datasheet.Range[date_range].Cells[i - 1, 1].Value;
                        newworksheet.Cells[k, 2].Value = "Annually";
                        if (subheader_range == "Empty")
                        {
                            newworksheet.Cells[k, 3].Value = nk + "_" + datasheet.Range[mainheader_range].Cells[1, j - 1].Value;
                        }
                        else
                        {
                            if (subheader_pos == "Mainheader")
                            {
                                newworksheet.Cells[k, 3].Value = nk + "_" + datasheet.Range[mainheader_range].Cells[1, j - 1].Value + "_" + datasheet.Range[subheader_range].Cells[1, j - 1].Value;
                            }
                            else
                            {
                                newworksheet.Cells[k, 3].Value = nk + "_" + datasheet.Range[mainheader_range].Cells[1, j - 1].Value + "_" + datasheet.Range[subheader_range].Cells[j - 1, 1].Value;
                            }
                        }

                        k++;
                    }
                }
            }
            */
        }

    }
}
