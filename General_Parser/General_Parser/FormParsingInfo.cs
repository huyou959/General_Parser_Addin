using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace General_Parser
{
    public partial class FormParsingInfo : Form
    {
        public FormParsingInfo()
        {
            InitializeComponent();

            if (System.IO.File.Exists("c:/saver_temp/currentUrl.txt"))
            {
                this.textBox2.Text = System.IO.File.ReadAllText("c:/saver_temp/currentUrl.txt");
            }
        }



        String _url;
        String _tableName;
        String _metricNumber;
        String _processingNotes;

        public string Url { get => _url; set => _url = value; }
        public string TableName { get => _tableName; set => _tableName = value; }
        public string MetricNumber { get => _metricNumber; set => _metricNumber = value; }
        public string ProcessingNotes { get => _processingNotes; set => _processingNotes = value; }

        private void button3_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {

            TableName = textBox1.Text;
            _url = textBox2.Text;
            MetricNumber = textBox4.Text;

            _processingNotes = textBox3.Text;

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void FormParsingInfo_Load(object sender, EventArgs e)
        {

        }
    }
}
