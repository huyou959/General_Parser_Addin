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
    public partial class Form_ISTSSelect : Form
    {
        public Form_ISTSSelect()
        {
            InitializeComponent();
        }

        public bool IsTimeseries = false;

        private void button1_Click(object sender, EventArgs e)
        {
            IsTimeseries = true;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            IsTimeseries = true;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
