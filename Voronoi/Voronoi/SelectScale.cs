using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Voronoi
{
    public partial class SelectScale : Form
    {
        public int OldScale;
        public int NewScale;
        public SelectScale()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OldScale = (int)numericUpDown1.Value;
            NewScale = (int)numericUpDown2.Value;
            this.DialogResult = DialogResult.OK;
        }
    }
}