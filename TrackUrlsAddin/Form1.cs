using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TrackUrlsAddin
{
    public partial class Form1 : Form
    {
        Ribbon1 ribbon;

        public Form1(Ribbon1 ribbon)
        {
            InitializeComponent();
            this.ribbon = ribbon;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ribbon.run(this);
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }
    }
}
