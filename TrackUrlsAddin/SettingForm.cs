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
    public partial class SettingForm : Form
    {
        public SettingForm()
        {
            InitializeComponent();
            urlTextBox.Text = Properties.Settings.Default.backend_url;
            oauthTokenTextBox.Text = Properties.Settings.Default.oauth_key;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.backend_url = urlTextBox.Text;
            Properties.Settings.Default.oauth_key = oauthTokenTextBox.Text;
            Properties.Settings.Default.Save();
            this.Close();
        }
    }
}
