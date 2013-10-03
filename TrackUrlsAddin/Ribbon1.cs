using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Word;
using System.Net;
using System.IO;
using System.Windows.Forms;

namespace TrackUrlsAddin
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        public void run(object o)
        {
            // disable certificate checks
            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

            int count = 0;
            string name = Globals.ThisAddIn.Application.ActiveDocument.Name;
            string[] words = name.Split(' ');
            string username = words[0];

            Hyperlinks links = Globals.ThisAddIn.Application.ActiveDocument.Hyperlinks;
            int total = links.Count;
            foreach (Hyperlink c in links)
            {
                // Hyperlink c = links[0];
                
                String address = c.Address;

                count += 1;

                if (o != null)
                {
                    Form1 f = (Form1)o;
                    f.progressBar1.Value = ( 100 / total ) * count;
                    f.statusLabel.Text = "Processing " + address;
                }

                if (address.StartsWith("http"))
                {
                    string redirectUrl = getRedirectUrl(address, username);
                    // MessageBox.Show(redirectUrl, "URL Builder", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    c.Address = redirectUrl;
                    
                }                               
            }

            bool writePDF = true;

            if (o != null)
                {
                    Form1 f = (Form1)o;
                    f.progressBar1.Value = 100;
                    f.statusLabel.Text = "Done";
                    writePDF = f.checkBox1.Checked;
            }
            string folder = Globals.ThisAddIn.Application.ActiveDocument.Path;

            string basename = username + " dgd report";
            string fileName =  basename + ".docx";

            Globals.ThisAddIn.Application.ActiveDocument.SaveAs2(Path.Combine(folder, fileName));
            if (writePDF)
            {
                exportPDF(folder, basename + ".pdf");
            }

            MessageBox.Show(String.Format("Replaced {0} urls in report for {1}.\n File Saved as {2}", count, username, fileName), "URL Builder", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }


        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            run(null);
        }

        private string getRedirectUrl(String url, String userName)
        {
                        
            HttpWebRequest httpWReq =
                 (HttpWebRequest)WebRequest.Create(Properties.Settings.Default.backend_url);
            string accessToken = Properties.Settings.Default.oauth_key;
            Encoding encoding = new UTF8Encoding();
            string postData = "{\"username\":\"" + userName + "\",\"url\":\"" + url + "\"}";
            byte[] data = encoding.GetBytes(postData);

            httpWReq.ProtocolVersion = HttpVersion.Version11;
            httpWReq.Method = "POST";
            httpWReq.ContentType = "application/json";//charset=UTF-8";


            httpWReq.Headers.Add(HttpRequestHeader.Authorization,
                "Bearer " + accessToken);
            httpWReq.ContentLength = data.Length;


            Stream stream = httpWReq.GetRequestStream();
            stream.Write(data, 0, data.Length);
            stream.Close();

            HttpWebResponse response = null;
            try
            {
                response = (HttpWebResponse)httpWReq.GetResponse();
            }
            catch (WebException e)
            {
                MessageBox.Show("Something went wrong with creating the redirect url. " + String.Format("email: {0}, url: {1}", userName, url), "URL Builder", MessageBoxButtons.OK, MessageBoxIcon.Error);                  
            }
            if (response.StatusCode != HttpStatusCode.Created)
            {
                throw new System.ArgumentException("Something went wrong with creating the redirect url",
                   String.Format("email: {0}, url: {1}", userName, url));
            }
            string s = response.ToString();
            StreamReader reader = new StreamReader(response.GetResponseStream());
            String jsonresponse = "";
            String temp = null;
            while ((temp = reader.ReadLine()) != null)
            {
                jsonresponse += temp;
            }
            string sub = jsonresponse.Substring(1, jsonresponse.Length - 2);
            // Console.WriteLine("Substring: {0}", sub);          

            return sub;

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            string url = Properties.Settings.Default.backend_url;
            if (url == "")
            {
               MessageBox.Show("Set backend url and auth token first.", "URL Builder", MessageBoxButtons.OK, MessageBoxIcon.Error);
               new SettingForm().ShowDialog();
               return;
            }

            Hyperlinks links = Globals.ThisAddIn.Application.ActiveDocument.Hyperlinks;

            Form1 form = new Form1(this);

            string name = Globals.ThisAddIn.Application.ActiveDocument.Name;
            string[] words = name.Split(' ');
            string email = words[0];

            form.textBox1.Text = email;

            foreach (Hyperlink c in links)
            {
                // Hyperlink c = links[0];

                String address = c.Address;

                if (address.StartsWith("http"))
                {
                    form.listBox1.Items.Add(address);        
                }

            }

            form.ShowDialog();
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            new SettingForm().ShowDialog();
        }

     
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            string name = Globals.ThisAddIn.Application.ActiveDocument.Name;
            string folder = Globals.ThisAddIn.Application.ActiveDocument.Path;
            
            string fileName = name + "_redirect.pdf";

            exportPDF(folder, fileName);
        }

        private static void exportPDF(string folder, string fileName)
        {
            Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(
                Path.Combine(folder, fileName), Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF,
                OpenAfterExport: true);
        }
    }
}
