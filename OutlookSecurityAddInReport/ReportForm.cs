using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using SharedResources;

namespace SecurityAddInReport
{

    public partial class ReportForm : Form
    {
        public DataGridView Grid
        {
            get
            {
                return dataGridViewSelections;
            }
        }

        protected Outlook.Application App
        {
            get
            {
                return Globals.ThisAddIn.Application;
            }
        }

        public string report_to = "suspicious-box@outlook.com";
        
        public List<Outlook.MailItem> items = new List<Outlook.MailItem>();

        public ReportForm()
        {
            InitializeComponent();

            this.AcceptButton = buttonReport;
            this.CancelButton = CancelButton;
        }

        private void ProcessMail(Outlook.MailItem mailObject)
        {
            
            string path = Path.Combine(Path.GetTempPath(), OutlookFunctions.ValidFilename(mailObject.Subject + ".msg"));
            try
            {
                mailObject.SaveAs(path, Outlook.OlSaveAsType.olMSGUnicode);
            }
            catch (Exception exc)
            {
                Debug.WriteLine(exc.Message);
                return;
            }

            string type = string.Empty;
            if (radioButtonSpam.Checked)
                type = radioButtonSpam.Text;
            else if (radioButtonPhishing.Checked)
                type = radioButtonPhishing.Text;

            type = string.Format("[ [ {0} TYPE ] ]", type.ToUpper());

            Outlook.MailItem newMail = App.CreateItem(Outlook.OlItemType.olMailItem);
            newMail.Subject = mailObject.Subject;

            Outlook.MAPIFolder parentFolder = mailObject.Parent as Outlook.MAPIFolder;

            // SMTP address
            string rootFolder = parentFolder.Store.GetRootFolder().Name;

            var account = OutlookFunctions.GetAccount(App, rootFolder);
            newMail.Sender = account.CurrentUser.AddressEntry;
            newMail.SendUsingAccount = account;
            newMail.Attachments.Add(path);
            newMail.To = report_to;

            // force generation of full HTML from Inspector
            var inspector = newMail.GetInspector;

            // signature
            newMail.HTMLBody = type + "<br><br>" + textBoxComment.Text + "<br><br>" + newMail.HTMLBody;

            try
            {
                newMail.Send();
            } 
            catch (Exception exc)
            {
                Debug.WriteLine(exc.Message);
                return;
            }
            
            mailObject.Categories = "Reported";
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {
            foreach (var item in items)
            {
                ProcessMail(item);
            }
            this.Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
