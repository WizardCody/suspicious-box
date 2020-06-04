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

namespace OutlookSecurityAddInReport
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

        public string Report_Mail { get; } = "suspicious-box@outlook.com";
        
        public List<Outlook.MailItem> Items { get; } = new List<Outlook.MailItem>();

        private FormProgressManager Manager { get; } = new FormProgressManager();

        public ReportForm()
        {
            InitializeComponent();

            Manager.Worker.DoWork += Worker_DoWork;
            Manager.Worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
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
            newMail.To = Report_Mail;

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
            
            //mailObject.Categories = "Reported";
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;

            Manager.CurrentStage = FormProgressManager.ProcessStage.Process;

            int currentItem = 0;
            foreach (var item in Items)
            {
                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }

                currentItem++;

                Manager.SetStatus(currentItem, Items.Count);

                ProcessMail(item);
            }

            if (currentItem < Items.Count)
                Manager.SetStatus(Items.Count, Items.Count);
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {
            Manager.Run();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Close();
        }

        private void ReportForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Manager.FormProgress != null)
                Manager.FormProgress.Close();
        }
    }
}
