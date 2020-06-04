using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

using System.Windows.Forms;

namespace ExcelAddIn
{
    public partial class RibbonSecurity
    {
        public RibbonEditBox MailboxEdit
        {
            get
            {
                return editBoxMailbox;
            }
        }

        public static Excel.Application ExcelApp
        {
            get
            {
                return ThisAddIn.Application;
            }
        }

        public static ThisAddIn ThisAddIn
        {
            get
            {
                return Globals.ThisAddIn;
            }
        }

        private void buttonGenerate1_Click(object sender, RibbonControlEventArgs e)
        {
            string mailbox = MailboxEdit.Text;

            if (string.IsNullOrWhiteSpace(mailbox))
                throw new Exception("wrong mailbox address");

            Manager.Run(new FormProgressArgs()
            {
                Mailbox = mailbox,
                Type = FormProgressArgs.Types.table
            });

            
        }

        private void buttonGenerate2_Click(object sender, RibbonControlEventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog()
            {
                FileName = "Mail",
                Filter = "csv file|*.csv"
            };

            if (dialog.ShowDialog() != DialogResult.OK)
                return;

            string path = dialog.FileName;

            if (!IsPathValid(path))
                throw new Exception("path not valid");

            string mailbox = MailboxEdit.Text;

            if (string.IsNullOrWhiteSpace(mailbox))
                throw new Exception("wrong mailbox address");


            Manager.Run(new FormProgressArgs()
            {
                Path = path,
                Mailbox = mailbox,
                Type = FormProgressArgs.Types.csv
            });
        }

        private void RibbonSecurity_Load(object sender, RibbonUIEventArgs e)
        {
            Manager.Worker.DoWork += Worker_DoWork;
        }
    }
}
