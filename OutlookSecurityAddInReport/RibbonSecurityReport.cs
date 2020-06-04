using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookSecurityAddInReport
{

    public partial class RibbonSecurityReport
    {
        public ReportForm Form { get; private set; } = null;

        protected Outlook.Application App
        {
            get
            {
                return ThisAddIn.Application;
            }
        }

        protected ThisAddIn ThisAddIn
        {
            get
            {
                return Globals.ThisAddIn;
            }
        }

        private void buttonReport_Click(object sender, RibbonControlEventArgs e)
        {
            if (Form == null || Form.IsDisposed)
            {
                Form = new ReportForm();
            }

            LoadData();

            if (!Form.Visible) {
                Form.Show();
            }

            Form.Activate();
        }

        protected virtual void LoadData()
        {

        }

    }
}
