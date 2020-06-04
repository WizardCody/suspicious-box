using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Outlook = Microsoft.Office.Interop.Outlook;

namespace SecurityAddInReport
{
    class RibbonSecurityReportInspector : RibbonSecurityReport
    {

        public RibbonSecurityReportInspector() : base()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // RibbonSecurityReportInspector
            // 
            this.Name = "RibbonSecurityReportInspector";
            this.RibbonType = "Microsoft.Outlook.Mail.Read";
            this.ResumeLayout(false);

        }

        protected override void LoadData()
        {
            var activeInspector = App.ActiveWindow() as Outlook.Inspector;

            Form.Grid.Rows.Clear();
            Form.Items.Clear();

            dynamic item = activeInspector.CurrentItem;

            if (item.Class == (int)Outlook.OlObjectClass.olMail)
            {
                Form.Grid.Rows.Add(item.Subject, item.ReceivedTime);
                Form.Items.Add(item);
            }
        }

    }
}
