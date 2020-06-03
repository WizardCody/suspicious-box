using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Outlook = Microsoft.Office.Interop.Outlook;

namespace SecurityAddInReport
{
    class RibbonSecurityReportExplorer : RibbonSecurityReport
    {

        public RibbonSecurityReportExplorer() : base()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // RibbonSecurityReportExplorer
            // 
            this.Name = "RibbonSecurityReportExplorer";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.ResumeLayout(false);

        }

        protected override void LoadData()
        {
            var activeExplorer = App.ActiveWindow() as Outlook.Explorer;
            var selections = activeExplorer.Selection;

            Form.Grid.Rows.Clear();
            Form.items.Clear();

            foreach (dynamic selection in selections)
            {
                if (selection.Class == (int)Outlook.OlObjectClass.olMail)
                {
                    Form.Grid.Rows.Add(selection.Subject, selection.ReceivedTime);
                    Form.items.Add(selection);
                }
            }
        }

    }
}
