using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using SharedResources;

namespace SecurityAddInAnalysis
{
    class RibbonSecurityAnalysisInspector : RibbonSecurityAnalysis
    {
        public RibbonSecurityAnalysisInspector() : base()
        {
            InitializeComponent();
        }

        protected override void RibbonSecurityAnalysis_Load(object sender, RibbonUIEventArgs e)
        {
            base.RibbonSecurityAnalysis_Load(sender, e);

            if (TaskPaneHeaderAnalysis != null & TaskPaneHeaderAnalysis.Visible)
                ApplySelection();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // RibbonSecurityAnalysisInspector
            // 
            this.Name = "RibbonSecurityAnalysisInspector";
            this.RibbonType = "Microsoft.Outlook.Mail.Read";
            this.ResumeLayout(false);

        }

        public void ApplySelection()
        {
            var activeExplorer = App.ActiveWindow() as Outlook.Inspector;

            dynamic item = activeExplorer.CurrentItem;

            if (item.Class == (int)Outlook.OlObjectClass.olMail)
            {
                HeaderAnalysisControl.TextBox.Text = MailItemProperties.GetHeader(item);
            }
        }

        protected override void toggleButtonShowHeaders_Click(object sender, RibbonControlEventArgs e)
        {
            base.toggleButtonShowHeaders_Click(sender, e);

            if (TaskPaneHeaderAnalysis != null & TaskPaneHeaderAnalysis.Visible)
            {
                ApplySelection();
            }
        }

    }
}