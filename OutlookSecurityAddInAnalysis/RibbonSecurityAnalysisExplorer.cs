using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools;
using System.Diagnostics;
using SharedResources;

namespace SecurityAddInAnalysis
{
    class RibbonSecurityAnalysisExplorer : RibbonSecurityAnalysis
    {
        public RibbonSecurityAnalysisExplorer() : base()
        {
            InitializeComponent();
        }

        protected override void HeaderAnalysisControl_Load(object sender, EventArgs e)
        {
            base.HeaderAnalysisControl_Load(sender, e);

            var activeExplorer = App.ActiveWindow() as Outlook.Explorer;

            activeExplorer.SelectionChange += ActiveExplorer_SelectionChange;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // RibbonSecurityAnalysisExplorer
            // 
            this.Name = "RibbonSecurityAnalysisExplorer";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.ResumeLayout(false);

        }
        

        private void ActiveExplorer_SelectionChange()
        {
            if (!TaskPaneHeaderAnalysis.Visible)
                return;

            ApplySelection();
        }

        public void ApplySelection()
        {
            var activeExplorer = App.ActiveWindow() as Outlook.Explorer;

            if (activeExplorer.Selection.Count > 0)
            {
                dynamic item = activeExplorer.Selection[1];
                if (item.Class == (int)Outlook.OlObjectClass.olMail)
                {
                    HeaderAnalysisControl.TextBox.Text = MailItemProperties.GetHeader(item);
                }
            }
        }

        private void ProcessSelections(Classification type)
        {
            foreach (dynamic selection in App.ActiveExplorer().Selection)
            {
                if (selection.Class == (int)Outlook.OlObjectClass.olMail)
                {
                    ProcessMail(selection, type);
                }
            }
        }

        protected override void buttonSample_Click(object sender, RibbonControlEventArgs e)
        {
            base.buttonSample_Click(sender, e);

            ProcessSelections(Classification.noSample);
        }

        protected override void buttonLegit_Click(object sender, RibbonControlEventArgs e)
        {
            base.buttonLegit_Click(sender, e);

            ProcessSelections(Classification.legit);
        }

        protected override void buttonSpam_Click(object sender, RibbonControlEventArgs e)
        {
            base.buttonSpam_Click(sender, e);

            ProcessSelections(Classification.spam);
        }

        protected override void buttonMalicious_Click(object sender, RibbonControlEventArgs e)
        {
            base.buttonMalicious_Click(sender, e);

            ProcessSelections(Classification.malicious);
        }

        protected override void buttonSimulation_Click(object sender, RibbonControlEventArgs e)
        {
            base.buttonSimulation_Click(sender, e);

            ProcessSelections(Classification.simulation);
        }

        protected override void toggleButtonShowHeaders_Click(object sender, RibbonControlEventArgs e)
        {
            base.toggleButtonShowHeaders_Click(sender, e);

            if (TaskPaneHeaderAnalysis.Visible)
            {
                ApplySelection();
            }
        }


    }
}
