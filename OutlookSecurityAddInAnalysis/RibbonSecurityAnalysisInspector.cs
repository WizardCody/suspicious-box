using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using SharedResources;

namespace OutlookSecurityAddInAnalysis
{
    class RibbonSecurityAnalysisInspector : RibbonSecurityAnalysis
    {
        public RibbonSecurityAnalysisInspector() : base()
        {
            InitializeComponent();

            Manager.Worker.DoWork += Worker_DoWork;
        }

        private void Worker_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            var worker = sender as System.ComponentModel.BackgroundWorker;

            var type = (Classification)e.Argument;

            Manager.CurrentStage = FormProgressManager.ProcessStage.Process;

            Manager.SetStatus(0, 1);

            if (worker.CancellationPending)
            {
                e.Cancel = true;
                return;
            }

            dynamic item = App.ActiveInspector().CurrentItem;

            if (item.Class == (int)Outlook.OlObjectClass.olMail)
            {
                ProcessMail(item, type);
            }

            Manager.SetStatus(1, 1);
        }

        protected override void RibbonSecurityAnalysis_Load(object sender, RibbonUIEventArgs e)
        {
            base.RibbonSecurityAnalysis_Load(sender, e);

            if (TaskPaneHeaderAnalysis != null && TaskPaneHeaderAnalysis.Visible)
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
                HeaderAnalysisControl.CurrentItem = item;
            }
        }

        protected override void toggleButtonShowHeaders_Click(object sender, RibbonControlEventArgs e)
        {
            base.toggleButtonShowHeaders_Click(sender, e);

            if (TaskPaneHeaderAnalysis != null && TaskPaneHeaderAnalysis.Visible)
            {
                ApplySelection();
            }
        }

        private void ProcessCurrentItem(Classification type)
        {
            Manager.Run(type);
        }

        protected override void buttonSample_Click(object sender, RibbonControlEventArgs e) 
        {
            ProcessCurrentItem(Classification.noSample);
        }

        protected override void buttonLegit_Click(object sender, RibbonControlEventArgs e) 
        {
            ProcessCurrentItem(Classification.legit);
        }

        protected override void buttonSpam_Click(object sender, RibbonControlEventArgs e)
        {
            ProcessCurrentItem(Classification.spam);
        }

        protected override void buttonSimulation_Click(object sender, RibbonControlEventArgs e) 
        {
            ProcessCurrentItem(Classification.simulation);
        }

        protected override void buttonMalicious_Click(object sender, RibbonControlEventArgs e) 
        {
            ProcessCurrentItem(Classification.malicious);
        }

    }
}