namespace OutlookSecurityAddInAnalysis
{
    partial class RibbonSecurityAnalysis : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabSecurity = this.Factory.CreateRibbonTab();
            this.groupAnalysis = this.Factory.CreateRibbonGroup();
            this.buttonSample = this.Factory.CreateRibbonButton();
            this.buttonLegit = this.Factory.CreateRibbonButton();
            this.buttonSimulation = this.Factory.CreateRibbonButton();
            this.buttonSpam = this.Factory.CreateRibbonButton();
            this.buttonMalicious = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.toggleButtonShowHeaders = this.Factory.CreateRibbonToggleButton();
            this.buttonAnalyse = this.Factory.CreateRibbonButton();
            this.tabSecurity.SuspendLayout();
            this.groupAnalysis.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabSecurity
            // 
            this.tabSecurity.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabSecurity.Groups.Add(this.groupAnalysis);
            this.tabSecurity.Groups.Add(this.group1);
            this.tabSecurity.Label = "Security";
            this.tabSecurity.Name = "tabSecurity";
            // 
            // groupAnalysis
            // 
            this.groupAnalysis.Items.Add(this.buttonSample);
            this.groupAnalysis.Items.Add(this.buttonLegit);
            this.groupAnalysis.Items.Add(this.buttonSimulation);
            this.groupAnalysis.Items.Add(this.buttonSpam);
            this.groupAnalysis.Items.Add(this.buttonMalicious);
            this.groupAnalysis.Label = "Response";
            this.groupAnalysis.Name = "groupAnalysis";
            // 
            // buttonSample
            // 
            this.buttonSample.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonSample.Image = global::OutlookSecurityAddInAnalysis.Properties.Resources.mail;
            this.buttonSample.Label = "No sample";
            this.buttonSample.Name = "buttonSample";
            this.buttonSample.ShowImage = true;
            this.buttonSample.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSample_Click);
            // 
            // buttonLegit
            // 
            this.buttonLegit.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonLegit.Image = global::OutlookSecurityAddInAnalysis.Properties.Resources.mail;
            this.buttonLegit.Label = "Legitimate";
            this.buttonLegit.Name = "buttonLegit";
            this.buttonLegit.ShowImage = true;
            this.buttonLegit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonLegit_Click);
            // 
            // buttonSimulation
            // 
            this.buttonSimulation.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonSimulation.Image = global::OutlookSecurityAddInAnalysis.Properties.Resources.mail;
            this.buttonSimulation.Label = "Simulation";
            this.buttonSimulation.Name = "buttonSimulation";
            this.buttonSimulation.ShowImage = true;
            this.buttonSimulation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSimulation_Click);
            // 
            // buttonSpam
            // 
            this.buttonSpam.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonSpam.Image = global::OutlookSecurityAddInAnalysis.Properties.Resources.mail;
            this.buttonSpam.Label = "Spam";
            this.buttonSpam.Name = "buttonSpam";
            this.buttonSpam.ShowImage = true;
            this.buttonSpam.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSpam_Click);
            // 
            // buttonMalicious
            // 
            this.buttonMalicious.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonMalicious.Image = global::OutlookSecurityAddInAnalysis.Properties.Resources.mail2;
            this.buttonMalicious.Label = "Malicious";
            this.buttonMalicious.Name = "buttonMalicious";
            this.buttonMalicious.ShowImage = true;
            this.buttonMalicious.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonMalicious_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.toggleButtonShowHeaders);
            this.group1.Items.Add(this.buttonAnalyse);
            this.group1.Label = "Analysis";
            this.group1.Name = "group1";
            // 
            // toggleButtonShowHeaders
            // 
            this.toggleButtonShowHeaders.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleButtonShowHeaders.Image = global::OutlookSecurityAddInAnalysis.Properties.Resources.zoom;
            this.toggleButtonShowHeaders.Label = "Show headers";
            this.toggleButtonShowHeaders.Name = "toggleButtonShowHeaders";
            this.toggleButtonShowHeaders.ShowImage = true;
            this.toggleButtonShowHeaders.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonShowHeaders_Click);
            // 
            // buttonAnalyse
            // 
            this.buttonAnalyse.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAnalyse.Image = global::OutlookSecurityAddInAnalysis.Properties.Resources.vision;
            this.buttonAnalyse.Label = "Analyse";
            this.buttonAnalyse.Name = "buttonAnalyse";
            this.buttonAnalyse.ShowImage = true;
            // 
            // RibbonSecurityAnalysis
            // 
            this.Name = "RibbonSecurityAnalysis";
            this.RibbonType = "";
            this.Tabs.Add(this.tabSecurity);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonSecurityAnalysis_Load);
            this.tabSecurity.ResumeLayout(false);
            this.tabSecurity.PerformLayout();
            this.groupAnalysis.ResumeLayout(false);
            this.groupAnalysis.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAnalysis;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSample;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonLegit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSpam;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonMalicious;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSimulation;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabSecurity;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonShowHeaders;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAnalyse;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonSecurityAnalysis RibbonSecurityAnalysis
        {
            get { return this.GetRibbon<RibbonSecurityAnalysis>(); }
        }
    }
}
