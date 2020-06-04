namespace OutlookSecurityAddInReport
{
    partial class RibbonSecurityReport : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonSecurityReport()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            this.groupHelp = this.Factory.CreateRibbonGroup();
            this.buttonReport = this.Factory.CreateRibbonButton();
            this.tabSecurity.SuspendLayout();
            this.groupHelp.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabSecurity
            // 
            this.tabSecurity.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabSecurity.Groups.Add(this.groupHelp);
            this.tabSecurity.Label = "Security";
            this.tabSecurity.Name = "tabSecurity";
            // 
            // groupHelp
            // 
            this.groupHelp.Items.Add(this.buttonReport);
            this.groupHelp.Label = "Help";
            this.groupHelp.Name = "groupHelp";
            // 
            // buttonReport
            // 
            this.buttonReport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonReport.Image = global::OutlookSecurityAddInReport.Properties.Resources.shield;
            this.buttonReport.Label = "Report email";
            this.buttonReport.Name = "buttonReport";
            this.buttonReport.ShowImage = true;
            this.buttonReport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonReport_Click);
            // 
            // RibbonSecurityReport
            // 
            this.Name = "RibbonSecurityReport";
            this.RibbonType = "";
            this.Tabs.Add(this.tabSecurity);
            this.tabSecurity.ResumeLayout(false);
            this.tabSecurity.PerformLayout();
            this.groupHelp.ResumeLayout(false);
            this.groupHelp.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonReport;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabSecurity;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonSecurityReport RibbonSecurity
        {
            get { return this.GetRibbon<RibbonSecurityReport>(); }
        }
    }
}
