namespace ExcelAddIn
{
    partial class RibbonSecurity : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonSecurity()
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
            this.groupData = this.Factory.CreateRibbonGroup();
            this.menu1 = this.Factory.CreateRibbonMenu();
            this.buttonGenerate1 = this.Factory.CreateRibbonButton();
            this.buttonGenerate2 = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.box1 = this.Factory.CreateRibbonBox();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.editBoxMailbox = this.Factory.CreateRibbonEditBox();
            this.tabSecurity.SuspendLayout();
            this.groupData.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabSecurity
            // 
            this.tabSecurity.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabSecurity.Groups.Add(this.groupData);
            this.tabSecurity.Label = "Security";
            this.tabSecurity.Name = "tabSecurity";
            // 
            // groupData
            // 
            this.groupData.Items.Add(this.menu1);
            this.groupData.Items.Add(this.separator1);
            this.groupData.Items.Add(this.box1);
            this.groupData.Label = "Data";
            this.groupData.Name = "groupData";
            // 
            // menu1
            // 
            this.menu1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu1.Image = global::ExcelAddIn.Properties.Resources.file;
            this.menu1.Items.Add(this.buttonGenerate1);
            this.menu1.Items.Add(this.buttonGenerate2);
            this.menu1.Label = "Get mailbox data into...";
            this.menu1.Name = "menu1";
            this.menu1.ShowImage = true;
            // 
            // buttonGenerate1
            // 
            this.buttonGenerate1.Image = global::ExcelAddIn.Properties.Resources.file;
            this.buttonGenerate1.Label = "table";
            this.buttonGenerate1.Name = "buttonGenerate1";
            this.buttonGenerate1.ShowImage = true;
            this.buttonGenerate1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGenerate1_Click);
            // 
            // buttonGenerate2
            // 
            this.buttonGenerate2.Image = global::ExcelAddIn.Properties.Resources.file;
            this.buttonGenerate2.Label = "csv file";
            this.buttonGenerate2.Name = "buttonGenerate2";
            this.buttonGenerate2.ShowImage = true;
            this.buttonGenerate2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGenerate2_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.label1);
            this.box1.Items.Add(this.editBoxMailbox);
            this.box1.Name = "box1";
            // 
            // label1
            // 
            this.label1.Label = "mailbox";
            this.label1.Name = "label1";
            // 
            // editBoxMailbox
            // 
            this.editBoxMailbox.Label = "editBox1";
            this.editBoxMailbox.Name = "editBoxMailbox";
            this.editBoxMailbox.ShowLabel = false;
            this.editBoxMailbox.SizeString = "_____________________________";
            this.editBoxMailbox.Text = null;
            // 
            // RibbonSecurity
            // 
            this.Name = "RibbonSecurity";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabSecurity);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonSecurity_Load);
            this.tabSecurity.ResumeLayout(false);
            this.tabSecurity.PerformLayout();
            this.groupData.ResumeLayout(false);
            this.groupData.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabSecurity;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGenerate1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGenerate2;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxMailbox;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonSecurity Ribbon1
        {
            get { return this.GetRibbon<RibbonSecurity>(); }
        }
    }
}
