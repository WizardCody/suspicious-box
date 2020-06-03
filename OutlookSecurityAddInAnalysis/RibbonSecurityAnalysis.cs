using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using Microsoft.Office.Tools;
using System.Reflection;
using SharedResources;

namespace SecurityAddInAnalysis
{

    public partial class RibbonSecurityAnalysis
    {

        public string TaskPaneTitle { get; } = "Message headers";

        public RibbonSecurityAnalysis()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        protected virtual void RibbonSecurityAnalysis_Load(object sender, RibbonUIEventArgs e)
        {
            foreach (var TaskPane in ThisAddIn.CustomTaskPanes)
            {
                if (TaskPane.Title == TaskPaneTitle & TaskPane.Window == App.ActiveWindow())
                {
                    TaskPaneHeaderAnalysis = TaskPane;
                    toggleButtonShowHeaders.Checked = TaskPane.Visible;
                    break;
                }
            }
        }

        public static Outlook.Application App
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

        public HeaderAnalysisControl HeaderAnalysisControl { get; private set; } = null;

        // auto property
        public CustomTaskPane TaskPaneHeaderAnalysis { get; private set; } = null;


        protected virtual CustomTaskPane SetupTaskPaneHeaderAnalysis()
        {
            HeaderAnalysisControl = new HeaderAnalysisControl();
            HeaderAnalysisControl.Load += HeaderAnalysisControl_Load;
            return ThisAddIn.CustomTaskPanes.Add(HeaderAnalysisControl, TaskPaneTitle);
        }

        protected virtual void HeaderAnalysisControl_Load(object sender, EventArgs e) { }



        protected void ProcessMail(Outlook.MailItem item, Classification type)
        {
            var account = OutlookFunctions.GetAccount(App, Properties.Settings.Default.SendFromMailbox);

            if (account == null)
                return;

            // Przerwij jeżeli napotkasz błąd na jakimś etapie.
            try
            {
                Outlook.MailItem template = App.CreateItemFromTemplate(type.GetTemplate());

                var newMail = item.Reply();
                var inspector = newMail.GetInspector;

                newMail.HTMLBody = template.HTMLBody + newMail.HTMLBody;

                newMail.Sender = account.CurrentUser.AddressEntry;
                newMail.SendUsingAccount = account;

                newMail.Send();

                item.Categories = type.ToString();

                // root folder
                Outlook.MAPIFolder moveFolder = OutlookFunctions.GetRootFolder(App, Properties.Settings.Default.RequestMailbox);

                moveFolder = moveFolder.Folders[Properties.Settings.Default.RequestMailboxDoneFolder].Folders[type.ToString()];

                item.Move(moveFolder);
            }
            catch (Exception exc)
            {
                Debug.WriteLine(exc.Message);
                return;
            }            
        }

        protected virtual void buttonSample_Click(object sender, RibbonControlEventArgs e) { }

        protected virtual void buttonLegit_Click(object sender, RibbonControlEventArgs e) { }

        protected virtual void buttonSpam_Click(object sender, RibbonControlEventArgs e) { }

        protected virtual void buttonSimulation_Click(object sender, RibbonControlEventArgs e) { }

        protected virtual void buttonMalicious_Click(object sender, RibbonControlEventArgs e) { }

        protected virtual void toggleButtonShowHeaders_Click(object sender, RibbonControlEventArgs e)
        {
            if (TaskPaneHeaderAnalysis == null)
                TaskPaneHeaderAnalysis = SetupTaskPaneHeaderAnalysis();

            if (TaskPaneHeaderAnalysis != null)
                TaskPaneHeaderAnalysis.Visible = toggleButtonShowHeaders.Checked;
        }
    }

    public static class ClassificationInfo
    { 

        public static string GetTemplate(this Classification type)
        {
            switch (type)
            {
                case Classification.noSample:
                    return Properties.Settings.Default.NoAttachmentTemplate;
                case Classification.legit:
                    return Properties.Settings.Default.LegitTemplate;
                case Classification.spam:
                    return Properties.Settings.Default.SpamTemplate;
                case Classification.malicious:
                    return Properties.Settings.Default.MaliciousTemplate;
                case Classification.simulation:
                    return Properties.Settings.Default.SimulationTemplate;
                default: 
                    return string.Empty;
            }
        }
    }

    public enum Classification
    {
        noSample,
        legit,
        spam,
        malicious,
        simulation
    }

}
