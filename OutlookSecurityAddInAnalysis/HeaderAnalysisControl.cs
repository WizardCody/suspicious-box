using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using Microsoft.Office.Tools;
using System.Text.RegularExpressions;
using SharedResources;
using System.IO;
using System.Runtime.InteropServices;


namespace OutlookSecurityAddInAnalysis
{
    public partial class HeaderAnalysisControl : UserControl
    {
        private Outlook.MailItem _currentItem = null;
        public Outlook.MailItem CurrentItem
        {
            get
            {
                return _currentItem;
            }
            set
            {
                _currentItem = value;
                ApplyHeader();
            }
        }

        public void ApplyHeader()
        {
            bool found = false;
            if (CheckBox.Checked)
            {
                foreach (Outlook.Attachment attachment in CurrentItem.Attachments)
                {
                    if (attachment.Type == Outlook.OlAttachmentType.olEmbeddeditem)
                    {
                        string path = string.Format(@"C:\Users\Marcin\Desktop\{0}", "test.msg");
                        
                        attachment.SaveAsFile(path);

                        var item = Globals.ThisAddIn.Application.Session.OpenSharedItem(path) as Outlook.MailItem;

                        if (item.Class == Outlook.OlObjectClass.olMail)
                        {
                            var prop = item.PropertyAccessor;
                            TextBox.Text = prop.GetProperty(MailItemProperties.Schemas.PidTagTransportMessageHeaders());

                            Marshal.FinalReleaseComObject(prop);
                        }

                        Marshal.FinalReleaseComObject(item);

                        found = true;
                        break;
                    }
                }
            }

            if (!CheckBox.Checked || !found)
            {
                TextBox.Text = CurrentItem.PropertyAccessor.GetProperty(MailItemProperties.Schemas.PidTagTransportMessageHeaders());
            }
        }

        public RichTextBox TextBox
        {
            get
            {
                return richTextBox1;
            }
        }

        public CheckBox CheckBox
        {
            get
            {
                return checkBox1;
            }
        }

        public HeaderAnalysisControl()
        {
            InitializeComponent();
        }

        private Regex MatchHeader(string header)
        {
            return new Regex(string.Format(@"^{0}:(.|\n)*?(?=\n\w)", header), RegexOptions.Multiline | RegexOptions.ECMAScript | RegexOptions.IgnoreCase);
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            foreach (var header in Properties.Settings.Default.HeaderWhitelist)
            {
                Regex regex = MatchHeader(header);
                MatchCollection matches = regex.Matches(TextBox.Text);

                if (matches.Count > 0)
                {
                    foreach (Match m in matches)
                    {
                        TextBox.Select(m.Index, m.Length);
                        TextBox.SelectionBackColor = Color.LightGreen;
                    }
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            ApplyHeader();
        }
    }
}
