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

namespace OutlookSecurityAddInAnalysis
{
    public partial class HeaderAnalysisControl : UserControl
    {

        public RichTextBox TextBox
        {
            get
            {
                return richTextBox1;
            }
        }

        public HeaderAnalysisControl()
        {
            InitializeComponent();
        }

    }
}
