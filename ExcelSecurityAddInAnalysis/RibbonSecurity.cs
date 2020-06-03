using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelAddIn
{
    public partial class RibbonSecurity
    {
        public RibbonEditBox MailboxEdit
        {
            get
            {
                return editBoxMailbox;
            }
        }

        public static Excel.Application ExcelApp
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
    }
}
