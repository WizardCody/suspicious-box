using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

namespace SharedResources
{
    public class OutlookFunctions
    {
        public static Outlook.Account GetAccount(Outlook.Application app, string smtpAddress)
        {
            foreach (Outlook.Account account in app.Session.Accounts)
            {
                if (account.SmtpAddress.ToUpper() == smtpAddress.ToUpper())
                {
                    return account;
                }
            }
            return null;
        }

        public static Outlook.Folder GetRootFolder(Outlook.Application app, string smtpAddress)
        {
            foreach (Outlook.Folder folder in app.Session.Folders)
            {
                if (folder.Name.ToUpper() == smtpAddress.ToUpper())
                {
                    return folder;
                }
            }
            return null;
        }

        public static string ValidFilename(string filename)
        {
            return string.Join("_", filename.Split(Path.GetInvalidFileNameChars()));
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool AllocConsole();

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool FreeConsole();

    }


}
