using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using SharedResources;
using System.IO;
using CsvHelper;
using System.Globalization;
using System.Windows.Forms;

namespace ExcelSecurityAddInAnalysis
{
    public partial class RibbonSecurity
    {
        protected FormProgressManager Manager { get; } = new FormProgressManager();

        public class CSVTemplate
        {
            public string FullFolderPath { get; set; }
            public string Subject { get; set; }
            public string ReceivedTime { get; set; }
            public string LastModificationTime { get; set; }
            public bool IsRead { get; set; }
            public bool IsReply { get; set; }
            public string Replied { get; set; }
            public string RepliedTime { get; set; }
        }

        private bool IsPathValid(string path)
        {
            try
            {
                return Directory.Exists(Path.GetDirectoryName(path));
            }
            catch
            {
                return false;
            }
        }

        private List<Outlook.Folder> RecurseFolder(Outlook.Folder origin)
        {
            var folders = new List<Outlook.Folder>();
            folders.Add(origin);
            foreach (Outlook.Folder folder in origin.Folders)
            {
                folders.AddRange(RecurseFolder(folder));
            }
            return folders;
        }

        private class FormProgressArgs
        {
            public string Path { get; set; }
            public string Mailbox { get; set; }
            public enum Types
            {
                table,
                csv
            }

            public Types Type { get; set; }

        }


        private FormProgress FormProgress { get; set; } = null;

        private void Worker_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();

            var worker = sender as System.ComponentModel.BackgroundWorker;
            var arguments = (FormProgressArgs)e.Argument;

            var path = arguments.Path;
            var mailbox = arguments.Mailbox;
            var type = arguments.Type;

            Debug.WriteLine(string.Format("path: {0}",path));
            Debug.WriteLine(string.Format("mailbox: {0}", mailbox));

            Manager.CurrentStage = FormProgressManager.ProcessStage.Prepare;

            Outlook.Application app = new Outlook.Application();

            var rootFolder = OutlookFunctions.GetRootFolder(app, mailbox);

            if (rootFolder == null)
                throw new Exception("root folder not found");

            var folders = RecurseFolder(rootFolder);

            var mailsList = new List<Outlook.MailItem>();
            
            int itemcount = 0;
            int currentitem = 0;
            foreach (Outlook.Folder folder in folders)
            {
                itemcount += folder.Items.Count;

                foreach (dynamic item in folder.Items)
                {
                    currentitem++;

                    Manager.SetStatus(currentitem, itemcount);

                    if (worker.CancellationPending)
                    {
                        e.Cancel = true;
                        return;
                    }

                    if (item.Class == (int)Outlook.OlObjectClass.olMail)
                    {
                        mailsList.Add(item);
                    }
                }
            }

            Manager.CurrentStage = FormProgressManager.ProcessStage.Process;
            itemcount = mailsList.Count;
            currentitem = 0;
            switch (type)
            {
                case FormProgressArgs.Types.csv:

                    using (var writer = new StreamWriter(path))
                    using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                    {
                        csv.WriteHeader<CSVTemplate>();
                        csv.NextRecord();

                        foreach (var mailItem in mailsList)
                        {
                            currentitem++;

                            Manager.SetStatus(currentitem, itemcount);

                            if (worker.CancellationPending)
                            {
                                e.Cancel = true;
                                return;
                            }

                            try
                            {
                                Outlook.Folder parentFolder = mailItem.Parent;

                                var entry = new CSVTemplate
                                {
                                    FullFolderPath = parentFolder.FullFolderPath,
                                    Subject = mailItem.Subject,
                                    ReceivedTime = mailItem.ReceivedTime.ToString(),
                                    LastModificationTime = mailItem.LastModificationTime.ToString(),
                                    IsRead = !mailItem.UnRead,
                                    IsReply = MailItemProperties.GetIsReply(mailItem),
                                    Replied = MailItemProperties.GetReplied(mailItem),
                                    RepliedTime = MailItemProperties.GetRepliedTime(mailItem)
                                };

                                csv.WriteRecord(entry);
                                csv.NextRecord();
                            }

                            catch (Exception exc)
                            {
                                Debug.WriteLine(exc.Message);
                            }
                        }
                    }
                    break;

                case FormProgressArgs.Types.table:

                    Excel.Worksheet data_worksheet = null;
                    foreach (Excel.Worksheet sheet in ExcelApp.ActiveWorkbook.Worksheets)
                    {
                        if (sheet.Name == "mailbox")
                        {
                            data_worksheet = sheet;
                            break;
                        }
                    }

                    if (data_worksheet == null)
                    {
                        data_worksheet = ExcelApp.ActiveWorkbook.Worksheets.Add();
                        data_worksheet.Name = "mailbox";
                    }

                    if (data_worksheet == null)
                        return;

                    foreach (Excel.ListObject listObject in data_worksheet.ListObjects)
                    {
                        if (listObject.Name == "Mail")
                        {
                            listObject.Delete();
                            break;
                        }
                    }

                    Excel.ListObject data_table = data_worksheet.ListObjects.Add(SourceType: Excel.XlListObjectSourceType.xlSrcRange);
                    data_table.Name = "Mail";
                    data_table.ListColumns[1].Name = "FullFolderPath";
                    data_table.ListColumns.Add().Name = "Subject";
                    data_table.ListColumns.Add().Name = "ReceivedTime";
                    data_table.ListColumns.Add().Name = "LastModificationTime";
                    data_table.ListColumns.Add().Name = "IsRead";
                    data_table.ListColumns.Add().Name = "IsReply";
                    data_table.ListColumns.Add().Name = "Replied";
                    data_table.ListColumns.Add().Name = "RepliedTime";

                    
                    data_table.Resize(data_worksheet.Range[data_table.Range.Cells[1, 1],
                                                          data_table.Range.Cells[itemcount + 1, data_table.Range.Columns.Count]
                                                          ]);


                    
                    foreach (Outlook.MailItem mailItem in mailsList)
                    {
                        currentitem++;

                        Manager.SetStatus(currentitem, itemcount);

                        if (worker.CancellationPending)
                        {
                            e.Cancel = true;
                            return;
                        }

                        try
                        {
                            var row = data_table.ListRows[currentitem];

                            Outlook.Folder parentFolder = mailItem.Parent;
                            row.Range.Columns[1].Value = parentFolder.FullFolderPath;
                            row.Range.Columns[2].Value = mailItem.Subject;
                            row.Range.Columns[3].Value = mailItem.ReceivedTime;
                            row.Range.Columns[4].Value = mailItem.LastModificationTime;
                            row.Range.Columns[5].Value = !mailItem.UnRead;
                            row.Range.Columns[6].Value = MailItemProperties.GetIsReply(mailItem);
                            row.Range.Columns[7].Value = MailItemProperties.GetReplied(mailItem);
                            row.Range.Columns[8].Value = MailItemProperties.GetRepliedTime(mailItem);
                        }
                        catch (Exception exc)
                        {
                            Debug.WriteLine(exc.Message);
                        }
                        
                    }
                    data_worksheet.Columns.AutoFit();

                    break;
            }

            Manager.SetStatus(itemcount, itemcount);

            watch.Stop();

            Debug.WriteLine(watch.ElapsedMilliseconds);
        }

    }
}
