using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using EmailCounter.Shared.Models;

namespace EmailCounter.Shared.Services
{
    public class OutlookService
    {
        private dynamic? _outlookApp;
        private dynamic? _ns;

        public OutlookService()
        {
            try 
            {
                Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType != null)
                {
                    _outlookApp = Activator.CreateInstance(outlookType);
                    _ns = _outlookApp!.GetNamespace("MAPI");
                    _ns!.Logon("", "", false, false);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Nie udało się połączyć z Outlookiem: {ex.Message}");
            }
        }
        
        public List<OutlookFolder> GetFolders()
        {
            var folders = new List<OutlookFolder>();
            if (_ns == null) return folders;

            try {
                foreach (dynamic root in _ns!.Folders)
                {
                    folders.Add(MapFolder(root));
                }
            } catch { /* Ignoruj błędy pojedynczych folderów */ }
            
            return folders;
        }

        public List<EmailData> GetEmailsForExport(string folderName, DateTime startDate, DateTime endDate)
        {
            var results = new List<EmailData>();
            string filter =$"[ReceivedTime]>='{startDate:g}' AND [ReceivedTime]<='{endDate:g}'";

            foreach (Outlook.MAPIFolder rootFolder in _ns!.Folders)
            {
                foreach (Outlook.MAPIFolder subFolder in rootFolder.Folders)
                {
                    if (subFolder.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase))
                    {
                        Outlook.Items items = subFolder.Items.Restrict(filter);
                        foreach (object item in items)
                        {
                            if (item is Outlook.MailItem mail)
                            {
                                results.Add(new EmailData
                                {
                                    Subject = mail.Subject,
                                    ReceivedTime = mail.ReceivedTime,
                                    Sender = mail.SenderName,
                                    ConversationID = mail.ConversationID,
                                    ConversationTopic = mail.ConversationTopic
                                });
                            }
                        }
                        return results;
                    }
                }
            }
            return results;
        }
        
        public int GetMessageCount(string folderName, DateTime? startDate=null, DateTime? endDate=null)
        {
            foreach (Outlook.MAPIFolder rootFolder in _ns!.Folders)
            {
                foreach (Outlook.MAPIFolder subFolder in rootFolder.Folders)
                {
                    if (subFolder.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (startDate.HasValue && endDate.HasValue)
                        {
                            string filter =$"[ReceivedTime]>='{startDate:g}' AND [ReceivedTime]<='{endDate:g}'";
                            Outlook.Items filteredItems =subFolder.Items.Restrict(filter);
                            return filteredItems.Count;
                        }
                        return subFolder.Items.Count;
                    }
                }
            }
            return -1;
        }

        private OutlookFolder MapFolder(dynamic outlookFolder)
        {
            var folder = new OutlookFolder {
                FolderName = outlookFolder?.Name ?? "Nieznany",
                FullPath = outlookFolder?.FolderPath ?? "Nieznany",
                SubFolders = new List<OutlookFolder>()
            };

            try {
                foreach (dynamic sub in outlookFolder!.Folders)
                {
                    folder.SubFolders.Add(MapFolder(sub));
                }
            } catch { }

            return folder;
        }
        public void Logout()
        {
            _ns?.Logoff();
        }
    }
}