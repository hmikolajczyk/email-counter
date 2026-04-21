using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EmailCounter.ConsoleApp
{
    public class OutlookService
    {
        private Outlook.Application _outlookApp;
        private Outlook.NameSpace _ns;

        public OutlookService()
        {
            _outlookApp = new Outlook.Application();
            _ns = _outlookApp.GetNamespace("MAPI");
            _ns.Logon(Type.Missing, Type.Missing, false, false);
        }

        public int GetMessageCount(string folderName, DateTime? startDate=null, DateTime? endDate=null)
        {
            foreach (Outlook.MAPIFolder rootFolder in _ns.Folders)
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

        public List<EmailData> GetEmailsForExport(string folderName, DateTime startDate, DateTime endDate)
        {
            var results = new List<EmailData>();
            string filter =$"[ReceivedTime]>='{startDate:g}' AND [ReceivedTime]<='{endDate:g}'";

            foreach (Outlook.MAPIFolder rootFolder in _ns.Folders)
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
        public void Logout()
        {
            _ns.Logoff();
        }
    }
}