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

        public List<string> GetLatestEmailSubjects(string folderName, int count)
        {
            var subjects = new List<string>();

            foreach (Outlook.MAPIFolder rootFolder in _ns.Folders)
            {
                foreach (Outlook.MAPIFolder subFolder in rootFolder.Folders)
                {
                    if (subFolder.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase))
                    {
                        // Pobieramy elementy i sortujemy je od najnowszych
                        Outlook.Items items = subFolder.Items;
                        items.Sort("[ReceivedTime]", true); // true = malejąco

                        int totalItems = items.Count;
                        // Wyciągamy tematy (zabezpieczając się przed małą liczbą maili)
                        for (int i = 1; i <= Math.Min(count, totalItems); i++)
                        {
                            // Sprawdzamy, czy to na pewno jest mail (a nie np. zaproszenie)
                            if (items[i] is Outlook.MailItem mail)
                            {
                                string dateInfo = mail.ReceivedTime.ToString("g");
                                subjects.Add($"[{dateInfo}] {mail.Subject}");
                            }
                        }
                        return subjects;
                    }
                }
            }
            return subjects;
        }

        public void Logout()
        {
            _ns.Logoff();
        }
    }
}