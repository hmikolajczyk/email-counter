using System.Collections.Generic;
using System.IO;
using System.Text;
using EmailCounter.Shared.Models;

namespace EmailCounter.Shared.Services
{
    public class CsvExportService
    {
        public void ExportEmails(List<EmailData> emails, string filePath)
        {
            using (var writer = new StreamWriter(filePath, false, System.Text.Encoding.UTF8))
            {
                writer.WriteLine("Temat;Data Otrzymania;Nadawca;ID Konwersacji;Temat Konwersacji");

                foreach (var email in emails)
                {
                    string row = $"{(email.Subject ?? "").Replace(";", " ")};"+
                    $"{email.ReceivedTime};" +
                    $"{email.Sender.Replace(";", " ")};" +
                    $"{email.ConversationID.Replace(";", " ")};" +
                    $"{email.ConversationTopic.Replace(";", " ")};";
                    writer.WriteLine(row);
                }
            }
        }
    }
}