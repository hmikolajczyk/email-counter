using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;
using EmailCounter.Shared.Models;
using EmailCounter.Shared.Services;

namespace EmailCounter.Tests
{
    public class CsvExportServiceTests
    {
        [Fact]
        public void Export_ValidListOfEmails_GeneratesFileWithCorrectHeaderRow()
        {
            var service = new CsvExportService();
            var emails = new List<EmailData>();
            
            string tempFilePath = Path.GetTempFileName(); 
            
            string expectedHeader = "Temat;Data Otrzymania;Nadawca;ID Konwersacji;Temat Konwersacji";

            service.ExportEmails(emails, tempFilePath);

            string ?actualHeader = File.ReadLines(tempFilePath).FirstOrDefault();
            
            Assert.Equal(expectedHeader, actualHeader);

            if (File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }
            
        }

        [Fact]
        public void Export_EmptyList_GeneratesFileWithOnlyHeaderRow()
        {
            var service = new CsvExportService();
            var emails = new List<EmailData>();
            
            string tempFilePath = Path.GetTempFileName(); 
            
            service.ExportEmails(emails, tempFilePath);

            int lineCount = File.ReadLines(tempFilePath).Count();

            Assert.Equal(1, lineCount);

            if (File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }
        }

        [Fact]
        public void Export_PolishCharactersEncoding_ShouldPreservePolishCharacters()
        {
            var service = new CsvExportService();
            var emails = new List<EmailData>
            {
                new EmailData
                {
                    Subject = "Chrząszcz brzmi w trzcinie",
                    ReceivedTime = DateTime.Now,
                    Sender = "Grzegorz Brzęczyszczykiewicz",
                    ConversationID = "3423214231",
                    ConversationTopic = "ęóąśłżźń"
                }
            };

            string tempFilePath = Path.GetTempFileName(); 

            service.ExportEmails(emails, tempFilePath);

            string ?dataRow = File.ReadLines(tempFilePath).Skip(1).FirstOrDefault();

            Assert.Contains("Chrząszcz brzmi w trzcinie", dataRow);
            Assert.Contains("Grzegorz Brzęczyszczykiewicz", dataRow);
            Assert.Contains("ęóąśłżźń", dataRow);

            if (File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }

        }

        [Fact]
        public void Export_NullFieldsInEmail_ShouldNotThrowExceptions()
        {
            var service = new CsvExportService();
            var emails = new List<EmailData>
            {
                new EmailData
                {
                    Subject = null,
                    ReceivedTime = DateTime.Now,
                    Sender = "",
                    ConversationID = "",
                    ConversationTopic = ""
                }
            };

            string tempFilePath = Path.GetTempFileName();

            var exception = Record.Exception(() => service.ExportEmails(emails, tempFilePath));

            Assert.Null(exception);

            if (File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }
        }
    }
}