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
    }
}