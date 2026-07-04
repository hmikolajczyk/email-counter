using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;
using EmailCounter.Shared;

namespace EmailCounter.Tests
{
    public class CsvExportServiceTests
    {
        [Fact]
        public void Export_ValidListOfEmails_GeneratesFileWithCorrectHeaderRow()
        {
            // 1. Arrange (Przygotowanie)
            // TUTAJ: Stwórz obiekt serwisu, pustą listę EmailData oraz pobierz Path.GetTempFileName()
            // Zdefiniuj też string z oczekiwanym nagłówkiem.


            // 2. Act (Działanie)
            // TUTAJ: Wywołaj metodę ExportEmails z przygotowanymi danymi


            // 3. Assert (Sprawdzenie)
            // TUTAJ: Przeczytaj pierwszą linię pliku za pomocą File.ReadLines(...).FirstOrDefault()
            // Porównaj ją z oczekiwanym nagłówkiem za pomocą Assert.Equal


            // 4. Clean (Sprzątanie)
            // TUTAJ: Skasuj plik tymczasowy za pomocą File.Delete
            
        }
    }
}