using System;
using EmailCounter.Shared;

namespace EmailCounter.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var service = new OutlookService();
            
            string folderName;
            do
            {   Console.Write("Podaj nazwę folderu (aby wyjść wpisz 'exit'): ");
                folderName = Console.ReadLine() ?? "";

                int count = service.GetMessageCount(folderName);
                if (count != -1)
                {
                    DateTime startDate;
                    bool isDateValid=false;
                    do
                    {
                        Console.Write("Podaj datę początkową (RRRR-MM-DD): ");
                        string startDateString=Console.ReadLine()??"";
                        isDateValid=DateTime.TryParse(startDateString,out startDate);
                        if (!isDateValid)
                        {  
                            Console.WriteLine("BŁĄD: Niepoprawna data.");
                        }
                    }while(!isDateValid);

                    DateTime endDate;
                    isDateValid=false;
                    do
                    {
                        Console.Write("Podaj datę końcową (RRRR-MM-DD): ");
                        string endDateString=Console.ReadLine()??"";
                        isDateValid=DateTime.TryParse(endDateString,out endDate);
                        if (!isDateValid)
                        {  
                            Console.WriteLine("BŁĄD: Niepoprawna data.");
                        }
                    }while(!isDateValid);
                    endDate=endDate.Date.AddDays(1).AddTicks(-1);

                    count=service.GetMessageCount(folderName,startDate,endDate);

                    Console.WriteLine($"Liczba maili w folderze '{folderName}': {count}");

                    string generateFiles="?";
                    do
                    {
                        Console.Write($"Generować '{folderName}.csv' (t/n): ");
                        generateFiles=Console.ReadLine()??"";
                        generateFiles=generateFiles.ToLower();
                        if (generateFiles == "t")
                        {
                            List<EmailData> emailsToExport = service.GetEmailsForExport(folderName, startDate, endDate);

                            if (emailsToExport.Count > 0)
                            {
                                var csvService = new CsvExportService();
                                string pwd = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                                string fileName = folderName+".csv";
                                string fullPath = Path.Combine(pwd, fileName);

                                csvService.ExportEmails(emailsToExport, fullPath);
                                
                                Console.WriteLine($"Sukces! Raport zapisany na pulpicie: {fullPath}");
                            }
                            else
                            {
                                Console.WriteLine("Brak maili do eksportu w podanym zakresie.");
                            }
                        } else if (generateFiles != "n")
                        {
                            Console.WriteLine("BŁĄD: Zła opcja.");
                        }
                    } while(generateFiles!="t"&generateFiles!="n");
                    break;
                }
                else if(folderName!="exit")
                {
                    Console.WriteLine("BŁĄD: Nie znaleziono takiego folderu.");
                }
            } while (folderName!="exit");

            service.Logout();
        }
    }
}