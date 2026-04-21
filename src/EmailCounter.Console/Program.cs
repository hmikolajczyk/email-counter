using System;

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