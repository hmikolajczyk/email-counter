using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using EmailCounter.Shared.Models;
using EmailCounter.Shared.Services;
using ReactiveUI;
using System.IO;
using System.Threading.Tasks;
using Microsoft.VisualBasic;

namespace EmailCounter.Gui.ViewModels;
public class MainWindowViewModel : ViewModelBase
{
    private readonly OutlookService _outlookService = new();
    private readonly CsvExportService _csvService = new();
    private DateTimeOffset _startDate;
    public DateTimeOffset StartDate
    { 
        get => _startDate; 
        set {
            this.RaiseAndSetIfChanged(ref _startDate, value);
            ResetStatus();
        }
    }
    private DateTimeOffset _endDate;
    public DateTimeOffset EndDate 
    { 
        get => _endDate; 
        set {
            this.RaiseAndSetIfChanged(ref _endDate, value);
            ResetStatus();
        }
    }
    private string _statusMessage = "";
    public string StatusMessage 
    { 
        get => _statusMessage; 
        set => this.RaiseAndSetIfChanged(ref _statusMessage, value); 
    }
    private OutlookFolder? _selectedFolder;
    public OutlookFolder? SelectedFolder 
    { 
        get => _selectedFolder;
        set
        {
            this.RaiseAndSetIfChanged(ref _selectedFolder, value);
            ResetStatus();
            //if (value != null) StatusMessage = $"Wybrano: {value.FolderName}";
        }
    }
    private string _statusColor = "Gray";
    public string StatusColor 
    { 
        get => _statusColor; 
        set => this.RaiseAndSetIfChanged(ref _statusColor, value); 
    }
    public async void GenerateReportCommand()
    {
        StatusMessage = "Przetwarzanie...";
        StatusColor = "Gray";
        await Task.Delay(100);
        if (SelectedFolder == null)
        {
            StatusMessage = "Wybierz folder!";
            StatusColor = "Red";
            await Task.Delay(100);
            return;
        }
        try 
        {
            EndDate = EndDate.Date.AddDays(1).AddTicks(-1);
            List<EmailData> emailsToExport = _outlookService.GetEmailsForExport(
                SelectedFolder.FullPath, 
                StartDate.DateTime, 
                EndDate.DateTime);

            if (emailsToExport.Count > 0)
            {
                StatusMessage = "Generowanie...";
                StatusColor = "Blue";
                await Task.Delay(100);
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string fileName = $"{SelectedFolder.FolderName}.csv";
                string fullPath = Path.Combine(desktopPath, fileName);

                _csvService.ExportEmails(emailsToExport, fullPath);

                StatusMessage = "Sukces! Raport zapisano na pulpicie.";
                StatusColor = "Green";
                await Task.Delay(100);
                System.Diagnostics.Debug.WriteLine($"Zapisano w: {fullPath}");
            }
            else
            {
                StatusMessage = "Brak maili do eksportu w podanym przedziale czasowym.";
                StatusColor = "Orange";
                await Task.Delay(100);
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Błąd: {ex.Message}";
            StatusColor = "Red";
            await Task.Delay(100);
            System.Diagnostics.Debug.WriteLine($"Błąd exportu: {ex.Message}");
        }
    }
    private void ResetStatus()
    {
        StatusMessage = "Przetwarzanie informacji (generuj raport).";
        StatusColor = "Gray";
    }
    public ObservableCollection<OutlookFolder> Folders { get; } = new();

    public MainWindowViewModel()
    {
        DateTimeOffset currentDate = DateTimeOffset.Now;
        StartDate = new DateTimeOffset(currentDate.Year, currentDate.Month, 1, 0, 0, 0, currentDate.Offset).AddMonths(-1);
        EndDate = StartDate.AddMonths(1).AddTicks(-1);
        LoadOutlookFolders();
    }

    private void LoadMockFolders()
    {
        Folders.Clear();
        
        var root = new OutlookFolder 
        { 
            FolderName = "twoje.imię@outlook.com",
            SubFolders = new List<OutlookFolder>
            {
                new OutlookFolder { FolderName = "Skrzynka odbiorcza", SubFolders = new() {
                    new OutlookFolder { FolderName = "Praca" },
                    new OutlookFolder { FolderName = "Prywatne" }
                }},
                new OutlookFolder { FolderName = "Elementy wysłane" },
                new OutlookFolder { FolderName = "Kosz" }
            }
        };
        
        Folders.Add(root);
    }

    private void LoadOutlookFolders()
    {
        Folders.Clear();
        try
        {
            var service = new OutlookService();
            var realFolders = service.GetFolders();

            if (realFolders.Count > 0)
            {
                foreach (var f in realFolders) Folders.Add(f);
            }
            else
            {
                LoadMockFolders();
                Folders.Add(new OutlookFolder { FolderName = "Nie znaleziono folderów Outlooka - używam Mocków" });
            }
        }
        catch (Exception ex)
        {
            LoadMockFolders();
            System.Diagnostics.Debug.WriteLine($"Błąd podczas ładowania: {ex.Message}");
        }
    }
}