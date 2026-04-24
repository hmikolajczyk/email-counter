using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using EmailCounter.Shared.Models;
using EmailCounter.Shared.Services;

namespace EmailCounter.Gui.ViewModels;

public class MainWindowViewModel : ViewModelBase
{
    public ObservableCollection<OutlookFolder> Folders { get; } = new();

    public MainWindowViewModel()
    {
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
                new OutlookFolder { FolderName = "Kosz" },
                new OutlookFolder { FolderName = "Elementy wysłane" },
                new OutlookFolder { FolderName = "Kosz" },
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