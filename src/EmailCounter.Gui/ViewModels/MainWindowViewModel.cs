using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using EmailCounter.Shared.Models;
using System.Threading.Tasks;

namespace EmailCounter.Gui.ViewModels;

public class MainWindowViewModel : ViewModelBase
{
    public ObservableCollection<OutlookFolder> Folders { get; } = new();

    public MainWindowViewModel()
    {
        LoadMockFolders();
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
}