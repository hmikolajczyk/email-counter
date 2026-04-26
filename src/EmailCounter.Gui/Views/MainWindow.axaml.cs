using Avalonia.Controls;
using EmailCounter.Gui.ViewModels;
using EmailCounter.Shared.Models;

namespace EmailCounter.Gui.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        DataContext = new MainWindowViewModel();
        var treeView = this.FindControl<TreeView>("FoldersTreeView");
        treeView!.SelectionChanged += (s, e) => 
        {
            if (DataContext is MainWindowViewModel vm)
            {
                vm.SelectedFolder = treeView.SelectedItem as OutlookFolder;
            }
        };
    }
}