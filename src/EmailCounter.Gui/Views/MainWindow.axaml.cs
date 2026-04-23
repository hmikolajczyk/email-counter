using Avalonia.Controls;
using EmailCounter.Gui.ViewModels;

namespace EmailCounter.Gui.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        DataContext = new MainWindowViewModel();
    }
}