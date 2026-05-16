using EmailCounter.Gui.ViewModels;
using Xunit;

namespace EmailCounter.Tests;

public class StatusTests
{
    [Fact]
    public void ResetStatus_ShouldSetDefaultValues()
    {
        var vm = new MainWindowViewModel();
        vm.StatusMessage = "ERROR";
        vm.StatusColor = "Red";
        vm.SelectedFolder = null; 

        Assert.Equal("", vm.StatusMessage);
        Assert.Equal("Gray", vm.StatusColor);
    }
}