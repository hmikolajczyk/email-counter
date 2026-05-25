using EmailCounter.Gui.ViewModels;
using EmailCounter.Shared;
using Xunit;

namespace EmailCounter.Tests;

public class DateHelperTests
{
    [Fact]
    public void GetPreviousMonthRange_ShouldReturnCorrectDateRange()
    {
        DateTimeOffset inputDate = new DateTimeOffset(2026, 5, 25, 0, 0, 0, TimeSpan.Zero);
        DateTimeOffset expectedStart = new DateTimeOffset(2026, 4, 1, 0, 0, 0, TimeSpan.Zero);
        DateTimeOffset expectedEnd = new DateTimeOffset(2026, 4, 30, 23, 59, 59, TimeSpan.Zero);

        var (actualStart, actualEnd) = DateHelper.GetPreviousMonthRange(inputDate);

        Assert.Equal(expectedStart, actualStart);
        Assert.Equal(expectedEnd, actualEnd, TimeSpan.FromSeconds(1));
    }

    [Fact]
    public void GetPreviousMonthRange_YearTransitioning_ShouldReturnDecemberRange()
    {
        DateTimeOffset inputDate = new DateTimeOffset(2026, 1, 25, 0, 0, 0, TimeSpan.Zero);
        DateTimeOffset expectedStart = new DateTimeOffset(2025, 12, 1, 0, 0, 0, TimeSpan.Zero);
        DateTimeOffset expectedEnd = new DateTimeOffset(2025, 12, 31, 23, 59, 59, TimeSpan.Zero);

        var (actualStart, actualEnd) = DateHelper.GetPreviousMonthRange(inputDate);

        Assert.Equal(expectedStart, actualStart);
        Assert.Equal(expectedEnd, actualEnd, TimeSpan.FromSeconds(1));
    }

    [Fact]
    public void GetPreviousMonthRange_NonLeapYear_ShouldReturnCorrectFebruaryRange()
    {
        DateTimeOffset inputDate = new DateTimeOffset(2026, 3, 25, 0, 0, 0, TimeSpan.Zero);
        DateTimeOffset expectedStart = new DateTimeOffset(2026, 2, 1, 0, 0, 0, TimeSpan.Zero);
        DateTimeOffset expectedEnd = new DateTimeOffset(2026, 2, 28, 23, 59, 59, TimeSpan.Zero);

        var (actualStart, actualEnd) = DateHelper.GetPreviousMonthRange(inputDate);

        Assert.Equal(expectedStart, actualStart);
        Assert.Equal(expectedEnd, actualEnd, TimeSpan.FromSeconds(1));
    }

    [Fact]
    public void GetPreviousMonthRange_LeapYear_ShouldReturnCorrectFebruaryRange()
    {
        DateTimeOffset inputDate = new DateTimeOffset(2024, 3, 25, 0, 0, 0, TimeSpan.Zero);
        DateTimeOffset expectedStart = new DateTimeOffset(2024, 2, 1, 0, 0, 0, TimeSpan.Zero);
        DateTimeOffset expectedEnd = new DateTimeOffset(2024, 2, 29, 23, 59, 59, TimeSpan.Zero);

        var (actualStart, actualEnd) = DateHelper.GetPreviousMonthRange(inputDate);

        Assert.Equal(expectedStart, actualStart);
        Assert.Equal(expectedEnd, actualEnd, TimeSpan.FromSeconds(1));
    }
}