using System.Globalization;

namespace EmailCounter.Shared
{
    public static class DateHelper
    {
        public static (DateTimeOffset start, DateTimeOffset end) GetPreviousMonthRange(DateTimeOffset currentDate)
        {
            DateTimeOffset startDate = new DateTimeOffset(currentDate.Year, currentDate.Month, 1, 0, 0, 0, currentDate.Offset).AddMonths(-1);
            DateTimeOffset endDate = startDate.AddMonths(1).AddTicks(-1);

            return (startDate, endDate);
        }
        public static string FormatDateRange(DateTime? date)
        {
            if (date == null)
            {
                return string.Empty;
            }
            return date.Value.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
        }
    }
}