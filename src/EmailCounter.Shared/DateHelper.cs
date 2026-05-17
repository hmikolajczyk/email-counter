namespace EmailCounter.Shared
{
    public static class DateHelper
    {
        public static (DateTimeOffset start, DateTimeOffset end) CalculatePreviousMonthRange(DateTimeOffset currentDate)
        {
            DateTimeOffset startDate = new DateTimeOffset(currentDate.Year, currentDate.Month, 1, 0, 0, 0, currentDate.Offset).AddMonths(-1);
            DateTimeOffset endDate = startDate.AddMonths(1).AddTicks(-1);

            return (startDate, endDate);
        }
    }
}