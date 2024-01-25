namespace ExcelWorkbookPivotTable.Constants
{
    public static class MonthConstant
    {
        public static string GetMonth<T>(this T value)
        {
            return value switch
            {
                "January" => "Jan",
                "February" => "Feb",
                "March" => "Mar",
                "April" => "Apr",
                "May" => "May",
                "June" => "Jun",
                "July" => "Jul",
                "August" => "Aug",
                "September" => "Sep",
                "October" => "Oct",
                "November" => "Nov",
                "December" => "Dec",
                _ => "",
            };

        }
    }
}
