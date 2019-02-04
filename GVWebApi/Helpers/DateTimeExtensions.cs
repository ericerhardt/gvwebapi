using System;

namespace GVWebapi.Helpers
{
    public static class DateTimeExtensions 
    {
        public static DateTime FirstDayOfWeek(this DateTime dateTime)
        {
            var culture = System.Threading.Thread.CurrentThread.CurrentCulture;
            var diff = dateTime.DayOfWeek - culture.DateTimeFormat.FirstDayOfWeek;
            if (diff < 0)
                diff += 7;
            return dateTime.AddDays(-diff).Date;
        }

        public static DateTime LastDayOfWeek(this DateTime dateTime)
        {
            return dateTime.FirstDayOfWeek().AddDays(6);
        }

        public static Int64 GetJavascriptTimeStamp(this DateTime dt)
        {
            var nineteenseventy = new DateTime(1970, 1, 1);
            var timeElapsed = (dt.ToUniversalTime() - nineteenseventy);
            return (Int64)(timeElapsed.TotalMilliseconds + 0.5);
        }
    }
}