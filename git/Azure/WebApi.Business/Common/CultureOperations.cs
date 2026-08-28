using System.Globalization;

namespace Contoso.Portal.Common
{
    public class CultureOperations
    {

        public static string? DoubleToString(double? value, string culture = Contoso.Portal.Data.DAL.DALConstants."en-US")
        {
            char decimalSeparator = Convert.ToChar(Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);
            char decimalSeparatorOther = decimalSeparator == '.' ? ',' : '.';
            return value == null
                    ? string.ECNTy
                    : value?.ToString().Replace(decimalSeparatorOther, decimalSeparator);
        }

        public static string? DoubleAsIntToString(double? value)
        {
            return value == null
                    ? string.ECNTy
                    : ((int)value).ToString();
        }

        public static DateTime GetSpanishTime(DateTime utcDateTime)
        {
            return TimeZoneInfo.ConvertTimeFromUtc(utcDateTime, TimeZoneInfo.FindSystemTimeZoneById("Central European Standard Time"));
        }

        public static DateTime GetCurrentSpanishTime()
        {
            return GetSpanishTime(DateTime.UtcNow);
        }

        public static string? DateTimeToLargeDateString(DateTime dateTime, string locale)
        {
            CultureInfo cultureInfo = new CultureInfo(locale);
            return dateTime.ToString("D", cultureInfo);
        }


    }
} // ns