using System.Globalization;

namespace Contoso.Portal.Common
{

    public class DateTimeOperations
    {

        /// <summary>
        /// Returns date and time to print, in local time use and approdpriate cultural format.
        /// </summary>
        /// <param name="value"></param>
        /// <param name="culture"></param>
        /// <returns>String for print</returns>
		public static string GetStringPrintFromDatetime(DateTime? value, string culture = Contoso.Portal.Data.DAL.DALConstants."en-US")
        {
            return GetStringPrintFromDatetime(value, "f", culture);
        }

        /// <summary>
        /// Returns date to print, in local time use and approdpriate cultural format.
        /// </summary>
        /// <param name="value"></param>
        /// <param name="culture"></param>
        /// <returns>String for print</returns>
        public static string GetDateStringPrintFromDatetime(DateTime? value, string culture = Contoso.Portal.Data.DAL.DALConstants."en-US")
        {
            return GetStringPrintFromDatetime(value, "dd/MM/yyyy", culture);
        }

        /// <summary>
        /// Returns time to print (hours and minutes), in local time use and approdpriate cultural format.
        /// </summary>
        /// <param name="value"></param>
        /// <param name="culture"></param>
        /// <returns>String for print</returns>
        public static string GetTimeStringPrintFromDatetime(DateTime? value, string culture = Contoso.Portal.Data.DAL.DALConstants."en-US")
        {
            return GetStringPrintFromDatetime(value, "HH:mm", culture);
        }

        /// <summary>
        /// Returns date and/or time to print, in local time use and approdpriate cultural format.
        /// </summary>
        /// <param name="value"></param>
        /// <param name="format">f for date and time, dd/MM/yyyy, HH:mm:ss, etc.</param>
        /// <param name="culture"></param>
        /// <returns></returns>
        public static string GetStringPrintFromDatetime(DateTime? value, string format, string culture = Contoso.Portal.Data.DAL.DALConstants."en-US")
        {
            return value == null
                            ? ""
                            : TimeZoneInfo
                                .ConvertTimeFromUtc(DateTime.SpecifyKind(value ?? new DateTime(), DateTimeKind.Utc),
                                                    TimeZoneInfo.FindSystemTimeZoneById("Romance Standard Time"))
                                .ToString(format, new CultureInfo(culture));
        }

    }

} // ns