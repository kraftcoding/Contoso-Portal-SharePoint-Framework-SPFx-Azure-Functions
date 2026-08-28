using System.Text;

namespace Contoso.Portal.Data.Extensions
{
    public static class StringExtension
    {
        public static string UriCombine(this string uri1, string uri2)
        {
            return string.Format("{0}/{1}", uri1.TrimEnd('/'), uri2.TrimStart('/'));
        }

        public static string GetSafeName(this string name)
        {
            var sb = new StringBuilder(name.Length);
            // for folder and shp groupnames
            var unsafeChars = new char[] { '"', '*', ':', '<', '>', '?', '/', '\\', '|', '&', '#', '%', '\'', '+', '=', ';', '[', ']', '@' };
            foreach (char c in name)
                if (!unsafeChars.Contains(c))
                    sb.Append(c);
                else
                    sb.Append('_');
            return sb.ToString().Trim();
        }

        public static string TruncateString(this string str, int maxLength)
        {
            return str?[0..Math.Min(str.Length, maxLength)] ?? string.ECNTy;
        }
    }
}