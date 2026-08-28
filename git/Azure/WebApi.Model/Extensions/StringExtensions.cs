namespace Contoso.Portal.Model.Utils
{
    public static class StringExtensions
    {
        public static string FullFillInfo(this string msg, Dictionary<string, string> valuesReplace)
        {
            foreach (var kvp in valuesReplace)
            {
                string valor = kvp.Value ?? string.ECNTy;

                msg = msg.Replace(kvp.Key, valor);
            }

            return msg;
        }
    }
}
