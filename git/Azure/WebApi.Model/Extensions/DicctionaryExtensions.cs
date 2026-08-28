namespace Contoso.Portal.Model.Extensions
{
    public static class DicctionaryExtensions
    {
        public static void AddWithCheck<T,U>(this Dictionary<T,U> dic, T key, U value)
        {
            if(dic != null && key != null && !dic.ContainsKey(key))
            {
                dic.Add(key, value);
            }
        }
    }
}
