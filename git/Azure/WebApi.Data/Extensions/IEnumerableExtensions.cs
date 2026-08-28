using System.Text;

namespace Contoso.Portal.Data.Extensions
{
    public static class IEnumerableExtension
    {
        public static IEnumerable<IEnumerable<T>> Partition<T>(this IEnumerable<T> source, int size)
        {
            int count = 0;
            T[]? group = null; // use arrays as buffer
            foreach (T item in source)
            {
                group ??= new T[size];
                group[count++] = item;
                if (count == size)
                {
                    yield return group;
                    group = null;
                    count = 0;
                }
            }
            if (count > 0)
            {
                Array.Resize(ref group, count);
                yield return group;
            }
        }

        public static IEnumerable<T> WhereNotNull<T>(this IEnumerable<T?>? source)
        {
            foreach (var item in source ?? [])//collection or item might be null
                if (item is T notNullItem)//non exception throwing conversion with success check
                    yield return notNullItem;
        }
    }
}