using PnP.Core.Model.SharePoint;

namespace Contoso.Portal.Data.Extensions
{
    public static class PnPSearchExtensions
    {
        public delegate T SearchResultToEntityMapper<T>(IDictionary<string, object> result);

        private static async Task<(IEnumerable<T> PageResults, long TotalRows)> PaginateResults<T>(IWeb web, SearchOptions sOptions, SearchResultToEntityMapper<T> mapper, int startRow)
        {
            var searchResults = new List<Dictionary<string, object>>();
            var currentSearchOptions = new SearchOptions(sOptions.Query)
            {
                ClientType = sOptions.ClientType,
                RefinementFilters = sOptions.RefinementFilters,
                Refineproperties = sOptions.Refineproperties,
                RowLimit = sOptions.RowLimit,
                RowsPerPage = sOptions.RowsPerPage,
                Selectproperties = sOptions.Selectproperties,
                Sortproperties = sOptions.Sortproperties,
                StartRow = startRow,
                TrimDuplicates = sOptions.TrimDuplicates
            };
            var searchResult = await web.SearchAsync(currentSearchOptions);
            return (searchResult.Rows.Select(r => mapper(r)), searchResult.TotalRows);
        }

        public static async Task<IEnumerable<T>> GetAllSearchResultsAsEntityAsync<T>(this IWeb web, SearchOptions sOptions, SearchResultToEntityMapper<T> mapper)
        {
            var results = new List<T>();

            bool paging = true;
            int startRow = 0;

            while (paging)
            {
                (var pageResults, var totalRows) = await PaginateResults(web, sOptions, mapper, startRow);
                results.AddRange(pageResults);

                if (results.Count < totalRows)
                    startRow = results.Count;
                else
                    paging = false;
            }

            return results;
        }

        public async static IAsyncEnumerable<T> EnumerateAllSearchResultsAsEntityAsync<T>(this IWeb web, SearchOptions sOptions, SearchResultToEntityMapper<T> mapper)
        {
            long totalResultsCount = 0;
            bool paging = true;
            int startRow = 0;

            while (paging)
            {
                (var pageResults, var totalRows) = await PaginateResults(web, sOptions, mapper, startRow);
                foreach (var result in pageResults)
                {
                    totalResultsCount++;
                    yield return result;
                }

                if (totalResultsCount < totalRows)
                    startRow = pageResults.Count();
                else
                    paging = false;
            }
        }
    }
}