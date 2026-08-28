using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;

namespace Contoso.Portal.Data.Extensions
{
    public static class PnPListQueryByRenderListDataExtensions
    {
        public delegate T QueryResultToEntityMapper<T>(IListItem result);

        private static async Task<(IEnumerable<T> PageResults, string? NextPage)> PaginateResults<T>(IList list, RenderListDataOptions rlOptions, QueryResultToEntityMapper<T> mapper, string? nextPage)
        {
            list.Items.Clear();
            var results = await list.LoadListDataAsStreamAsync(new RenderListDataOptions()
            {
                ViewXml = rlOptions.ViewXml,
                RenderOptions = RenderListDataOptionsFlags.ListData,
                Paging = nextPage,
                FolderServerRelativeUrl = rlOptions.FolderServerRelativeUrl,
                DatesInUtc = rlOptions.DatesInUtc,
                AllowMultipleValueFilterForTaxonomyFields = rlOptions.AllowMultipleValueFilterForTaxonomyFields,
                AddRequiredFields = rlOptions.AddRequiredFields,
                AudienceTarget = rlOptions.AudienceTarget,
                ImageFieldsToTryRewriteToCdnUrls = rlOptions.ImageFieldsToTryRewriteToCdnUrls,
                DeferredRender = rlOptions.DeferredRender,
                ExpandGroups = rlOptions.ExpandGroups,
                FirstGroupOnly = rlOptions.FirstGroupOnly,
                ReplaceGroup = rlOptions.ReplaceGroup,
                OverrideViewXml = rlOptions.OverrideViewXml
            }).ConfigureAwait(false);
            return (list.Items.AsRequested().Select(r => mapper(r)), results.ContainsKey("NextHref") ? results["NextHref"]?.ToString()?.Substring(1) : null);
        }

        public static async Task<IEnumerable<T>> GetAllQueryResultsAsEntityAsync<T>(this IList list, RenderListDataOptions rlOptions, QueryResultToEntityMapper<T> mapper)
        {
            var results = new List<T>();
            string? nextPage = null;
            do
            {
                (var pageResults, nextPage) = await PaginateResults(list, rlOptions, mapper, nextPage);
                results.AddRange(pageResults);
            } while (!string.IsNullOrECNTy(nextPage));
            return results;
        }

        public async static IAsyncEnumerable<T> EnumerateQueryResultsAsEntityAsync<T>(this IList list, RenderListDataOptions rlOptions, QueryResultToEntityMapper<T> mapper)
        {
            var results = new List<T>();
            string? nextPage = null;
            do
            {
                (var pageResults, nextPage) = await PaginateResults(list, rlOptions, mapper, nextPage);
                foreach (var result in pageResults)
                    yield return result;
            } while (!string.IsNullOrECNTy(nextPage));
        }

        /* do not use this method, it is not working prodperly */
        // private static async Task<(IEnumerable<T> PageResults, string? NextPage)> PaginateResultsCamlQuery<T>(IList list, RenderListDataOptions rlOptions, QueryResultToEntityMapper<T> mapper, string? nextPage)
        // {
        //     list.Items.Clear();
        // 
        //     await list.LoadItemsByCamlQueryAsync(new CamlQueryOptions()
        //     {
        //         ViewXml = rlOptions.ViewXml,
        //         DatesInUtc = true,
        //         PagingInfo = nextPage,
        //         FolderServerRelativeUrl = rlOptions.FolderServerRelativeUrl
        //     }, p => p.FieldValuesForEdit);
        // 
        //     return (list.Items.AsRequested().Select(r => mapper(r)), null);
        // }
    }
}
