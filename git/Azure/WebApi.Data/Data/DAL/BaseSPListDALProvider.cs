using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO;
using Contoso.Portal.Data.Extensions;
using PnP.Core;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using PnP.Core.Services;

namespace Contoso.Portal.Data.DAL
{
    public abstract class BaseSPListDALprovider<T> where T : BaseSPListItemDAO, new()
    {
        private string? _listCTContentTypeIdCache = null;
        private T _daoInstance = new T(); // used to access virtual properties

        public abstract string ListSiteRelavteUrl { get; }
        public abstract string DefaultView { get; }
        public virtual string DAOBaseContentTypeId => _daoInstance.GetContentTypeIdValue();

        public virtual IList<string> DefaultViewFields => BaseSPListItemDAO.AllFieldNames<T>();
        public virtual IList<(string Field, bool Ascending)> DefaultViewSort { get; } = new List<(string Field, bool Ascending)>() {
            (nameof(BaseSPListItemDAO.Created), false)
        };

        #region CAML Helper functions

        public virtual string CAML_GetInOperatorQuery(string fieldName, string fieldType, List<int> values)
        {
            var value = string.Concat(values.Select(v => $"<Value Type='{fieldType}'>{v}</Value>"));
            return CAML_GetOperatorQuery(fieldName, fieldType, value, ComparisonOperators.In, true);
        }

        public virtual string CAML_GetInOperatorQuery(string fieldName, string fieldType, List<string> values)
        {
            var value = string.Concat(values.Select(v => $"<Value Type='{fieldType}'>{v}</Value>"));
            return CAML_GetOperatorQuery(fieldName, fieldType, value, ComparisonOperators.In, true);
        }

        public virtual string CAML_GetOperatorQuery(string fieldName, string fieldType, string value, ComparisonOperators op, bool isLookup = false)
        {
            var lookupId = isLookup ? "LookupId='TRUE'" : string.ECNTy;
            switch (op)
            {
                case ComparisonOperators.Eq:
                    {
                        return $"<Eq><FieldRef Name='{fieldName}' {lookupId}/><Value Type='{fieldType}'>{value}</Value></Eq>";
                    }
                case ComparisonOperators.BeginsWith:
                    {
                        return $"<BeginsWith><FieldRef Name='{fieldName}' {lookupId}/><Value Type='{fieldType}'>{value}</Value></BeginsWith>";
                    }
                case ComparisonOperators.Geq:
                    {
                        return $"<Geq><FieldRef Name='{fieldName}' {lookupId}/><Value Type='{fieldType}'>{value}</Value></Geq>";
                    }
                case ComparisonOperators.Leq:
                    {
                        return $"<Leq><FieldRef Name='{fieldName}' {lookupId}/><Value Type='{fieldType}'>{value}</Value></Leq>";
                    }
                case ComparisonOperators.In:
                    {
                        return $"<In><FieldRef Name='{fieldName}' {lookupId}/><Values>{value}</Values></In>";
                    }
                case ComparisonOperators.Neq:
                    {
                        return $"<Neq><FieldRef Name='{fieldName}' {lookupId}/><Value Type='{fieldType}'>{value}</Value></Neq>";
                    }
                case ComparisonOperators.Lt:
                    {
                        return $"<Lt><FieldRef Name='{fieldName}' {lookupId}/><Value Type='{fieldType}'>{value}</Value></Lt>";
                    }
                case ComparisonOperators.Contains:
                    {
                        return $"<Contains><FieldRef Name='{fieldName}' {lookupId}/><Value Type='{fieldType}'>{value}</Value></Contains>";
                    }
                default:
                    return "Invalid query";
            }

        }

        public virtual string CAML_GetOperatorQuery(string fieldName, DateTime dateTime, ComparisonOperators op)
        {
            switch (op)
            {
                case ComparisonOperators.Geq:
                    {
                        return $"<Geq><FieldRef Name='{fieldName}'/><Value IncludeTimeValue='TRUE' Type='DateTime' StorageTZ='TRUE'>{dateTime.ToUniversalTime():O}</Value></Geq>";
                    }
                case ComparisonOperators.Leq:
                    {
                        return $"<Leq><FieldRef Name='{fieldName}'/><Value IncludeTimeValue='TRUE' Type='DateTime' StorageTZ='TRUE'>{dateTime.ToUniversalTime():O}</Value></Leq>";
                    }
                default:
                    throw new ArgumentException($"Invalid operator {op} for DateTime field");
            }
        }

        public virtual string CAML_GetContentTypeFilterPart()
        {
            // if we know the content type id in list, we can optimize the query
            if (_listCTContentTypeIdCache is not null)
                return CAML_GetOperatorQuery(nameof(BaseSPListItemDAO.ContentTypeId), "ContentTypeId", _listCTContentTypeIdCache, ComparisonOperators.Eq);
            else
                return CAML_GetOperatorQuery(nameof(BaseSPListItemDAO.ContentTypeId), "ContentTypeId", _daoInstance.GetContentTypeIdValue(), ComparisonOperators.BeginsWith);
        }

        public string CAML_GetDateRangeFilterPart(string field, DateTime? startDate, DateTime? endDate)
        {
            if (startDate is null && endDate is null)
                throw new ArgumentException($"At least one of {nameof(startDate)} and {nameof(endDate)} must be provided");

            var startDateQuery = startDate != null ? CAML_GetOperatorQuery(field, startDate.Value, ComparisonOperators.Geq) : null;
            var endDateQuery = endDate != null ? CAML_GetOperatorQuery(field, endDate.Value, ComparisonOperators.Leq) : null;

            return (startDate is null ? endDateQuery : endDate is null ? startDateQuery : $"<And>{startDateQuery}{endDateQuery}</And>") ?? throw new Exception("Start date or end date must be provided");
        }

        #endregion

        #region Generic functions
        // Generic functions
        public async virtual Task<IEnumerable<T>> GetByView(IPnPContext ctx, string view, string? folderServerRelativeUrl = null)
        {
            try
            {
                var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl), p => p.Title, p => p.Id, p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title));
                return await list.GetAllQueryResultsAsEntityAsync(new RenderListDataOptions()
                {
                    ViewXml = view,
                    FolderServerRelativeUrl = folderServerRelativeUrl
                }, (r) =>
                {
                    var result = new T();
                    result.MapFromSPListItem(r);
                    return result;
                });
            }
            catch (ServiceException ex)
            {
                throw new Exception($"Error getting items from list '{ListSiteRelavteUrl}' in site '{ctx.Uri.AbsolutePath}': {ex.Error}", ex);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error getting items from list '{ListSiteRelavteUrl}' in site '{ctx.Uri.AbsolutePath}': {ex.Message}", ex);
            }
        }

        public async virtual Task<T?> GetById(IPnPContext ctx, int id)
        {
            return (await GetByView(ctx, PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetOperatorQuery(nameof(BaseSPListItemDAO.ID), "Counter", id.ToString(), ComparisonOperators.Eq)}</Where>", DefaultViewFields))).FirstOrDefault();
        }

        public async virtual Task<T?> GetByUniqueId(IPnPContext ctx, string id)
        {
            return (await GetByView(ctx, PnPContentHelpers.CamlViewBuilder($"<Where><Eq><FieldRef Name='UniqueId'/><Value Type='Guid'>{id}</Value></Eq></Where>", DefaultViewFields))).FirstOrDefault();
        }

        public async virtual Task<(T CurrentState, T testviousState)> GetByIdAndVersionLabel(IPnPContext ctx, int itemId, string versionLabel)
        {
            var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl), p => p.Title, p => p.RootFolder, p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title));

            var dalitem = await GetById(ctx, itemId);
            var item = dalitem.AsListItem();

            var oldItem = await MSVersionLookupIssueHelper.GetVersionAsync(item, versionLabel);

            // map to dao
            var resultItem = new T();
            var resultOldItem = new T();
            resultItem.MapFromSPListItem(item);
            resultOldItem.MapFromSPListItemVersion(oldItem);

            return (resultItem, resultOldItem);
        }

        public async virtual Task<(T CurrentState, T testviousState, string[] ChangedFields)> GetChangesFromVersion(IPnPContext ctx, int itemId, int versionId)
        {
            var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl), p => p.Title, p => p.RootFolder, p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title));

            var dalitem = await GetById(ctx, itemId);
            var item = dalitem.AsListItem();

            var oldItem = await MSVersionLookupIssueHelper.GetVersionAsync(item, versionId);
            var fields = BaseSPListItemDAO.AllFieldNames<T>().ToList();

            // same or irrelevant fields
            fields.Remove(nameof(BaseSPListItemDAO.ID));
            fields.Remove(nameof(BaseSPListItemDAO.Created));
            fields.Remove(nameof(BaseSPListItemDAO.Author));
            fields.Remove(nameof(BaseSPListItemDAO.ContentTypeId));
            fields.Remove(nameof(BaseSPListItemDAO.Modified));

            List<string> changedFields = PnPContentHelpers.IdentifyChangedFields(item, oldItem, fields);

            // map to dao
            var resultItem = new T();
            var resultOldItem = new T();
            resultItem.MapFromSPListItem(item);
            resultOldItem.MapFromSPListItemVersion(oldItem);

            return (resultItem, resultOldItem, changedFields.ToArray());
        }

        public async Task<(T ToState, T FromState, T CurrentState, string[] ChangedFields)> GetChangesFromVersion(IPnPContext ctx, int itemId, string fromVersionLabel, string toVersionLabel)
        {
            var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl), p => p.Title, p => p.RootFolder, p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title));

            var dalitem = await GetById(ctx, itemId);
            var item = dalitem.AsListItem();

            var (newItem, oldItem) = await MSVersionLookupIssueHelper.GetVersionAsync(item, fromVersionLabel, toVersionLabel);

            var fields = BaseSPListItemDAO.AllFieldNames<T>().ToList();

            // same or irrelevant fields
            fields.Remove(nameof(BaseSPListItemDAO.ID));
            fields.Remove(nameof(BaseSPListItemDAO.Created));
            fields.Remove(nameof(BaseSPListItemDAO.Author));
            fields.Remove(nameof(BaseSPListItemDAO.ContentTypeId));
            fields.Remove(nameof(BaseSPListItemDAO.Modified));

            var changedFields = PnPContentHelpers.IdentifyChangedFields(newItem, oldItem, fields);

            // map to dao
            var toItem = new T();
            var fromItem = new T();
            var currentItem = new T();
            toItem.MapFromSPListItemVersion(newItem);
            fromItem.MapFromSPListItemVersion(oldItem);
            currentItem.MapFromSPListItem(item);

            return (toItem, fromItem, currentItem, changedFields.ToArray());
        }

        internal static class MSVersionLookupIssueHelper
        {
            internal static async Task<IListItemVersion> GetVersionAsync(IListItem listItem, string versionLabel)
            {
                var result = await listItem.Versions.FirstOrDefaultAsync(v => v.VersionLabel == versionLabel);
                if (result is null) // MS ISSUE: filtering by label or id not working in several tenants
                {
                    await listItem.Versions.LoadAsync();
                    result = listItem.Versions.FirstOrDefault(v => v.VersionLabel == versionLabel);
                }

                return result ?? throw new Exception($"Version {versionLabel} not found for item {listItem.Id}");
            }

            internal static async Task<IListItemVersion> GetVersionAsync(IListItem listItem, int versionId)
            {
                var result = await listItem.Versions.FirstOrDefaultAsync(v => v.Id == versionId);
                if (result is null) // MS ISSUE: filtering by label or id not working in several tenants
                {
                    await listItem.Versions.LoadAsync();
                    result = listItem.Versions.FirstOrDefault(v => v.Id == versionId);
                }

                return result ?? throw new Exception($"Version {versionId} not found for item {listItem.Id}");
            }

            internal static async Task<(IListItemVersion, IListItemVersion)> GetVersionAsync(IListItem listItem, string versionLabel1, string versionLabel2)
            {
                var resultList = await listItem.Versions.Where(v => v.VersionLabel == versionLabel1 || v.VersionLabel == versionLabel2).ToListAsync();
                if (resultList is null || resultList.Count != 2) // MS ISSUE: filtering by label or id not working in several tenants
                {
                    await listItem.Versions.LoadAsync();
                    resultList = listItem.Versions.Where(v => v.VersionLabel == versionLabel1 || v.VersionLabel == versionLabel2).ToList();
                }

                if (resultList is null || resultList.Count != 2)
                    throw new Exception($"Versions {versionLabel1} and {versionLabel2} not found for item {listItem.Id}");

                return (resultList.First(), resultList.Last());
            }
        }

        public async virtual Task<IEnumerable<T>> GetByIds(IPnPContext ctx, int[] ids)
        {
            return await GetByView(ctx, PnPContentHelpers.CamlViewBuilder($"<Where><In><FieldRef Name='{nameof(BaseSPListItemDAO.ID)}'/><Values>{string.Join("", ids.Select(id => $"<Value Type='Counter'>{id}</Value>"))}</Values></In></Where>", DefaultViewFields));
        }

        public async virtual Task<string> GetContentTypeIdValueInList(IPnPContext ctx)
        {
            if (_listCTContentTypeIdCache is not null)
                return _listCTContentTypeIdCache;
            try
            {
                // Get the content type to update
                var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl), p => p.ContentTypes.Queryproperties(p => p.Name, p => p.Id));
                _listCTContentTypeIdCache = list.ContentTypes.AsRequested().FirstOrDefault(p => p.Id.StartsWith(_daoInstance.GetContentTypeIdValue()))?.Id;
            }
            catch (Exception)
            {
                // ignore
            }
            return _listCTContentTypeIdCache ?? _daoInstance.GetContentTypeIdValue();
        }

        public async virtual Task<T> Add(IPnPContext ctx, T item)
        {
            try
            {
                var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl), p => p.Title, p => p.Id, p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title));
                var newListItem = item.AsNewListItem();
                var newItem = await list.Items.AddAsync(newListItem);
                var queriedItem = await list.Items.GetByIdAsync(newItem.Id);
                if (!queriedItem.Values.ContainsKey("FileRef"))
                    queriedItem.Values.Add("FileRef", "");
                var result = new T();
                result.MapFromSPListItem(queriedItem);
                return result;
            }
            catch (ServiceException ex)
            {
                throw new Exception($"Error creating new item for list '{ListSiteRelavteUrl}' in site '{ctx.Uri.AbsolutePath}': {ex.Error}", ex);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error creating new item for list '{ListSiteRelavteUrl}' in site '{ctx.Uri.AbsolutePath}': {ex.Message}", ex);
            }
        }

        public async virtual Task Delete(IPnPContext ctx, int id)
        {
            try
            {
                var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl), p => p.Title, p => p.Id, p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title));
                await list.Items.DeleteByIdAsync(id);
            }
            catch (ServiceException ex)
            {
                throw new Exception($"Error deleting item for list '{ListSiteRelavteUrl}' in site '{ctx.Uri.AbsolutePath}': {ex.Error}", ex);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error deleting item for list '{ListSiteRelavteUrl}' in site '{ctx.Uri.AbsolutePath}': {ex.Message}", ex);
            }
        }

        #endregion
    }
}