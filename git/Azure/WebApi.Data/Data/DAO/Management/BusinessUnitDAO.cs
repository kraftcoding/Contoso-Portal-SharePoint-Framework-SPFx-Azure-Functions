using Contoso.Portal.Data.DAL;
using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.Management;

public class BusinessUnitDAO : BaseSPListItemDAO
{
    public override string GetContentTypeIdValue() => ContentTypeIds.BusinessUnit;

    public FieldTaxonomyValue? BusinessArea { get; set; }
    public FieldLookupValue? DivisionLookup { get; set; }
    public DateTime? StartDate { get; set; }
    public DateTime? EndDate { get; set; }

    public override void MapToSPListItem(IListItem item)
    {
        base.MapToSPListItem(item);

        item[nameof(BusinessArea)] = BusinessArea;
        item[nameof(DivisionLookup)] = DivisionLookup;
        item[nameof(StartDate)] = StartDate;
        item[nameof(EndDate)] = EndDate;
    }

    public override void MapFromSPListItem(IListItem item)
    {
        base.MapFromSPListItem(item);

        BusinessArea = item[nameof(BusinessArea)] as FieldTaxonomyValue;
        DivisionLookup = item[nameof(DivisionLookup)] as FieldLookupValue;
        StartDate = item[nameof(StartDate)] is not null ? (DateTime)item[nameof(StartDate)] : null;
        EndDate = item[nameof(EndDate)] is not null ? (DateTime)item[nameof(EndDate)] : null;
    }

    public override Dictionary<string, object> AsNewListItem()
    {
        var values = base.AsNewListItem();
        values[nameof(BusinessArea)] = BusinessArea;
        values[nameof(DivisionLookup)] = DivisionLookup;
        values[nameof(StartDate)] = StartDate ?? new DateTime();
        values[nameof(EndDate)] = EndDate ?? new DateTime();
        return values;
    }
}
