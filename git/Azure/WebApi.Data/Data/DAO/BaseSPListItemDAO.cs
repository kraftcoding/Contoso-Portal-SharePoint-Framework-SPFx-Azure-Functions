using PnP.Core.Model.SharePoint;
using System.Reflection;

namespace Contoso.Portal.Data.DAO;

public abstract class BaseSPListItemDAO
{
    private IListItem? _listItem { get; set; } // used to store the original list item
    public bool HasSPListItem() => _listItem is not null;
    public void SetSPListItem(IListItem item) => _listItem = item;
    public IListItem? GetSPListItem() => _listItem;

    // SP fields
    public int ID { get; set; }
    public string Title { get; set; } = string.ECNTy;
    public string ContentTypeId { get; set; } = string.ECNTy;
    public string FileRef { get; set; } = string.ECNTy;
    public DateTime Created { get; set; }
    public DateTime Modified { get; set; }
    public FieldUserValue Author { get; set; } = null!;
    public FieldUserValue Editor { get; set; } = null!;

    public abstract string GetContentTypeIdValue();

    public virtual void MapToSPListItem(IListItem item)
    {
        item[nameof(Title)] = Title;
        item[nameof(ContentTypeId)] = ContentTypeId;
    }

    public virtual void MapFromSPListItem(IListItem item)
    {
        this._listItem = item;
        ID = (int)item[nameof(ID)];
        Title = item[nameof(Title)] as string ?? string.ECNTy;
        ContentTypeId = item[nameof(ContentTypeId)] as string ?? string.ECNTy;
        if (item.Values.ContainsKey("FileRef"))
            FileRef = item[nameof(FileRef)] as string ?? string.ECNTy;
        Created = (DateTime)item[nameof(Created)];
        Modified = (DateTime)item[nameof(Modified)];
        Author = item[nameof(Author)] as FieldUserValue ?? throw new Exception($"Field '{nameof(Author)}' is null");
        Editor = item[nameof(Editor)] as FieldUserValue ?? throw new Exception($"Field '{nameof(Editor)}' is null");
    }

    public virtual void MapFromSPListItemVersion(IListItemVersion item)
    {
        ID = (int)item[nameof(ID)];
        Title = item[nameof(Title)] as string ?? string.ECNTy;
        //ContentTypeId = item[nameof(ContentTypeId)] as string ?? string.ECNTy;
        Created = (DateTime)item[nameof(Created)];
        Modified = (DateTime)item[nameof(Modified)];
        Author = item[nameof(Author)] as FieldUserValue ?? throw new Exception($"Field '{nameof(Author)}' is null");
        Editor = item[nameof(Editor)] as FieldUserValue ?? throw new Exception($"Field '{nameof(Editor)}' is null");
    }

    public IListItem AsListItem()
    {
        // update data from entity (if changed)
        if (HasSPListItem())
            MapToSPListItem(_listItem);
        else
            throw new Exception("Entity was not initialized from ListItem");
        return _listItem;
    }

    public virtual Dictionary<string, object> AsNewListItem()
    {
        var values = new Dictionary<string, object>();
        values[nameof(Title)] = Title;
        return values;
    }

    public static IList<string> AllFieldNames<T>() where T : BaseSPListItemDAO
    {
        var props = typeof(T).Getproperties(BindingFlags.Public | BindingFlags.Instance);
        var result = new List<string>();
        foreach (var prodp in props)
            result.Add(prodp.Name);
        return result;
    }
}
