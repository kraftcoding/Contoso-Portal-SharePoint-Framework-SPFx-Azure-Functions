using Contoso.Portal.Data.DAO;

namespace Contoso.Portal.Data.DTO;

public abstract class JsonFileListItemAbstractionDTO<T, U>
        where T : BaseSPListItemDAO, new()
        where U : BaseJsonDataDAO, new()
{
    public int Id { get => ItemDAO.ID; set => ItemDAO.ID = value; }

    public string Title { get => ItemDAO.Title; set => ItemDAO.Title = value; }
    public string ContentTypeId { get => ItemDAO.ContentTypeId; set => ItemDAO.ContentTypeId = value; }
    public string FileRef { get => ItemDAO.FileRef; }

    public T ItemDAO { get; set; }
    public U JsonDAO { get; set; }

    public JsonFileListItemAbstractionDTO()
    {
        ItemDAO = new T();
        JsonDAO = new U();
    }

    public JsonFileListItemAbstractionDTO(T itemDAO, U jsonDAO)
    {
        ItemDAO = itemDAO;
        JsonDAO = jsonDAO;
    }
}

public abstract class JsonFileAbstractionDTO<U>
    where U : BaseJsonDataDAO, new()
{
    public U JsonDAO { get; set; }

    public JsonFileAbstractionDTO()
    {
        JsonDAO = new U();
    }

    public JsonFileAbstractionDTO(U jsonDAO)
    {
        JsonDAO = jsonDAO;
    }
}