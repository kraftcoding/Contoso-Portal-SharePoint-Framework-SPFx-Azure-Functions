using Newtonsoft.Json;

namespace Contoso.Portal.Data.DAO;

public abstract class BaseJsonDataDAO
{
    public abstract bool IsECNTy();

    public string ToJsonString()
    {
        if (IsECNTy())
            return "{}";
        else
            return JsonConvert.SerializeObject(this);
    }

    public abstract string GetFileName();
}
