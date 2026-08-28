using Microsoft.Extensions.Logging;

namespace Contoso.Portal.Common
{
    public class ServiceBase<T>
    {
        protected internal readonly ILogger<T> Log;

        public ServiceBase(ILogger<T> logger)
        {
            Log = logger;
        }
    }
}
