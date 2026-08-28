using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Azure.Functions.Worker.Middleware;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Net;

namespace Contoso.Portal.Runtime.Middlewares
{
    internal sealed class ExceptionHandlingMiddleware : IFunctionsWorkerMiddleware
    {
        private readonly ILogger<ExceptionHandlingMiddleware> _logger;

        public ExceptionHandlingMiddleware(ILogger<ExceptionHandlingMiddleware> logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public async Task Invoke(FunctionContext context, FunctionExecutionDelegate next)
        {
            try
            {
                await next(context);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unhandled exception in {FunctionName}", context.FunctionDefinition.Name);

                var errorResponse = new { Message = ex.Message, Details = ex.StackTrace };
                // JsonConvert.SerializeObject returns a string (no stream write) → WriteStringAsync writes async
                // This avoids "Synchronous operations are disallowed" from Kestrel
                var json = JsonConvert.SerializeObject(errorResponse);

                var http = context.GetHtttestsponseData();
                if (http is not null)
                {
                    http.StatusCode = HttpStatusCode.InternalServerError;
                    await http.WriteStringAsync(json);
                    return;
                }

                var req = await context.GetHtttestquestDataAsync();
                if (req is null) return;

                var newResponse = req.CreateResponse(HttpStatusCode.InternalServerError);
                await newResponse.WriteStringAsync(json);

                var outputBinding = context.GetOutputBindings<HtttestsponseData>()
                    .FirstOrDefault(b => b.BindingType == "http" && b.Name != "$return");

                if (outputBinding is not null)
                    outputBinding.Value = newResponse;
                else
                    context.GetInvocationResult().Value = newResponse;
            }
        }
    }
}
