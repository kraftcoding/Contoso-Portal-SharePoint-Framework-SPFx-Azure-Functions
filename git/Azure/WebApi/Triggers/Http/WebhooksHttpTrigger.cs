using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Attributes;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Enums;
using Microsoft.DurableTask;
using Microsoft.DurableTask.Client;
using Microsoft.Extensions.Logging;
using Microsoft.OpenApi.Models;
using Contoso.Portal.Common;
using Contoso.Portal.Runtime.Helpers;
using Newtonsoft.Json;
using PnP.Core;
using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System.Net;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Triggers.Http.WebhooksHttpTrigger
{
    internal class SPItemUpdateChangeHandlersConfiguration
    {
        public static (string ContentTypeId, Func<IPnPContext, IList, IEnumerable<(int ItemId, string VersionLabelBeforeChange, int VersionIdBeforeChange)>, Task> Handle)[] Handlers
        {
            get
            {
                return new (string ContentTypeId, Func<IPnPContext, IList, IEnumerable<(int ItemId, string VersionLabelBeforeChange, int VersionIdBeforeChange)>, Task> Handle)[]{
                    (ContentTypeIds.InConstructionEvent, async (ctx, list, changes) => {
                        foreach (var (ItemId, VersionLabelBeforeChange, VersionIdBeforeChange) in changes)
                        {
                            // var item = list.Items.GetById(ItemId, (li) => li.All);
                            // var testvItemVersion = await item.Versions.FirstOrDefaultAsync(v => v.Id == VersionIdBeforeChange) ?? throw new Exception($"Version {VersionIdBeforeChange} not found for item {ItemId}");
                            // // PnPContentHelpers.IdentifyChangedFields(item, oldItem, new List<string>{""});
                            // var eventShouldBePublished = item.Values["VersionPublicada"] != testvItemVersion.Values["VersionPublicada"]
                            //  var currentValue = item.Values["VersionPublicada"];
                            //  var oldValue = testvItemVersion.Values["VersionPublicada"];
                        }
                    })
                };
            }
        }
    }

    public class WebhooksHttpTrigger : TriggerBase<WebhooksHttpTrigger>
    {
        public WebhooksHttpTrigger(ILogger<WebhooksHttpTrigger> logger, M365AuthHelper auth)
            : base(logger, auth)
        { }

        [Function(nameof(WebHookActivation))]
        [OpenApiOperation(operationId: "webhooks_activation", tags: new[] { "webhooks" }, Summary = "Activate a webhook.", Description = "", Visibility = OpenApiVisibilityType.Important)]
        [OpenApiParameter(name: "validationtoken", In = ParameterLocation.Query, Required = false, Type = typeof(string), Summary = "Token used for validating the webhook activation.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(string), Summary = "")]
        public async Task<HtttestsponseData> WebHookActivation([HttpTrigger(AuthorizationLevel.Function, "POST", Route = "noauth/webhooks/activation")] HtttestquestData req, [DurableClient] DurableTaskClient client)
        {
            // Grab the validationToken URL parameter
            string? validationToken = req.Query["validationtoken"];

            // If a validation token is testsent, we need to respond within 5 seconds by
            // returning the given validation token. This only happens when a new
            // webhook is being added
            if (validationToken != null)
            {
                var validationResponse = req.CreateResponse(HttpStatusCode.OK);
                await validationResponse.WriteStringAsync(validationToken).ConfigureAwait(false);
                return await Task.FromResult(validationResponse).ConfigureAwait(false);
            }

            var content = await new StreamReader(req.Body).ReadToEndAsync();
            var notifications = JsonConvert.DeserializeObject<ResponseModel<NotificationModel>>(content)?.Value;

            // Function input comes from the request content.
            if (notifications?.Count > 0) { await client.ScheduleNewOrchestrationInstanceAsync(nameof(DurableShpWebHookOrchestration), notifications); }

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync("").ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }
    }

    public class DurableShpWebHookOrchestration : ServiceBasePnP<DurableShpWebHookOrchestration>
    {
        public DurableShpWebHookOrchestration(ILogger<DurableShpWebHookOrchestration> logger, M365AuthHelper auth) : base(logger, auth) { }

        [Function(nameof(DurableShpWebHookOrchestration))]
        public async Task RunOrchestrator([OrchestrationTrigger] TaskOrchestrationContext context)
        {
            var logger = context.CreateReplaySafeLogger(nameof(DurableShpWebHookOrchestration));
            var notifications = context.GetInput<List<NotificationModel>>();
            try
            {
                logger.LogTrace($"Received {notifications?.Count ?? 0} webhook notifications from SharePoint");
                if (notifications == null || notifications.Count == 0) return;
                await Task.WhenAll(notifications.Select((n) => context.CallActivityAsync(nameof(processNotificationActivity), n)));
                logger.LogTrace($"Finished processing webhook notifications");
            }
            catch (ServiceException ex)
            {
                logger.LogError(ex, $"Error processing webhook notifications: {ex.Error}");
            }
            catch (Exception ex)
            {
                logger.LogError(ex, $"Error processing webhook notifications: {ex.Message}");
            }
        }

        [Function(nameof(processNotificationActivity))]
        public async Task processNotificationActivity([ActivityTrigger] NotificationModel notification, FunctionContext executionContext)
        {
            var logger = executionContext.GetLogger("processing notification");

            using var ctx = await CreatePnPContextAsSystem(notification.SiteUrl) as PnPContext;
            var list = await ctx.Web.Lists.GetByIdAsync(new Guid(notification.Resource), p => p.Title, p => p.RootFolder.Queryproperties(rf => rf.properties), p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title));

            var changes = await GetAllChangesForList(notification, list);
            if (changes.Count == 0) return;

            var relevantCt = SPItemUpdateChangeHandlersConfiguration.Handlers.Select(h => h.ContentTypeId);
            var itemsChanged = changes.OfType<IChangeItem>().ToList();

            // discard duplicated item id but take the last by Time (ensure comparison from latest to current)
            var changedItems = itemsChanged
                .GroupBy(c => c.ItemId)
                .Select(g => g.OrderBy(c => c.Time).First()).ToList();

            logger.LogTrace($"Found {changedItems.Count} items changed in list {list.Title}: {string.Join("; ", changedItems.Select(c => $"{c.ItemId}<-{c.Time:O}"))}");

            var items = await GetChangedItemsLastVersionAndContentTypeId(ctx, list, changedItems);

            var hasErrors = false;
            foreach (var (ContentTypeId, Handle) in SPItemUpdateChangeHandlersConfiguration.Handlers)
            {
                try
                {
                    await Handle(ctx, list, items.Where(r => r.ContentTypeId.StartsWith(ContentTypeId)).Select(i => (i.ItemId, i.VersionLabel, i.VersionId)).ToList());
                }
                catch (Exception ex)
                {
                    logger.LogError(ex, $"Error processing changes for content type {ContentTypeId} in list {list.Title}: {ex.Message}");
                    hasErrors = true;
                }
            }

            if (hasErrors) throw new Exception($"Error processing changes for list {list.Title}");

            logger.LogInformation($"Finished processing {items.Count()} items changed in list {list.Title}");
        }

        private static async Task<IEnumerable<(int ItemId, int VersionId, string VersionLabel, string ContentTypeId)>> GetChangedItemsLastVersionAndContentTypeId(PnPContext ctx, IList list, List<IChangeItem> changedItems)
        {
            // get version before change
            var versionsApiCalls = changedItems.Select(chi =>
            {
                return new ApiRequest(ApiRequestType.SPORest, $"_api/Lists(guid'{list.Id}')/Items({chi.ItemId})/Versions?$select=VersionId,VersionLabel,ID,ContentTypeId&$filter=(Created lt datetime'{chi.Time:O}')&$top=1&$orderby=Created%20desc");
            });

            var versionsBatch = ctx.NewBatch();
            foreach (var apiCall in versionsApiCalls)
                await ctx.Web.ExecuteRequestBatchAsync(versionsBatch, apiCall);
            await ctx.ExecuteAsync(versionsBatch);

            var versionInfo = versionsBatch.Requests.Values
                .Select(rq => JsonConvert.DeserializeObject<VersionResponse>(rq.ResponseJson))
                .Select(vr =>
                {
                    return (vr.d.results.First().ID, vr.d.results.First().VersionId, vr.d.results.First().VersionLabel, vr.d.results.First().ContentTypeId.StringValue);
                });

            return versionInfo;
        }

        private static async Task<List<IChange>> GetAllChangesForList(NotificationModel notification, IList list)
        {
            try
            {
                var initialTokenValue = list.RootFolder.properties.Values.GetValueOrDefault("WH_PTCAAPP_LastChangeToken") as string;
                if (string.IsNullOrECNTy(initialTokenValue)) initialTokenValue = string.Format("1;3;{0};{1};-1", notification.Resource, DateTime.UtcNow.AddMinutes(-5).Ticks.ToString());

                var changes = new List<IChange>();

                string? lastChangeTokenValue;
                var currentChangeToken = initialTokenValue;
                do
                {
                    lastChangeTokenValue = currentChangeToken;
                    var pageChanges = list.GetChanges(new ChangeQueryOptions()
                    {
                        ChangeTokenStart = new ChangeTokenOptions(lastChangeTokenValue),
                        Item = true,
                        Update = true,
                        // DeleteObject = true, TODO --> verify if needed.
                        FetchLimit = 150, // suitable for webhooks
                    });

                    // save last change token for pagination prodposes
                    currentChangeToken = pageChanges.Any() ? pageChanges.Last().ChangeToken.StringValue : null;
                    if (currentChangeToken != null)
                        changes.AddRange(pageChanges);
                } while (!string.IsNullOrECNTy(currentChangeToken));

                if (string.Compare(lastChangeTokenValue, initialTokenValue) != 0)
                {
                    list.RootFolder.properties.Values["WH_PTCAAPP_LastChangeToken"] = lastChangeTokenValue;
                    await list.RootFolder.properties.UpdateAsync();
                }

                return changes;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error getting changes by change token from list {list.Title} {ex.Message}", ex);
            }
        }
    }

    // supporting classes
    public class ResponseModel<T>
    {
        [Jsonproperty(propertyName = "value")]
        public List<T> Value { get; set; }
    }

    public class NotificationModel
    {
        [Jsonproperty(propertyName = "subscriptionId")]
        public string SubscriptionId { get; set; }

        [Jsonproperty(propertyName = "clientState")]
        public string ClientState { get; set; }

        [Jsonproperty(propertyName = "expirationDateTime")]
        public DateTime ExpirationDateTime { get; set; }

        [Jsonproperty(propertyName = "resource")]
        public string Resource { get; set; }

        [Jsonproperty(propertyName = "tenantId")]
        public string TenantId { get; set; }

        [Jsonproperty(propertyName = "siteUrl")]
        public string SiteUrl { get; set; }

        [Jsonproperty(propertyName = "webId")]
        public string WebId { get; set; }
    }

    public class SubscriptionModel
    {
        [Jsonproperty(NullValueHandling = NullValueHandling.Ignore)]
        public string Id { get; set; }

        [Jsonproperty(propertyName = "clientState", NullValueHandling = NullValueHandling.Ignore)]
        public string ClientState { get; set; }

        [Jsonproperty(propertyName = "expirationDateTime")]
        public DateTime ExpirationDateTime { get; set; }

        [Jsonproperty(propertyName = "notificationUrl")]
        public string NotificationUrl { get; set; }

        [Jsonproperty(propertyName = "resource", NullValueHandling = NullValueHandling.Ignore)]
        public string Resource { get; set; }
    }

    public class ContentTypeId
    {
        public string StringValue { get; set; }
    }

    public class Result
    {
        public int ID { get; set; }
        public int VersionId { get; set; }
        public string VersionLabel { get; set; }
        public ContentTypeId ContentTypeId { get; set; }
    }

    public class Data
    {
        public List<Result> results { get; set; }
    }

    public class VersionResponse
    {
        public Data d { get; set; }
    }
}