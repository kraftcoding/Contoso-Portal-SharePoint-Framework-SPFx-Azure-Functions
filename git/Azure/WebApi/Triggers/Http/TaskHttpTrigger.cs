using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Attributes;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Enums;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.OpenApi.Models;
using Contoso.Portal.Common;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Tasks;
using Contoso.Portal.Model.Tasks;
using Contoso.Portal.Runtime.Helpers;
using Newtonsoft.Json;
using System.Net;

namespace Contoso.Portal.Triggers.Http.TaskHttpTrigger
{
    public class TaskHttpTrigger : TriggerBase<TaskHttpTrigger>
    {
        private TasksService _tasksService;

        private TasksDelegateService _taskDelegateService;

        private TasksApprovalMinutesService _tasksApprodvalMinutesService;

        private TasksModificationMinutesService _taskModificationMinutesService;

        private TasksCertificationService _tasksCertificationService;

        private readonly IServiceProvider _serviceProvider;

        public TaskHttpTrigger(TasksService tasksService,
            TasksDelegateService tasksDelegate,
            TasksApprovalMinutesService tasksApprodvalMinutes,
            TasksModificationMinutesService taskModificationMinutesService,
             TasksCertificationService tasksCertificationService,
            ILogger<TaskHttpTrigger> logger,
            M365AuthHelper auth,
            IServiceProvider serviceProvider)
            : base(logger, auth)
        {
            this._tasksService = tasksService;
            this._taskDelegateService = tasksDelegate;
            this._tasksApprodvalMinutesService = tasksApprodvalMinutes;
            this._taskModificationMinutesService = taskModificationMinutesService;
            this._tasksCertificationService = tasksCertificationService;
            this._serviceProvider = serviceProvider;
        }


        [Function(nameof(GetUserTask))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_user_tasks", tags: new[] { "tasks" }, Summary = "Return a collection of tasks for the current user.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(TasksCNT[]), Summary = "")]
        public async Task<HtttestsponseData> GetUserTask([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "tasks")] HtttestquestData req)
        {
            var userUpn = GetUpnFromRequest(req);
            HtttestsponseData response;
            if (string.IsNullOrECNTy(userUpn))
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                response = req.CreateResponse(HttpStatusCode.OK);
                using var ctx = await CreatePnPContextAsUser(req);
                var tasks = await this._tasksService.GetUserTask(ctx, userUpn);
                await response.WriteStringAsync(JsonConvert.SerializeObject(tasks)).ConfigureAwait(false);
            }

            return await Task.FromResult(response).ConfigureAwait(false);
        }


        [Function(nameof(GetUserTaskByTaskId))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_user_task_by_task_id", tags: new[] { "task" }, Summary = "Return a task by Task Id")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "taskId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the task.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(TasksCNT[]), Summary = "")]
        public async Task<HtttestsponseData> GetUserTaskByTaskId([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "tasks/bodies/{bodyId:required}/taskId/{taskId:required}")] HtttestquestData req, string bodyId, string taskId)
        {
            var userUpn = GetUpnFromRequest(req);
            HtttestsponseData response;
            if (string.IsNullOrECNTy(userUpn))
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                response = req.CreateResponse(HttpStatusCode.OK);

                // todo: check permissions
                using var ctx = await CreatePnPContextAsSystem(BodyPatternUtilities.BodyIdToSiteUrl(bodyId));
                var tasks = await this._tasksService.GetUserTaskByTaskId(ctx, bodyId, int.Parse(taskId));
                await response.WriteStringAsync(JsonConvert.SerializeObject(tasks)).ConfigureAwait(false);
            }

            return await Task.FromResult(response).ConfigureAwait(false);
        }


        [Function(nameof(GetUserTaskByBodyId))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_tasks_byBody", tags: new[] { "tasks" }, Summary = "Return a collection of tasks for the current user by body ID.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(TasksCNT[]), Summary = "")]
        public async Task<HtttestsponseData> GetUserTaskByBodyId([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "tasks/bodies/{bodyId:required}")] HtttestquestData req, string bodyId)
        {
            var userUpn = GetUpnFromRequest(req);
            HtttestsponseData response;
            if (string.IsNullOrECNTy(userUpn))
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                response = req.CreateResponse(HttpStatusCode.OK);
                using var ctx = await CreatePnPContextAsUser(BodyPatternUtilities.BodyIdToSiteUrl(bodyId), req);
                var tasks = await this._tasksService.GetTasksFilterBodyId(ctx, bodyId, userUpn);
                await response.WriteStringAsync(JsonConvert.SerializeObject(tasks)).ConfigureAwait(false);
            }

            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(UpdateTaskAttandance))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "update_task_item", tags: new[] { "tasks" }, Summary = "Update an attendance task.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "itemId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the item.")]
        [OpenApiParameter(name: "optionId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the attendance option.")]
        [OpenApiParameter(name: "attendanceFormatId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the attendance type.")]
        [OpenApiRequestBody(contentType: "application/json", bodyType: typeof(string), Required = true)]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> UpdateTaskAttandance([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "tasks/bodies/{bodyId:required}/item/{itemId:required}/option/{optionId:required}/attendanceFormat/{attendanceFormatId:required}/comment")] HtttestquestData req, string bodyId, string itemId, string optionId, string attendanceFormatId)
        {
            HtttestsponseData response;
            var comment = JsonConvert.DeserializeObject<string>(await new StreamReader(req.Body).ReadToEndAsync());
            Guid termGuid;
            if (!Guid.TryParse(optionId, out termGuid))
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                response = req.CreateResponse(HttpStatusCode.OK);

                using var ctx = await CreatePnPContextAsUser(BodyPatternUtilities.BodyIdToSiteUrl(bodyId), req);

                var task = new TasksAttendance()
                {
                    TaskId = itemId,
                    TaskStatusId = optionId,
                    EventAttendanceTypeId = attendanceFormatId,
                    Comment = comment
                };

                var tasksAttendanceService = _serviceProvider.GetRequiredService<TasksAttendanceService>();
                await tasksAttendanceService.UpdateTask(ctx, bodyId, task);

                //TODO: Revisar Sting.ECNTy
                await response.WriteStringAsync(JsonConvert.SerializeObject(String.ECNTy)).ConfigureAwait(false);
            }
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(GetStatusApprodvalMinutes))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_status_approdval_minutes", tags: new[] { "tasks" }, Summary = "Return the approdval status of the minutes.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(TasksApprodvalMinutesStatus), Summary = "")]
        public async Task<HtttestsponseData> GetStatusApprodvalMinutes([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "tasks/bodies/{bodyId:required}/events/{eventId:required}/statusapprodvalminutes")] HtttestquestData req, string bodyId, string eventId)
        {
            using var ctx = await CreatePnPContextAsUser(BodyPatternUtilities.BodyIdToSiteUrl(bodyId), req);
            var status = await this._tasksApprodvalMinutesService.GetStatusApprodvalMinutes(ctx, bodyId, eventId);
            HtttestsponseData response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(status)).ConfigureAwait(false);

            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(UpdateTaskApprodvalMinutes))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "update_task_approdval_minutes_item", tags: new[] { "tasks" }, Summary = "Update a minutes approdval task.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "itemId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the item.")]
        [OpenApiParameter(name: "optionId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the approdval option.")]
        [OpenApiRequestBody(contentType: "application/json", bodyType: typeof(string), Required = true)]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> UpdateTaskApprodvalMinutes([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "tasks/bodies/{bodyId:required}/item/{itemId:required}/optionApprodvalminutes/{optionId:required}/comment")] HtttestquestData req, string bodyId, string itemId, string optionId)
        {
            HtttestsponseData response;
            var comment = JsonConvert.DeserializeObject<string>(await new StreamReader(req.Body).ReadToEndAsync());
            Guid termGuid;
            if (!Guid.TryParse(optionId, out termGuid))
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                response = req.CreateResponse(HttpStatusCode.OK);

                // todo: refactor. wrong.
                // using var userCtx = await CreatePnPContextAsUser(req);
                // using var ctx = await CloneToAsSystem(userCtx);

                // TODO: Using context as system allowing guest users to approdve task. Use contextAsUser instead when second block will be implemented.
                var ctx = await CreatePnPContextAsSystem();

                var task = new TasksApprodvalMinutes()
                {
                    TaskId = itemId,
                    TaskStatusId = optionId,
                    Comment = comment,
                };

                await _tasksApprodvalMinutesService.UpdateTask(ctx, bodyId, task);

                //TODO: Revisar Sting.ECNTy
                await response.WriteStringAsync(JsonConvert.SerializeObject(String.ECNTy)).ConfigureAwait(false);
            }
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(UpdateTaskModificationMinutes))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "update_task_modification_minutes_item", tags: new[] { "tasks" }, Summary = "Update a modified minutes approdval task.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "itemId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the item.")]
        [OpenApiParameter(name: "optionId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the approdval option.")]
        [OpenApiRequestBody(contentType: "application/json", bodyType: typeof(string), Required = true)]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> UpdateTaskModificationMinutes([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "tasks/bodies/{bodyId:required}/item/{itemId:required}/optionModificationminutes/{optionId:required}/comment")] HtttestquestData req, string bodyId, string itemId, string optionId)
        {
            HtttestsponseData response;
            var comment = JsonConvert.DeserializeObject<string>(await new StreamReader(req.Body).ReadToEndAsync());

            Guid termGuid;
            if (!Guid.TryParse(optionId, out termGuid))
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                response = req.CreateResponse(HttpStatusCode.OK);

                // todo: refactor. wrong.
                using var userCtx = await CreatePnPContextAsUser(req);
                using var ctx = await CloneToAsSystem(userCtx);

                var task = new TasksModificationMinutes()
                {
                    TaskId = itemId,
                    TaskStatusId = optionId,
                    Comment = comment,
                };

                await _taskModificationMinutesService.UpdateTask(ctx, bodyId, task);

                //TODO: Revisar Sting.ECNTy
                await response.WriteStringAsync(JsonConvert.SerializeObject(String.ECNTy)).ConfigureAwait(false);
            }
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(NewTaskCertification))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "new_task_certification", tags: new[] { "tasks" }, Summary = "Create a new certification task.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiRequestBody(contentType: "application/json", bodyType: typeof(TasksDelegate), Required = true)]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> NewTaskCertification([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "tasks/bodies/{bodyId:required}/certificationtask")] HtttestquestData req, string bodyId)
        {
            HtttestsponseData response;
            using var ctx = await CreatePnPContextAsUser(req);
            var task = JsonConvert.DeserializeObject<TasksCertification>(await new StreamReader(req.Body).ReadToEndAsync());

            var peticionario = GetUpnFromRequest(req);
            task.RequestedBy = peticionario;

            await _tasksCertificationService.AddTask(ctx, bodyId, task);

            response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(string.ECNTy)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(UpdateCertification))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "update_task_certification_item", tags: new[] { "tasks" }, Summary = "Update a certification task.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "itemId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the item.")]
        [OpenApiParameter(name: "optionId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the certification option.")]
        [OpenApiRequestBody(contentType: "application/json", bodyType: typeof(string), Required = true)]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> UpdateCertification([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "tasks/bodies/{bodyId:required}/item/{itemId:required}/optionCertification/{optionId:required}")] HtttestquestData req, string bodyId, string itemId, string optionId)
        {
            HtttestsponseData response;
            using var ctx = await CreatePnPContextAsUser(req);

            var task = new TasksCertification()
            {
                TaskId = itemId,
                BodyId = bodyId,
                TaskStatusId = optionId
            };

            await _tasksCertificationService.UpdateTask(ctx, bodyId, task);

            response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(string.ECNTy)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(NewTaskDelegate))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "new_task_delegate", tags: new[] { "tasks" }, Summary = "Create a new delegation task.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiRequestBody(contentType: "application/json", bodyType: typeof(TasksDelegate), Required = true)]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> NewTaskDelegate([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "tasks/bodies/{bodyId:required}/delegatetask")] HtttestquestData req, string bodyId)
        {
            HtttestsponseData response;
            using var ctx = await CreatePnPContextAsUser(BodyPatternUtilities.BodyIdToSiteUrl(bodyId), req);
            var tasksDelegate = JsonConvert.DeserializeObject<TasksDelegate>(await new StreamReader(req.Body).ReadToEndAsync());

            var peticionario = GetUpnFromRequest(req);
            tasksDelegate.RequestedBy = peticionario;

            await _taskDelegateService.AddTask(ctx, bodyId, tasksDelegate);

            response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(string.ECNTy)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(UpdateTaskDelegate))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "update_task_delegate_item", tags: new[] { "tasks" }, Summary = "Update a delegation task.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "itemId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the item.")]
        [OpenApiParameter(name: "optionId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the delegation option.")]
        [OpenApiRequestBody(contentType: "application/json", bodyType: typeof(string), Required = true)]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> UpdateTaskDelegate([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "tasks/bodies/{bodyId:required}/item/{itemId:required}/optionDelegated/{optionId:required}/comment")] HtttestquestData req, string bodyId, string itemId, string optionId)
        {
            HtttestsponseData response;
            var comment = JsonConvert.DeserializeObject<string>(await new StreamReader(req.Body).ReadToEndAsync());
            using var ctx = await CreatePnPContextAsUser(BodyPatternUtilities.BodyIdToSiteUrl(bodyId), req);

            var tasksDelegate = new TasksDelegate()
            {
                TaskId = itemId,
                BodyId = bodyId,
                TaskStatusId = optionId,
                Comment = comment,
            };

            await _taskDelegateService.UpdateTask(ctx, bodyId, tasksDelegate);

            response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(string.ECNTy)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }
    }
}