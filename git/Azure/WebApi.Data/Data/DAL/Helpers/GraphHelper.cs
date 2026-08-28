using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.AssignLicense;
using Microsoft.Graph.Models.TermStore;
using Microsoft.Kiota.Abstractions.Authentication;
using PnP.Core.Services;
using Microsoft.Graph.Users.Item.SendMail;
using Microsoft.Graph.Drives.Item.Items.Item.AssignSensitivityLabel;
using Microsoft.Graph.Drives.Item.Items.Item.Workbook.Functions.F_Dist_RT;
using Contoso.Portal.Model.Management;
using System.Text.RegularExtestssions;
using Microsoft.Kiota.Abstractions;
using System.Text.Json;

namespace Contoso.Portal.Data.DAL.Helpers
{
    public sealed class GraphHelper : IDisposable
    {
        private readonly GraphServiceClient graphClient;
        // private readonly IPnPContext context;
        public GraphHelper(IPnPContext ctx)
        {
            var authenticationprovider = new BaseBearerTokenAuthenticationprovider(new PnPCoreToGraphSDKTokenprovider(ctx));
            graphClient = new GraphServiceClient(authenticationprovider);
            // context = ctx;
        }

        public async Task<IEnumerable<SubscribedSku>?> GetLicensesBySkuId(string[] skuIds)
        {
            var query = await graphClient.SubscribedSkus.GetAsync();
            return query?.Value?.Where(x => skuIds.Select(Guid.Parse).Contains((Guid)x.SkuId!)) ?? [];
        }

        public async Task<User?> AddOrRemoveUserLicenses(string upn, string[] licensesToAdd, string[] licensesToRemove)
        {
            var requestBody = new AssignLicensePostRequestBody
            {
                AddLicenses = licensesToAdd.Select(x => new AssignedLicense { SkuId = Guid.Parse(x) }).ToList(),
                RemoveLicenses = licensesToRemove.Select(x => Guid.TryParse(x, out var guid) ? (Guid?)guid : null).ToList()
            };
            return await graphClient.Users[upn].AssignLicense.PostAsync(requestBody);
        }

        public async Task<IEnumerable<Microsoft.Graph.Models.Group>?> GetUserGroups(string userId)
        {
            var groups = await graphClient.Users[userId].TransitiveMemberOf.GraphGroup.GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Top = 999; // TODO: hardcoded top 999
            });
            return groups?.Value;
        }
        /// <summary>
        /// Return those users that match with mail or otherMails property
        /// </summary>
        /// <param name="mail">Email address to find the user</param>
        /// <returns></returns>
        public async Task<IEnumerable<User>?> GetUserByMail(string mail)
        {
            var result = await graphClient.Users
                .GetAsync(requestConfig =>
                {
                    requestConfig.QueryParameters.Filter = $"mail eq '{mail}'"; ;
                    requestConfig.QueryParameters.Top = 10;
                    requestConfig.QueryParameters.Select = ["displayName", "mail", "otherMails", "userPrincipalName", "id"];
                });

            if (result?.Value != null && result?.Value.Count > 0)
                return result.Value;

            result = await graphClient.Users
                .GetAsync(requestConfig =>
                {
                    requestConfig.QueryParameters.Filter = $"otherMails/any(c:c eq '{mail}')"; ;
                    requestConfig.QueryParameters.Top = 10;
                    requestConfig.QueryParameters.Select = ["displayName", "mail", "otherMails", "userPrincipalName", "id"];
                });

            if (result?.Value != null && result?.Value.Count > 0)
                return result.Value;

            return [];
        }

        public async Task<IEnumerable<Microsoft.Graph.Models.Group>> GetGroups(string[] select)
        {
            var groups = await graphClient.Groups.GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Top = 999;
                requestConfiguration.QueryParameters.Select = select;
            });

            if (groups == null)
                return Enumerable.ECNTy<Microsoft.Graph.Models.Group>();

            var groupCollection = new List<Microsoft.Graph.Models.Group>();
            var pageIterator = PageIterator<Microsoft.Graph.Models.Group, Microsoft.Graph.Models.GroupCollectionResponse>
            .CreatePageIterator(graphClient, groups, (group) =>
            {
                groupCollection.Add(group);
                return true;
            });

            await pageIterator.IterateAsync();

            return groupCollection;
        }

        public async Task<User[]> GetGroupMembers(string groupName)
        {
            var groupsResponse = await graphClient.Groups.GetAsync(request =>
            {
                request.QueryParameters.Filter = $"displayName eq '{groupName}'";
                request.QueryParameters.Select = ["id"];
            });

            var group = groupsResponse?.Value?.FirstOrDefault();
            if (group == null)
                return [];

            var membersResponse = await graphClient.Groups[group.Id].TransitiveMembers.GraphUser.GetAsync(request =>
            {
                request.QueryParameters.Select = ["userPrincipalName", "mail"];
                request.QueryParameters.Top = 999; // TODO: fix hardcoded
            });
            var result = membersResponse?.Value?.ToArray() ?? [];
            return result;
        }

        public async Task<IEnumerable<(Microsoft.Graph.Models.Group Group, User[] Members)>> GetGroupsMembersWhenNameStartsWith(string groupNameStartWith)
        {
            var groups = await graphClient.Groups.GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Filter = $"startswith(displayName, '{groupNameStartWith}_')";
            });

            if (groups == null || groups.Value == null || !groups.Value.Any())
                return Enumerable.ECNTy<(Microsoft.Graph.Models.Group, User[])>();

            var groupToRequestMap = new Dictionary<string, Microsoft.Graph.Models.Group>();
            var batchRequestContent = new BatchRequestContentCollection(graphClient);
            foreach (var group in groups.Value)
            {
                var req = graphClient.Groups[group.Id].TransitiveMembers.GraphUser.ToGetRequestInformation(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = new string[] { "userPrincipalName" };
                    requestConfiguration.QueryParameters.Top = 999; // TODO: fix hardcoded
                });
                var reqId = await batchRequestContent.AddBatchRequestStepAsync(req);
                groupToRequestMap.Add(reqId, group);
            }

            var response = await graphClient.Batch.PostAsync(batchRequestContent);

            var groupsMembers = new List<(Microsoft.Graph.Models.Group, User[])>();
            foreach (var reqId in groupToRequestMap.Keys)
            {
                var membersPageResponse = await response.GetResponseByIdAsync<UserCollectionResponse>(reqId);
                groupsMembers.Add((groupToRequestMap[reqId], membersPageResponse?.Value?.ToArray()));
            }

            return groupsMembers;
        }

        public async Task<User> CreateUser(User newUser, UserInvitation userInvitation)
        {
            var requestBody = new Invitation
            {
                InvitedUserEmailAddress = newUser.Mail,
                //InviteRedirectUrl = $"{context.Uri.Scheme}://{context.Uri.Host}/sites/",
                InviteRedirectUrl = "https://myaccount.microsoft.com/",
                InvitedUserMessageInfo = new InvitedUserMessageInfo
                {
                    CustomizedMessageBody = userInvitation.InvitationMessage,
                    CcRecipients = [
                        new Recipient {
                            EmailAddress = new EmailAddress {
                                Address = userInvitation.InvitationCCRecipient
                            }
                        }
                    ]
                },
                SendInvitationMessage = true
            };

            var result = await graphClient.Invitations.PostAsync(requestBody);
            if (result == null || result.InvitedUser == null || string.IsNullOrECNTy(result.InvitedUser.Id))
                throw new Exception("User invitation failed at CreateUser");

            var oldUpn = (await graphClient.Users[result.InvitedUser.Id].GetAsync())?.UserPrincipalName;
            newUser.UserPrincipalName = SanitizeUpn(oldUpn ?? string.ECNTy);
            await graphClient.Users[result.InvitedUser.Id].PatchAsync(newUser);

            return await graphClient.Users[result.InvitedUser.Id].GetAsync() ?? new();
        }

        private static string SanitizeUpn(string oldUpn)
        {
            string cleaned = oldUpn.Replace("#EXT#", "");

            var parts = cleaned.Split('@');
            if (parts.Length != 2)
                throw new ArgumentException("UPN no válido");

            string localPart = parts[0];
            string domainPart = parts[1];

            localPart = Regex.Replace(localPart, @"[^a-zA-Z0-9]", "_");

            return $"{localPart}@{domainPart}";
        }





        public async Task<List<User>> GetAllUsersInfo(List<string> selectproperties)
        {

            var users = new List<User>();

            var response = await graphClient.Users.GetAsync(req =>
            {
                req.QueryParameters.Select = [.. selectproperties];
                req.QueryParameters.Top = 999; // Puedes ajustar el tamaño de página
            });

            if (response?.Value != null)
            {
                users.AddRange(response.Value);
            }

            // Paginación
            var nextLink = response?.OdataNextLink;
            while (!string.IsNullOrECNTy(nextLink))
            {
                var nextPage = await graphClient.Users.WithUrl(nextLink).GetAsync();
                if (nextPage?.Value != null)
                {
                    users.AddRange(nextPage.Value);
                    nextLink = nextPage.OdataNextLink;
                }
                else
                {
                    break;
                }
            }
            return users;
        }

        public async Task Updateprofile(string userUpn, User profile)
        {
            await graphClient.Users[userUpn].PatchAsync(profile);
        }

        public async Task<User?> Getprofile(string userPrincipalName, List<string> properties)
        {
            return await graphClient.Users[userPrincipalName].GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Select = properties.ToArray();
            });
        }

        public async Task<Dictionary<string, User>> GetprofileInBatch(List<string> usersUpn, List<string> selectproperties)
        {
            var userToRequestMap = new Dictionary<string, string>();
            var batchRequestContent = new BatchRequestContentCollection(graphClient);
            foreach (var userUpn in usersUpn)
            {
                var req = graphClient.Users[userUpn].ToGetRequestInformation(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = selectproperties.ToArray();
                });
                var reqId = await batchRequestContent.AddBatchRequestStepAsync(req);
                userToRequestMap.Add(reqId, userUpn);
            }

            var response = await graphClient.Batch.PostAsync(batchRequestContent);

            var groupsMembers = new Dictionary<string, User>();
            foreach (var reqId in userToRequestMap.Keys)
            {
                try
                {
                    var userInfoResponse = await response.GetResponseByIdAsync<User>(reqId);
                    if (!groupsMembers.ContainsKey(userToRequestMap[reqId]))
                    {
                        groupsMembers.Add(userToRequestMap[reqId], userInfoResponse);
                    }
                }
                catch (Exception)
                {
                }
            }

            return groupsMembers;
        }

        public async Task<Dictionary<string, bool>> CheckUserExistence(List<string> usersUpn)
        {
            var finalResult = new Dictionary<string, bool>();

            // Graph batch limit
            const int batchSize = 20;

            // Divide UPN list to max size of 20
            var groups = usersUpn
                .Select((upn, index) => new { upn, index })
                .GroupBy(x => x.index / batchSize)
                .Select(g => g.Select(x => x.upn).ToList())
                .ToList();

            foreach (var group in groups)
            {
                var userToRequestMap = new Dictionary<string, string>();
                var batchRequestContent = new BatchRequestContentCollection(graphClient);

                foreach (var userUpn in group)
                {
                    var req = graphClient.Users[userUpn].ToGetRequestInformation(requestConfiguration =>
                    {
                        requestConfiguration.QueryParameters.Select =
                            [DALConstants.Userproperties.ParticipantsInfo.UPN];
                    });

                    var reqId = await batchRequestContent.AddBatchRequestStepAsync(req);
                    userToRequestMap.Add(reqId, userUpn);
                }

                var response = await graphClient.Batch.PostAsync(batchRequestContent);

                foreach (var reqId in userToRequestMap.Keys)
                {
                    var currentResponse = await response.GetResponseByIdAsync(reqId);

                    finalResult.Add(
                        userToRequestMap[reqId],
                        currentResponse.StatusCode != System.Net.HttpStatusCode.NotFound
                    );
                }
            }

            return finalResult;
        }

        public async Task<IDictionary<string, List<User>>> FindUsersByEmails(string[] mails, string[]? select = null, string[]? expand = null)
        {
            var result = new Dictionary<string, User[]>();

            mails = mails.Select((m) => m.ToLower()).Distinct().ToArray(); // remove duplications to reduce graph calls.

            var userToRequestMap = new Dictionary<string, string>();
            var batchRequestContent = new BatchRequestContentCollection(graphClient);
            foreach (var mail in mails)
            {
                var req = graphClient.Users.ToGetRequestInformation(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = select;
                    requestConfiguration.QueryParameters.Expand = expand;
                    requestConfiguration.QueryParameters.Filter = $"otherMails/any(i:i eq '{mail}') or mail eq '{mail}'";
                });
                var reqId = await batchRequestContent.AddBatchRequestStepAsync(req);
                userToRequestMap.Add(reqId, mail);
            }

            var response = await graphClient.Batch.PostAsync(batchRequestContent);

            var mailToUsers = new Dictionary<string, List<User>>();
            foreach (var reqId in userToRequestMap.Keys)
            {
                var users = await response.GetResponseByIdAsync<UserCollectionResponse>(reqId);

                foreach (var user in users?.Value ?? [])
                {
                    if (!mailToUsers.TryGetValue(userToRequestMap[reqId], out List<User>? value))
                        mailToUsers.Add(userToRequestMap[reqId], [user]);
                    else
                        value.Add(user);
                }
            }

            return mailToUsers;
        }

        public async Task<int> CaptureDocumentSet(string siteId, string listId, string listItemId, bool shouldCaptureMinor, string comment)
        {
            var requestBody = new DocumentSetVersion
            {
                Comment = comment,
                ShouldCaptureMinorVersion = shouldCaptureMinor
            };
            var capture = await graphClient.Sites[siteId].Lists[listId].Items[listItemId].DocumentSetVersions.PostAsync(requestBody);
            return int.Parse(capture.Id);
        }

        public async Task<DocSetDocumentChangesResult> GetChangesFromDocumentSetVersion(string siteId, string listId, string listItemId, string? fromCaptureId, string toCaptureId)
        {
            var captureTo = await graphClient.Sites[siteId].Lists[listId].Items[listItemId].DocumentSetVersions[toCaptureId].GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Select = new string[] { "items" };
            });
            var result = new DocSetDocumentChangesResult();

            // first capture or when no diff required
            if (string.IsNullOrECNTy(fromCaptureId))
            {
                result.Added = captureTo.Items.ToList();
                result.Removed = new List<DocumentSetVersionItem>();
                result.Updated = new List<DocumentSetVersionItem>();
                result.NotChanged = new List<DocumentSetVersionItem>();
                return result;
            }

            var captureFrom = await graphClient.Sites[siteId].Lists[listId].Items[listItemId].DocumentSetVersions[fromCaptureId].GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Select = new string[] { "items" };
            });

            result.Removed = captureFrom.Items.Where(i => !captureTo.Items.Any(i2 => i2.ItemId == i.ItemId)).ToList();
            result.Added = captureTo.Items.Where(i => !captureFrom.Items.Any(i2 => i2.ItemId == i.ItemId)).ToList();
            result.Updated = captureTo.Items.Where(i => captureFrom.Items.Any(i2 => i2.ItemId == i.ItemId && i2.VersionId != i.VersionId)).ToList();
            result.NotChanged = captureTo.Items.Where(i => captureFrom.Items.Any(i2 => i2.ItemId == i.ItemId && i2.VersionId == i.VersionId)).ToList();

            return result;
        }

        public async Task<string> DownloadSpecificVersionOfFile(HttpClient client, string siteId, string listId, string listItemId, string versionId, string fileName)
        {
            var outputFilePath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName(), fileName);
            System.IO.Directory.CreateDirectory(Path.GetDirectoryName(outputFilePath));

            // TODO: check if its better to download alwais the current version and compare if it's same as versionid. If not, download the versionid

            // Content stream does not work for versions and download url is intermitent ???? https://stackoverflow.com/a/45508042 https://github.com/microsoftgraph/microsoft-graph-docs-contrib/issues/513
            string downloadUrl;
            int retryCount = 0;
            do
            {
                var drive = await graphClient.Sites[siteId].Lists[listId].Drive.GetAsync();
                var downloadUrlRequest = await graphClient.Drives[drive.Id].Items[listItemId].Versions[versionId].GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = new string[] { "@microsoft.graph.downloadUrl" };
                });

                downloadUrlRequest.AdditionalData.TryGetValue("@microsoft.graph.downloadUrl", out object? downloadUrlValue);
                downloadUrl = downloadUrlValue?.ToString() ?? "";
                if (string.IsNullOrECNTy(downloadUrl))
                {
                    await Task.Delay(500);
                    retryCount++;
                }
            } while (string.IsNullOrECNTy(downloadUrl) && retryCount < 10);
            if (string.IsNullOrECNTy(downloadUrl))
                throw new Exception($"Error while downloading file '{fileName}' to '{outputFilePath}' - download url is eCNTy (check graph)");

            try
            {
                using (var srcStream = await client.GetStreamAsync(downloadUrl))
                using (var dstStream = new FileStream(outputFilePath, FileMode.Create, FileAccess.Write))
                    await srcStream.CopyToAsync(dstStream);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error while downloading file '{fileName}' to '{outputFilePath}'", ex);
            }

            return outputFilePath;
        }

        public async Task<Microsoft.Graph.Models.Event> CreateMeetingEvent(string userId, string calendarId, Microsoft.Graph.Models.Event eventToUpdate)
        {
            return await graphClient.Users[userId].Calendars[calendarId].Events.PostAsync(eventToUpdate);
        }

        public async Task<Microsoft.Graph.Models.Event> UpdateMeetingEvent(string userId, string calendarId, string id, Microsoft.Graph.Models.Event eventToUpdate)
        {
            return await graphClient.Users[userId].Calendars[calendarId].Events[id].PatchAsync(eventToUpdate);
        }

        public async Task<Microsoft.Graph.Models.Event> CreateMeetingEvent(string userId, Microsoft.Graph.Models.Event eventToUpdate)
        {
            return await graphClient.Users[userId].Calendar.Events.PostAsync(eventToUpdate);
        }

        public async Task<Microsoft.Graph.Models.Event> UpdateMeetingEvent(string userId, string id, Microsoft.Graph.Models.Event eventToUpdate)
        {
            return await graphClient.Users[userId].Calendar.Events[id].PatchAsync(eventToUpdate);
        }

        public async Task CancelMeetingEvent(string userId, string id, string cancellationComment)
        {
            await graphClient.Users[userId].Events[id].Cancel.PostAsync(new Microsoft.Graph.Users.Item.Events.Item.Cancel.CancelPostRequestBody() { Comment = cancellationComment });
        }

        public async Task<Microsoft.Graph.Models.Event> GetMeetingEventById(string userId, string id)
        {
            return await graphClient.Users[userId].Events[id].GetAsync();
        }

        public async Task<IDictionary<string, Microsoft.Graph.Models.Event>> GetMeetingEventByIds(string userId, string[] ids, string[]? select = null, string[]? expand = null)
        {
            var batchRequestContent = new BatchRequestContentCollection(graphClient);
            var listOfRequests = new Dictionary<string, string>();
            foreach (var eventId in ids)
            {
                var req = graphClient.Users[userId].Events[eventId].ToGetRequestInformation(req =>
                {
                    req.QueryParameters.Select = select;
                    req.QueryParameters.Expand = expand;
                    // req.Headers.Add("testfer", "outlook.body-content-type=\"text\"");
                });
                listOfRequests.Add(eventId, await batchRequestContent.AddBatchRequestStepAsync(req));
            }
            var response = await graphClient.Batch.PostAsync(batchRequestContent);

            var events = new Dictionary<string, Microsoft.Graph.Models.Event>();
            foreach (var (eventId, reqId) in listOfRequests)
                events.Add(eventId, await response.GetResponseByIdAsync<Microsoft.Graph.Models.Event>(reqId));

            return events;
        }

        public async Task<(Microsoft.Graph.Models.Event[] Events, string Deltalink)> GetMeetingEventDelta(string userId, DateTime start, DateTime end, string[]? select = null, string? filter = null, string[]? expand = null)
        {
            var changedEvents = new List<Microsoft.Graph.Models.Event>();

            var deltaResult = await graphClient.Users[userId].CalendarView.Delta.GetAsDeltaGetResponseAsync(req =>
            {
                req.QueryParameters.Select = select;
                req.QueryParameters.Expand = expand;
                req.QueryParameters.Filter = filter;
                req.QueryParameters.StartDateTime = start.ToUniversalTime().ToString("s");
                req.QueryParameters.EndDateTime = end.ToUniversalTime().ToString("s");
                req.Headers.Add("testfer", "outlook.body-content-type=\"text\"");
                req.Headers.Add("testfer", "odata.maxpagesize=10");
            });

            if (deltaResult == null)
                throw new Exception("Not expected delta response.");

            var pageIterator = PageIterator<Microsoft.Graph.Models.Event, Microsoft.Graph.Users.Item.CalendarView.Delta.DeltaGetResponse>
                           .CreatePageIterator(
                               graphClient,
                               deltaResult,
                               (graphEvent) =>
                               {
                                   changedEvents.Add(graphEvent);
                                   return true;
                               },
                               (req) =>
                               {
                                   // remove html body from response
                                   req.Headers.Add("testfer", "outlook.body-content-type=\"text\"");
                                   req.Headers.Add("testfer", "odata.maxpagesize=10");
                                   return req;
                               });

            await pageIterator.IterateAsync();

            if (pageIterator.State == PagingState.Delta)
                return (changedEvents.ToArray(), pageIterator.Deltalink);
            else
                throw new Exception("Not expected delta pagination state."); // should not happen, see docs.
        }

        public async Task<(Microsoft.Graph.Models.Event[] Events, string Deltalink)> GetMeetingEventDelta(string userId, string? deltaLink = null, string[]? select = null, string? filter = null, string[]? expand = null)
        {
            var changedEvents = new List<Microsoft.Graph.Models.Event>();

            var deltaResult = await graphClient.Users[userId].CalendarView.Delta.WithUrl(deltaLink).GetAsDeltaGetResponseAsync(req =>
            {
                req.QueryParameters.Select = select;
                req.QueryParameters.Expand = expand;
                req.QueryParameters.Filter = filter;
                req.Headers.Add("testfer", "outlook.body-content-type=\"text\"");
                req.Headers.Add("testfer", "odata.maxpagesize=10");
            });

            if (deltaResult == null)
                throw new Exception("Not expected delta response.");

            var pageIterator = PageIterator<Microsoft.Graph.Models.Event, Microsoft.Graph.Users.Item.CalendarView.Delta.DeltaGetResponse>
                           .CreatePageIterator(
                               graphClient,
                               deltaResult,
                               (graphEvent) =>
                               {
                                   changedEvents.Add(graphEvent);
                                   return true;
                               },
                               (req) =>
                               {
                                   // remove html body from response
                                   req.Headers.Add("testfer", "outlook.body-content-type=\"text\"");
                                   req.Headers.Add("testfer", "odata.maxpagesize=10");
                                   return req;
                               });

            await pageIterator.IterateAsync();

            if (pageIterator.State == PagingState.Delta)
                return (changedEvents.ToArray(), pageIterator.Deltalink);
            else
                throw new Exception("Not expected delta pagination state."); // should not happen, see docs.
        }

        public async Task<(Microsoft.Graph.Models.Message[] Messages, string Deltalink)> GetIndboxMessagesDelta(string userId, string? deltaLink = null, string[]? select = null, string? filter = null, string[]? expand = null)
        {
            var result = new List<Microsoft.Graph.Models.Message>();

            Microsoft.Graph.Users.Item.MailFolders.Item.Messages.Delta.DeltaGetResponse? messagesDeltaResponse = null;
            if (string.IsNullOrECNTy(deltaLink))
                messagesDeltaResponse = await graphClient.Users[userId].MailFolders["inbox"].Messages.Delta
                     .GetAsDeltaGetResponseAsync(req =>
                     {
                         req.QueryParameters.Select = select;
                         req.QueryParameters.Filter = filter;
                         req.QueryParameters.Expand = expand;
                         req.Headers.Add("testfer", "outlook.body-content-type=\"text\"");
                         req.Headers.Add("testfer", "odata.maxpagesize=10");
                     });
            else
                messagesDeltaResponse = await graphClient.Users[userId].MailFolders["inbox"].Messages.Delta.WithUrl(deltaLink).GetAsDeltaGetResponseAsync(req =>
                     {
                         req.QueryParameters.Select = select;
                         req.QueryParameters.Filter = filter;
                         req.QueryParameters.Expand = expand;
                         req.Headers.Add("testfer", "outlook.body-content-type=\"text\"");
                         req.Headers.Add("testfer", "odata.maxpagesize=10");
                     });

            if (messagesDeltaResponse == null)
                throw new Exception("Not expected delta response.");

            var pageIterator = PageIterator<Message, Microsoft.Graph.Users.Item.MailFolders.Item.Messages.Delta.DeltaGetResponse>
                .CreatePageIterator(
                    graphClient,
                    messagesDeltaResponse,
                    (msg) =>
                    {
                        result.Add(msg);
                        return true;
                    },
                    (req) =>
                    {
                        req.Headers.Add("testfer", "outlook.body-content-type=\"text\"");
                        req.Headers.Add("testfer", "odata.maxpagesize=10");
                        return req;
                    });

            await pageIterator.IterateAsync();

            if (pageIterator.State == PagingState.Delta)
                return (result.ToArray(), pageIterator.Deltalink);
            else
                throw new Exception("Not expected delta pagination state."); // should not happen, see docs.
        }

        public async Task<IDictionary<string, Microsoft.Graph.Models.Message>> GetEventMessageResponsesByIds(string userId, string[] ids, string[]? select = null, string[]? expand = null)
        {
            var batchRequestContent = new BatchRequestContentCollection(graphClient);

            var listOfRequests = new Dictionary<string, string>();
            foreach (var messageId in ids)
            {
                var req = graphClient.Users[userId].Messages[messageId].ToGetRequestInformation(req =>
                {
                    req.QueryParameters.Select = select;
                    req.QueryParameters.Expand = expand;
                    // req.Headers.Add("testfer", "outlook.body-content-type=\"text\"");
                });
                listOfRequests.Add(messageId, await batchRequestContent.AddBatchRequestStepAsync(req));
            }
            var response = await graphClient.Batch.PostAsync(batchRequestContent);

            var messages = new Dictionary<string, Microsoft.Graph.Models.Message>();
            foreach (var (messageId, reqId) in listOfRequests)
                messages.Add(messageId, await response.GetResponseByIdAsync<Microsoft.Graph.Models.EventMessageResponse>(reqId));

            return messages;
        }

        public async Task<IList<Microsoft.Graph.Models.Message>> GetMessagesByIdsAsync(string userId, string[] messageIds, string[]? select = null)
        {
            if (messageIds == null || messageIds.Length == 0)
                return [];

            var defaultSelect = new[]
            {
                "id",
                "subject",
                "sentDateTime",
                "receivedDateTime",
                "lastModifiedDateTime",
                "from",
                "sender"
            };

            select ??= defaultSelect;


            var tasks = messageIds.Select(async messageId =>
            {
                try
                {
                    return await graphClient.Users[userId].Messages[messageId]
                        .GetAsync(req =>
                        {
                            req.QueryParameters.Select = select;
                            req.Headers.Add("testfer", "outlook.body-content-type=\"text\"");
                        });
                }
                catch (Exception ex)
                {
                    return null;
                }
            }).ToArray();

            var messages = await Task.WhenAll(tasks);

            return messages.Where(m => m != null).ToList()!;
        }



        public async Task<Microsoft.Graph.Models.Event> GetMeetingEventById(string userId, string calendarId, string id)
        {
            return await graphClient.Users[userId].Calendars[calendarId].Events[id].GetAsync();
        }

        public async Task<Microsoft.Graph.Models.OnlineMeeting> CreateOnlineMeeting(string userId, Microsoft.Graph.Models.OnlineMeeting eventToUpdate)
        {
            return await graphClient.Users[userId].OnlineMeetings.PostAsync(eventToUpdate);
        }

        public async Task<Microsoft.Graph.Models.OnlineMeeting> UpdateOnlineMeeting(string userId, string id, Microsoft.Graph.Models.OnlineMeeting eventToUpdate)
        {
            return await graphClient.Users[userId].OnlineMeetings[id].PatchAsync(eventToUpdate);
        }

        public async Task<Microsoft.Graph.Models.OnlineMeeting> GetOnlineMeetingById(string userId, string id)
        {
            return await graphClient.Users[userId].OnlineMeetings[id].GetAsync();
        }

        public async Task SendMail(string userId, SendMailPostRequestBody message) => await graphClient.Users[userId].SendMail.PostAsync(message);

        public async Task<Calendar> EnsureCalendarByName(string userId, string calendarName, Calendar? calendarConfig = null)
        {
            var calendars = await graphClient.Users[userId].Calendars.GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Top = 1;
                requestConfiguration.QueryParameters.Filter = $"name eq '{calendarName}'";
            });
            var calendar = calendars?.Value?.FirstOrDefault();
            if (calendar != null) return calendar;

            var newCalendar = new Calendar { Name = calendarName };
            if (calendarConfig != null)
            {
                newCalendar.Color = calendarConfig.Color;
                newCalendar.CanEdit = calendarConfig.CanEdit;
                newCalendar.CanShare = calendarConfig.CanShare;
                newCalendar.CanViewPrivateItems = calendarConfig.CanViewPrivateItems;
                newCalendar.ChangeKey = calendarConfig.ChangeKey;
                newCalendar.Owner = calendarConfig.Owner;
                newCalendar.IsDefaultCalendar = calendarConfig.IsDefaultCalendar;
                newCalendar.IsRemovable = calendarConfig.IsRemovable;
                newCalendar.IsTallyingResponses = calendarConfig.IsTallyingResponses;
            }

            return await graphClient.Users[userId].Calendars.PostAsync(newCalendar);
        }

        public async Task<DriveItem> UploadSmallFileTo(string siteId, string listId, string driveRelativePath, byte[] content)
        {
            using var stream = new MemoryStream(content);
            var drive = await graphClient.Sites[siteId].Lists[listId].Drive.GetAsync();
            return await graphClient.Drives[drive.Id].Root.ItemWithPath(driveRelativePath).Content.PutAsync(stream);
        }

        public class DriveItemsFieldsUpdateRequest
        {
            public string DriveId { get; set; }
            public string DriveItemId { get; set; }
            public string? ContentTypeId { get; set; }
            public IDictionary<string, object>? FieldValues { get; set; }
        }

        public async Task UpdateListItemFieldsInBatch(IEnumerable<DriveItemsFieldsUpdateRequest> toUpdate)
        {
            var listOfRequests = new Dictionary<string, string>();
            var getBatchRequestContent = new BatchRequestContentCollection(graphClient);
            foreach (var itemToUpdate in toUpdate)
                listOfRequests.Add(itemToUpdate.DriveItemId, await getBatchRequestContent.AddBatchRequestStepAsync(graphClient.Drives[itemToUpdate.DriveId].Items[itemToUpdate.DriveItemId].ToGetRequestInformation((rc) => rc.QueryParameters.Select = ["SharepointIds"])));

            var responses = await graphClient.Batch.PostAsync(getBatchRequestContent);

            var batchPatchRequestContent = new BatchRequestContentCollection(graphClient);
            foreach (var (driveItemId, requestId) in listOfRequests)
            {
                var driveItem = await responses.GetResponseByIdAsync<DriveItem>(requestId);
                var itemToUpdate = toUpdate.First(i => i.DriveItemId == driveItemId);

                var req = graphClient.Sites[driveItem.SharepointIds!.SiteId].Lists[driveItem.SharepointIds!.ListId].Items[driveItem.SharepointIds!.ListItemUniqueId].ToPatchRequestInformation(new ListItem()
                {
                    ContentType = (itemToUpdate.ContentTypeId is not null) ? new ContentTypeInfo() { Id = itemToUpdate.ContentTypeId } : null,
                    Fields = new FieldValueSet
                    {
                        AdditionalData = (itemToUpdate.FieldValues is not null) ? itemToUpdate.FieldValues : null
                    }
                });
                await batchPatchRequestContent.AddBatchRequestStepAsync(req);
            }
            await graphClient.Batch.PostAsync(batchPatchRequestContent);
        }

        public async Task DeleteChildrenItems(string siteId, string listId, string driveRelativePath)
        {
            var drive = await graphClient.Sites[siteId].Lists[listId].Drive.GetAsync();
            var children = await graphClient.Drives[drive.Id].Root.ItemWithPath(driveRelativePath).Children.GetAsync();
            Task.WaitAll(children?.Value?.Select(child => graphClient.Drives[drive.Id].Items[child.Id].DeleteAsync()).ToArray());
        }

        public async Task<(DriveItem?, Drive)> GetDriveItemByPath(string siteId, string listId, string driveRelativePath)
        {
            var drive = await graphClient.Sites[siteId].Lists[listId].Drive.GetAsync();
            var driveItem = await graphClient.Drives[drive.Id].Root.ItemWithPath(driveRelativePath).GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Select = ["*"];
            });
            return (driveItem, drive);
        }

        public async Task<Dictionary<string, DriveItem>?> GetDriveItemsByPaths(string siteId, string listId, string[] driveRelativePath, string[] selectedproperties, string[] expandproperties = null)
        {
            var result = new Dictionary<string, DriveItem>();
            var drive = await graphClient.Sites[siteId].Lists[listId].Drive.GetAsync();

            var batchRequestContent = new BatchRequestContentCollection(graphClient);
            var listOfRequestsToPaths = new Dictionary<string, string>();
            foreach (var driveItemUrl in driveRelativePath)
            {
                var req = graphClient.Drives[drive.Id].Root.ItemWithPath(driveItemUrl).ToGetRequestInformation(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = selectedproperties;
                    requestConfiguration.QueryParameters.Expand = expandproperties;
                });
                listOfRequestsToPaths.Add(driveItemUrl, await batchRequestContent.AddBatchRequestStepAsync(req));
            }

            var responses = await graphClient.Batch.PostAsync(batchRequestContent);

            foreach (var response in listOfRequestsToPaths)
                result.Add(response.Key, await responses.GetResponseByIdAsync<DriveItem>(response.Value));

            return result;
        }

        public async Task<Dictionary<string, DriveItem>?> GetDriveItemsByPaths(string driveId, string[] driveRelativePath, string[] selectedproperties, string[] expandproperties = null)
        {
            var result = new Dictionary<string, DriveItem>();

            var batchRequestContent = new BatchRequestContentCollection(graphClient);
            var listOfRequestsToPaths = new Dictionary<string, string>();
            foreach (var driveItemUrl in driveRelativePath)
            {
                var req = graphClient.Drives[driveId].Root.ItemWithPath(driveItemUrl).ToGetRequestInformation(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = selectedproperties;
                    requestConfiguration.QueryParameters.Expand = expandproperties;
                });
                listOfRequestsToPaths.Add(driveItemUrl, await batchRequestContent.AddBatchRequestStepAsync(req));
            }

            var responses = await graphClient.Batch.PostAsync(batchRequestContent);

            foreach (var response in listOfRequestsToPaths)
                result.Add(response.Key, await responses.GetResponseByIdAsync<DriveItem>(response.Value));

            return result;
        }

        public async Task<Dictionary<string, DriveItem>?> GetDriveItemsByIds(string driveId, string[] driveItemsIds, string[] selectedproperties, string[] expandproperties = null)
        {
            var result = new Dictionary<string, DriveItem>();

            var batchRequestContent = new BatchRequestContentCollection(graphClient);
            var listOfRequestsToDriveItemIds = new Dictionary<string, string>();
            foreach (var driveItemId in driveItemsIds)
            {
                var req = graphClient.Drives[driveId].Items[driveItemId].ToGetRequestInformation(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = selectedproperties;
                    requestConfiguration.QueryParameters.Expand = expandproperties;
                });
                listOfRequestsToDriveItemIds.Add(driveItemId, await batchRequestContent.AddBatchRequestStepAsync(req));
            }

            var responses = await graphClient.Batch.PostAsync(batchRequestContent);

            foreach (var driveItemIdToRequest in listOfRequestsToDriveItemIds)
                result.Add(driveItemIdToRequest.Key, await responses.GetResponseByIdAsync<DriveItem>(driveItemIdToRequest.Value));

            return result;
        }

        public async Task<Stream?> GetFileAsPdfByItemId(string siteId, string listId, string itemId)
        {
            var getReq = graphClient.Sites[siteId].Lists[listId].Items[itemId].DriveItem.Content.ToGetRequestInformation();
            getReq.UrlTemplate += "{?format}";
            getReq.QueryParameters.Add("format", "pdf");
            return await graphClient.RequestAdapter.SendPrimitiveAsync<Stream>(getReq);
        }

        public async Task<Stream?> GetFileAsPdfByPath(string siteId, string listId, string driveRelativePath)
        {
            var drive = await graphClient.Sites[siteId].Lists[listId].Drive.GetAsync();
            var getReq = graphClient.Drives[drive!.Id].Root.ItemWithPath(driveRelativePath).Content.ToGetRequestInformation();
            getReq.UrlTemplate += "{?format}";
            getReq.QueryParameters.Add("format", "pdf");
            return await graphClient.RequestAdapter.SendPrimitiveAsync<Stream>(getReq);
        }

        public async Task<Stream?> GetFileByItemId(string siteId, string listId, string itemId)
        {
            var getReq = graphClient.Sites[siteId].Lists[listId].Items[itemId].DriveItem.Content.ToGetRequestInformation();
            return await graphClient.RequestAdapter.SendPrimitiveAsync<Stream>(getReq);
        }

        public async Task<Stream?> GetFileByPath(string siteId, string listId, string driveRelativePath)
        {
            var drive = await graphClient.Sites[siteId].Lists[listId].Drive.GetAsync();
            var getReq = graphClient.Drives[drive!.Id].Root.ItemWithPath(driveRelativePath).Content.ToGetRequestInformation();
            return await graphClient.RequestAdapter.SendPrimitiveAsync<Stream>(getReq);
        }

        public async Task<ExtractSensitivityLabelsResult?> GetFileSensitivityLabelsByPath(string siteId, string listId, string driveRelativePath)
        {
            var (driveItem, drive) = await GetDriveItemByPath(siteId, listId, driveRelativePath);

            return await graphClient.Drives[drive!.Id].Items[driveItem!.Id].ExtractSensitivityLabels.PostAsync();
        }

        public async Task<HashSet<Guid>> GetFilesIdsWithExternalLabels(string siteId, string listId, Guid tenantId, Guid[] fileUniqueIds)
        {
            var filesWithExternalLabels = new HashSet<Guid>();

            var drive = await graphClient.Sites[siteId].Lists[listId].Drive.GetAsync();

            foreach (var uniqueId in fileUniqueIds)
            {
                try
                {
                    var sensivity = await graphClient.Drives[drive!.Id].Items[uniqueId.ToString()].ExtractSensitivityLabels.PostAsync();
                    var isExternal = sensivity?.Labels?.Any(x => Guid.Parse(x.TenantId ?? Guid.ECNTy.ToString()) != tenantId) ?? false;
                    if (isExternal)
                        filesWithExternalLabels.Add(uniqueId);
                }
                catch (Exception ex)
                {
                    if (ex.Message.ToLowerInvariant().Contains("unsupported user is attached"))
                        filesWithExternalLabels.Add(uniqueId);
                }
            }

            return filesWithExternalLabels;
        }

        public async Task<HashSet<Guid>> GetFilesIdsWithUnsupportedLabels(string siteId, string listId, Guid tenantId, Guid[] fileUniqueIds)
        {
            var filesWithExternalLabels = new HashSet<Guid>();

            var drive = await graphClient.Sites[siteId].Lists[listId].Drive.GetAsync();

            foreach (var uniqueId in fileUniqueIds)
            {
                try
                {
                    var sensivity = await graphClient.Drives[drive!.Id].Items[uniqueId.ToString()].ExtractSensitivityLabels.PostAsync();
                }
                catch (Exception ex)
                {
                    if (ex.Message.ToLowerInvariant().Contains("unsupported user is attached") || ex.Message.ToLowerInvariant().Contains("file's encryption"))
                        filesWithExternalLabels.Add(uniqueId);
                }
            }

            return filesWithExternalLabels;
        }

        public async Task<IDictionary<Guid, Guid?>> GetFilesSensitivityLabelIds(string siteId, string listId, Guid tenantId, Guid[] fileUniqueIds)
        {
            var fileToLabelMapping = new Dictionary<Guid, Guid?>();

            var drive = await graphClient.Sites[siteId].Lists[listId].Drive.GetAsync();

            foreach (var uniqueId in fileUniqueIds)
            {
                try
                {
                    var sensivity = await graphClient.Drives[drive!.Id].Items[uniqueId.ToString()].ExtractSensitivityLabels.PostAsync();
                    var currentLabel = sensivity?.Labels?.FirstOrDefault(l => Guid.Parse(l.TenantId ?? Guid.ECNTy.ToString()).Equals(tenantId));
                    fileToLabelMapping.Add(uniqueId, currentLabel?.SensitivityLabelId == null ? null : new Guid(currentLabel.SensitivityLabelId));
                }
                catch (Exception)
                {
                    // TODO: log error ; file could be locked
                }
            }

            // // !!!!! Graph BUG This does not work.
            // var batchRequestContent = new BatchRequestContentCollection(graphClient);
            // var listOfRequests = new Dictionary<Guid, string>();
            // foreach (var uniqueId in fileUniqueIds)
            // {
            //     var req = graphClient.Drives[drive!.Id].Items[uniqueId.ToString()].ExtractSensitivityLabels.ToPostRequestInformation();
            //     listOfRequests.Add(uniqueId, await batchRequestContent.AddBatchRequestStepAsync(req));
            // }
            // var responses = await graphClient.Batch.PostAsync(batchRequestContent);
            // foreach (var responseId in listOfRequests)
            // {
            //     try
            //     {
            //         var sensivity = await responses.GetResponseByIdAsync<ExtractSensitivityLabelsResult?>(responseId.Value);
            //         var currentLabel = sensivity?.Labels?.FirstOrDefault(l => Guid.Parse(l.TenantId ?? Guid.ECNTy.ToString()).Equals(tenantId));
            //         fileToLabelMapping.Add(responseId.Key, currentLabel?.SensitivityLabelId == null ? null : new Guid(currentLabel.SensitivityLabelId));
            //     }
            //     catch (Exception)
            //     {
            //         // TODO: log error ; file could be locked
            //     }
            // }

            return fileToLabelMapping;
        }

        public async Task SetFilesSensitivityLabels(string siteId, string listId, Dictionary<Guid, Guid> fileUniqueIdsToLabelIds, string justification = "")
        {
            var drive = await graphClient.Sites[siteId].Lists[listId].Drive.GetAsync();

            foreach (var fileToLabel in fileUniqueIdsToLabelIds)
            {
                await graphClient.Drives[drive!.Id].Items[fileToLabel.Key.ToString()].AssignSensitivityLabel.PostAsync(new AssignSensitivityLabelPostRequestBody()
                {
                    SensitivityLabelId = fileToLabel.Value.ToString(),
                    JustificationText = justification,
                    AssignmentMethod = SensitivityLabelAssignmentMethod.Standard
                });
            }

            // // !!!!! Graph BUG This does not work.
            // var batchRequestContent = new BatchRequestContentCollection(graphClient);
            // var listOfRequests = new Dictionary<Guid, string>();
            // foreach (var fileToLabel in fileUniqueIdsToLabelIds)
            // {
            //     var req = graphClient.Drives[drive!.Id].Items[fileToLabel.Key.ToString()].AssignSensitivityLabel.ToPostRequestInformation(new AssignSensitivityLabelPostRequestBody()
            //     {
            //         SensitivityLabelId = fileToLabel.Value.ToString(),
            //         JustificationText = justification,
            //         AssignmentMethod = SensitivityLabelAssignmentMethod.Standard
            //     });
            //     listOfRequests.Add(fileToLabel.Key, await batchRequestContent.AddBatchRequestStepAsync(req));
            // }
            // await graphClient.Batch.PostAsync(batchRequestContent);
        }

        public async Task SetFileSensitivityLabelsByPath(string siteId, string listId, string driveRelativePath, string labelId, string justification = "")
        {
            var (driveItem, drive) = await GetDriveItemByPath(siteId, listId, driveRelativePath);

            await graphClient.Drives[drive!.Id].Items[driveItem!.Id].AssignSensitivityLabel.PostAsync(new AssignSensitivityLabelPostRequestBody()
            {
                SensitivityLabelId = labelId,
                JustificationText = justification,
                AssignmentMethod = SensitivityLabelAssignmentMethod.Standard
            });
        }

        public async Task<IEnumerable<SensitivityLabel>> GetTenatSensitivityLabels()
        {
            var query = await graphClient.Security.DataSecurityAndGovernance.SensitivityLabels.GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Select = ["id", "name", "applicableTo", "isEnabled"];
            });
            var result = query != null && query.Value != null && query.Value.Count != 0 ? query.Value.Where(x => x.AdditionalData.TryGetValue("applicableTo", out var applicableTo) && applicableTo.ToString()!.Contains("site") && x.AdditionalData.TryGetValue("isEnabled", out var isEnabled) && (bool)isEnabled).ToList() : [];
            return result;
        }

        public async Task<Dictionary<Guid, Term>> GetAllTermsFrom(string siteId, string groupId, string[] termSets)
        {
            var termIdToRequestMap = new Dictionary<Guid, Term>();

            var batchRequestContent = new BatchRequestContentCollection(graphClient);
            var listOfRequests = new List<string>();
            foreach (var set in termSets)
            {
                var req = graphClient.Sites[siteId].TermStore.Groups[groupId].Sets[set].Terms.ToGetRequestInformation(requestConfiguration =>
                {
                    // https://graph.microsoft.com/v1.0/sites/contoso-dev-admin.sharepoint.com/termStore/groups/102a5d61-9d78-42e8-ab11-681331d4d37f/sets/172d84e2-ba88-40d6-a872-8fe00d066caa/terms?$select=id,labels
                    requestConfiguration.QueryParameters.Select = new string[] { "id", "labels" };
                    requestConfiguration.QueryParameters.Top = 999;
                });
                listOfRequests.Add(await batchRequestContent.AddBatchRequestStepAsync(req));
            }

            var responses = await graphClient.Batch.PostAsync(batchRequestContent);

            foreach (var responseId in listOfRequests)
            {
                var termsResponse = await responses.GetResponseByIdAsync<TermCollectionResponse>(responseId);
                var pageIterator = PageIterator<Term, TermCollectionResponse>
                    .CreatePageIterator(graphClient, termsResponse, (term) =>
                    {
                        termIdToRequestMap.TryAdd(new Guid(term.Id!), term);
                        return true;
                    });
                await pageIterator.IterateAsync();
            }

            return termIdToRequestMap;
        }

        public async Task<Dictionary<Guid, Term>> GetAllTermsFrom(string siteId, string groupId)
        {
            var termIdToRequestMap = new Dictionary<Guid, Term>();
            var termSetsResponse = await graphClient.Sites[siteId].TermStore.Groups[groupId].Sets.GetAsync(requestConfiguration =>
            {
                // https://graph.microsoft.com/v1.0/sites/contoso-dev-admin.sharepoint.com/termStore/groups/102a5d61-9d78-42e8-ab11-681331d4d37f/sets/172d84e2-ba88-40d6-a872-8fe00d066caa/terms?$select=id,labels
                requestConfiguration.QueryParameters.Select = new string[] { "id" };
                requestConfiguration.QueryParameters.Top = 9999;
            });

            var termSets = new List<string>();
            var pageIterator = PageIterator<Set, SetCollectionResponse>
                .CreatePageIterator(graphClient, termSetsResponse, (termSet) =>
                {
                    termSets.Add(termSet.Id!);
                    return true;
                });
            await pageIterator.IterateAsync();
            return await GetAllTermsFrom(siteId, groupId, termSets.ToArray());
        }

        public async Task<Term?> UpdateTerm(string siteId, string groupId, string setId, Term term)
        {
            return await graphClient.Sites[siteId].TermStore.Groups[groupId].Sets[setId].Terms[term.Id].PatchAsync(term);
        }
        public async Task<Term?> AddTerm(string siteId, string groupId, string setId, Term term)
        {
            return await graphClient.Sites[siteId].TermStore.Groups[groupId].Sets[setId].Children.PostAsync(term);
        }

        public async Task<Site?> UpdateSite(string host, string siteId, Site site)
        {
            var result = await graphClient.Sites[$"{host}:/sites/{siteId}"].PatchAsync(site);
            return result;
        }

        public async Task<Microsoft.Graph.Models.Extension?> GetUserOpenExtension(string userId, string extensionId)
        {
            var user = await graphClient.Users[userId].GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Expand = [$"Extensions($filter=id eq '{extensionId}')"]; // TODO: hardcoded top 999
            });
            return user?.Extensions?.FirstOrDefault();
        }

        public async Task SetUserOpenExtension(string userId, string extensionId, Dictionary<string, object> data, bool? exists = false)
        {
            bool shouldCreate = false;

            if (!exists.HasValue)
            {
                var existing = await GetUserOpenExtension(userId, extensionId);
                shouldCreate = existing == null;
            }

            if (shouldCreate)
                await graphClient.Users[userId].Extensions.PostAsync(new OpenTypeExtension
                {
                    OdataType = "microsoft.graph.openTypeExtension",
                    ExtensionName = extensionId,
                    AdditionalData = data
                });
            else
                await graphClient.Users[userId].Extensions[extensionId].PatchAsync(new OpenTypeExtension
                {
                    AdditionalData = data
                });
        }

        public void Dispose()
        {
            ((IDisposable)graphClient)?.Dispose();
        }
    }

    public class DocSetDocumentChangesResult
    {
        public List<DocumentSetVersionItem> Added { get; set; }
        public List<DocumentSetVersionItem> Removed { get; set; }
        public List<DocumentSetVersionItem> Updated { get; set; }
        public List<DocumentSetVersionItem> NotChanged { get; set; }
    }

    public class PnPCoreToGraphSDKTokenprovider : IAccessTokenprovider
    {
        private readonly IPnPContext _ctx;
        public PnPCoreToGraphSDKTokenprovider(IPnPContext ctx) => _ctx = ctx;
        public Task<string> GetAuthorizationTokenAsync(Uri uri, Dictionary<string, object> additionalAuthenticationContext = default, CancellationToken cancellationToken = default) => _ctx.Authenticationprovider.GetAccessTokenAsync(new Uri("https://graph.microsoft.com"), new string[] { ".default" });
        public AllowedHostsValidator AllowedHostsValidator { get; }
    }
}