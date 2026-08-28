using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Domains.Notifications;
using Contoso.Portal.Domains.Profile;
using Contoso.Portal.Model.Notifications;

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;
using Contoso.Portal.Data.Extensions;
using System.Linq;
using static Contoso.Portal.Data.DAL.DALConstants;
using Microsoft.Extensions.DependencyInjection;
using Contoso.Portal.Model.Bodies;

namespace WebApi.Tests.Notifications;

public class NotificationsBusinessServiceTests : BasePnPAppTest
{
    // CONFIGURAR .................
    private string BODY_ID = "demomvp-cs-ae";
    private string SHARED_EVENT_ID = "ab6a84a4da33438290de03fe11f957cc";
    private string USER_UPN = "gcastrma_emeal.nttdata.com@test.onmicrosoft.com";
    private string URL = "https://test.sharepoint.com/sites/Contoso";

    private readonly NotificationsBusinessService _notiBusinessService;

    public NotificationsBusinessServiceTests(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output)
    {
        _notiBusinessService = ServiceProvider().GetRequiredService<NotificationsBusinessService>();
    }

    [Fact]
    public async Task SendRememberEventNotificationTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);

            await _notiBusinessService.SendRememberEventNotification(ctx, BODY_ID, SHARED_EVENT_ID, RememberNotification.OneHour);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }



}
