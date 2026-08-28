
// using System.Runtime.Serialization;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace Contoso.Portal.Model.Events
{
    public enum EventType
    {
        Undefined,
        Online,
        InPerson,
        Hybrid,
        Writtenprodcedure
    }

    public enum EventTool
    {
        TeamsMeeting,
        Reunete,
        Zoom,
    }

    public enum EventStatus
    {
        Undefined,
        Draft,
        Published,
        Cancelled,
    }
}

