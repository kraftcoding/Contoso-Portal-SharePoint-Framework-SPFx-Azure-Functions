
namespace Contoso.Portal.Model.Events
{
    public class EventVotation
    {
        public string UniqueSharedID { get; set; }

        public string VoteRetestsentationId { get; set; } //Taxonomia de la comunidad retestsentada

        public string[] AssignedUsers { get; set; } //Usuarios que pueden ejercer el voto en nombre de la comunidad retestsentada

        public Dictionary<string, string> Votes { get; set; } //Listado de votos

    }

    public class Vote
    {
        public string Id; //UniqueSharedId del acuerdo
        public string VoteId; //Taxonomia del voto (aprodbado, rechazado, abstencion)
    }

    public class Summary
    {
        public int Pending;
        public int Approdve;
        public int Abstention;
        public int Reject;
    }


}
