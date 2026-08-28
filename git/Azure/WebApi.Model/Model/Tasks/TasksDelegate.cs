
namespace Contoso.Portal.Model.Tasks
{
    public class TasksDelegate : TasksCNT
    {
        //Delegado a
        public string DelegateTo { get; set; }

        public string DelegatedUserVote { get; set; }

        public bool? DelegateVote { get; set; }

    }
}