export interface IUserTaskModel {
    BodyId: string;
    BodyNameId: string;
    TaskId: string;
    TaskTypeId: string;
    TaskTitle: string;
    TaskStartDate: string;
    TaskEndDate: string;
    //
    AssignedTo: string[];
    BodyTypeId: string;
    Comment: string;
    EventAttendanceTypeId: string;
    Languaje: string;
    RequestedBy: string;
    SharedEventId: string;
    TaskStatusId: string;

    DelegatedUser: string;
    DelegatedUserVote: string;

    Vote: boolean;
    AgreementTitle: string;
}

export interface IUserTaskDelegateModel {
    TaskTitle: string;
    TaskStartDate: string;
    TaskEndDate: string;
    Comment: string;
    SharedEventId: string;
    TaskParentId: string;
    DelegateTo: string;
    DelegateVote: boolean;
    DelegatedUserVote: string;
}

export interface ITaskCertification {
    AgreementSharedId: string;
    SharedEventId: string;
    Comment: string;
    AgreementTitle: string;
}