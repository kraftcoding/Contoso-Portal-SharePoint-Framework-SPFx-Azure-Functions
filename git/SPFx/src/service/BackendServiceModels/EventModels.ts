
export interface IEvent {
    Id: string;
    Title: string;
    Description: string;
    StartDate: Date;
    EndDate: Date;
    Location: string;
    LocationDetails: string;
    MeetingToolUrl: string;
    StorageServerRelativeUrl: string;

    // taxonomy fields
    AttendanceTypeId: string;
    StatusId: string;
    BodyTypeId: string;
    EventToolId: string;
    BodyNameId: string;

    LastPublished?: Date;
    BodyId: string;
    ShowMeetingToolUrl?:boolean;
}

export enum EventStatus {
    Canceled = '47a651b0-921b-40ad-bf16-90d9687f4004',
    Celebrated = 'bd52a4c8-d276-4bb0-a1f8-12fb7adf8999',
    InCelebration = '7c80f44b-6586-4d98-82a8-a43acd09b99d',
    InConstruction = '1214ed8e-cfbd-4d7a-b5b7-867654823f74',
    testBooking = '96909cc0-ab33-48c0-9bfd-714add3bdb4d',
    Published = '8db5d28f-c8e4-463d-99bd-6e2bd4196216',
    Archived = '2522ab31-a67d-4ebd-abc4-733610aad9a9'
}

export interface INewOrUpdatedEvent {
    Title: string;
    Description: string;
    StartDate: Date;
    EndDate: Date;
    Location: string;
    LocationDetails: string;
    MeetingToolUrl: string;

    // taxonomy fields
    AttendanceTypeId: string;
    EventToolId: string;

    Attendees: string[]; // members
    Guests: string[]; // guests
}

export interface IEventAgendaItem {
    Id: string;
    Title: string;
    Description: string | null;
    Duration: number;
    Order: number;
    AgendaItemType: string | null;
    RelatedDocumentsIds: string[];
    ParentId: string |  null;
    OrderType: string | null;
    AgreementRelatedCertificateId?: string;
}

export interface IUserAttendance {
    AttendanceId: string;
    UserPrincipalName: string;
    RequestStatusId: string;
    RequestUserMessage: string;
    RequestStatusUpdate: Date;
    HasAttended: boolean;
    AttendanceTypeId: string;
    DelegationUserPrincipalName: string;
    DelegationUserVotePrincipalName: string;
    Vote: boolean;
    IsGuest: boolean;
    VoteDelegation: boolean;
}

export interface IEventAgreement {
    Id: string;
    Title: string;
    Description: string | null;
    StatusId: string;
    RelatedDocumentsIds: string[];
    RelatedCertificateId: string;
    PendingCertificateTask: boolean;
    Order: string;
}

export enum AgendaItemTypes {
    Decisive = 'c83ab3b1-00a1-4023-9a2f-1697b23c13d5',
    Coordination = '264eb3ec-23d4-4cc9-8ace-8a643483db0c',
    Informative = 'ca28d183-ce3e-4339-a516-f3aec1940622'
}

export enum AgreementStatus {
    Approdved = '04012a59-69bf-4f39-99b3-799d6d9e4f26',
    Declined = '76426a2a-907d-4717-9857-5c704e7b1c1b',
    OverTheTable = '22abc0e5-25f1-4c07-80c5-b715dba4a1a6',
    Inprodgress = "5996ee35-33ef-4fce-820d-06db4c5cfcad",
    UnanimouslyApprodved = "ddf339aa-3f6a-4a34-a24e-55b2b6efa7f5",
    Finished = 'Finished'
}

export enum OrderTypes {
    Alphabetical = "c2c12ed4-1001-466f-977c-07249f45c3a0",
    Numeric = "2cc86787-b808-4ea4-b5f7-d92f9cc9b741",
    Roman = "f980cb48-6220-45d0-a9f8-3266bc31c2ba"
}

export enum VotationStatus {
    Approdve = "bd558d1a-e24b-4882-865d-853c7a14289e",
    Abstention = "0ae772d2-0f28-4eaf-95a9-8c857d78d51c",
    Reject = "b0c9703c-16a2-4c65-966d-7f39b67c662e"
}

export interface IEventStorage {
    InConstruction: IEventContainer;
    Published: IEventContainer;
}

export interface IEventContainer {
    SiteId: string;
    DriveId: string;
    DriveItemId: string;
    ListItemId: string;
    ListItemUniqueId: string;
    ServerRelativeUrl: string;
    DriveRelativePath: string;
}

export interface IEventSendOperationInformation {
    ToCaptureState: IEventCaptureState;
    FromCaptureState?: IEventCaptureState;
    StatusHasChanged: boolean;
    NewEventStatusId: string;
    CurrentEventStatusId: string;
    AttendeesHasChanged: boolean;
    Attendees: IAttendeesChangesResult;
    DocumentsHasChanged: boolean;
    Documents: IDocSetChangesResult;
    AttendeesMustReconfirm: boolean;
    DetailsHasChanged: boolean;
    SentBy: string;
    BodyId: string;
    SharedEventId: string;
}

export interface IAttendeesChangesResult {
    Added: string[];
    Removed: string[];
    NotChanged: string[];
}

export interface IEventCaptureState {
    VersionLabel: string;
    VersionId: number;
    CaptureId: number;
}

export interface IDocSetChangesResult {
    Added: IDocSetDocumentInformation[];
    Updated: IDocSetDocumentInformation[];
    NotChanged: IDocSetDocumentInformation[];
    Removed: { ItemUniqueId: string, VersionLabel: string }[];
    NotAvalible: { ItemUniqueId: string, VersionLabel: string }[];
}

export interface IDocSetDocumentInformation {
    ItemId: number;
    ItemUniqueId: string;
    FileServerRelativePath: string;
    DocSetRelativeUrl: string;
    FileName: string;
    CurrentVersionLabel: string;
    DocSetVersionLabel: string;
    ContentTypeId: string;
    FileSize: string;
}

export enum RelatedDocumentCType {
    //tipos de contenido de los Documents relacionados
    MeetingMinute = '0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC01020201',
    MeetingAsistanceReport = '0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC01020202',
    AgreementCertification = '0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC01020203',
    MeetingDocument = '0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC0102'
}

export interface IEventVotationDetails {
    Id: string;
    VoteRetestsentationId: string;
    AssignedUsers: string[];
    Votes: Record<string,string>;
}

export interface IVote {
    Id: string;
    VoteId: string;
}

export interface ITenantSensitivityLabel{
    Id: string;
    Name: string;
}

export interface IDocumentSensitivityLabel
{
    BodyId :string;
    SharedEventId :string;
    FilePath :string;
    SensitivityLabelId :string;
    IsExternalLabel :boolean;
    IsSignedDocument :boolean;
    IsOtherError :boolean;
    IsprocessingLabel :boolean;
    ExpectedLabel :string;
}
export interface IEventFromBbdd {
    StartDate: Date;
    EndDate: Date;
    Title: string;
    Department: string;
    EstadoMeeting: string;
    NombreDepartment: string;
    MeetingType: string;
    TipoDepartment: string;
    RowKey: string;
}