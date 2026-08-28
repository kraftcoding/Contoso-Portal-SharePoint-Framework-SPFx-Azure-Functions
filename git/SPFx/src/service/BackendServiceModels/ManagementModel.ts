export interface IMemberdepartmentManagement {
  UserPrincipalName: string;
  UserRoles: string[];
  RoleId?:string;
  Ccaa?:string;
}
  export interface IUserOrganList{
    OrganId:string;
    DepartmentName:string;
    Role:string;
    RoleId:string;
}

export interface IUsersOrganRelation{
  UserUpn: string;
  BodyId: string;
  Role?:string;
}

export interface IdepartmentManagement {
  BodyId: string;
  Department: string;
  TipoDepartment: string;
  Secretaria: string;
  Division: string;
  Description:string;
  Intersectorial:boolean;
  dir: string;
  SIA: string;
  Materia: string;
  Abreviatura: string;
  Period: string;
  Activo: boolean;
  FechaConstitucion?: string;
  FechaExtincion?: string;
  Observaciones: string;
  IdEmailBodies: number;
  IdEmailBodiesOnline?: number;
  IdEmailBodiesOnlineInPerson?: number;
  IdEmailBodiesInPerson?: number;
  IdEmailBodiesWrittenprodcedure?: number;
  IdEmailBodiesDocumentationReferral?: number;
  IdActBodies: number;
  IdCertificateBodies: number;
  IdAgendaBodies: number;
  IdAttendanceBodies: number;
  DiasAprodbacionMinutes: number;
  BusinessArea: string;
  NombreDepartment:string;
  IsEXTERNAL?:boolean;
  TipoMembresia: string;
  IdentificadorConferencia: string;
  Users: IMemberdepartmentManagement[];
}

export interface IOrganEXTERNAL {
  BodyId: string;
  ID?: string;
  Department: string;
  TipoDepartmentEXTERNAL: string;
  SecretariaText: string;
  Division?: string;
  DocumentSetDescription:string;
  Intersectorial:boolean;
  dir: string;
  SIA: string;
  Abreviatura: string;
  Activo: boolean;
  FechaConstitucion?: string;
  FechaExtincion?: string;
  FechaInscripcion?:string;
  Observaciones: string;
  IdCertificateBodies: number;
  BusinessArea: string;
  NombreDepartment:string;
  StatusDepartment:string;
  IsEXTERNAL?:boolean;
  CodigoDepartment: string;
  FileRef?:string;
  Inscrito:boolean;
  FechaCreacion?: string;
  DepartmentAdscripcion?: string;
}

export interface IUserManagement {
  DisplayName: string;
  PrincipalMail: string;
  AssignedLicenses: string[];
  FirstName: string;
  LastName: string;
  Email: string;
  CellPhone: string;
  JobTitle: string;
  Office: string;
  UserPrincipalName: string;
  CompanyName?: string;
  Department?: string;
  EmployeeType?: string;
  BusinessPhone?: string;
  Organ?:IUserOrganList[];
}

export interface IRelationBusinessArea{
  Id?:string;
  BusinessArea: string;
  Division: string;
  StartDate: string;
  EndDate: string;
}

export interface ILicenseInformation {
  SkuId: string;
  Total: string;
  Assigned: string;
  Available: string;
}
export interface IRelationBusinessArea{
  Id?:string;
  BusinessArea: string;
  Division: string;
  StartDate: string;
  EndDate: string;
}

export interface BodyUserInformation{
  UserBodyRole?:UserBodyRoleInformation;
  UserInfo:profileInformation;
  UserInvitation?: NewUserInvitation;
}

export interface NewUserInvitation {
    InvitationMessage?:string;
    InvitationCCRecipient?:string;
}

export interface UserBodyRoleInformation {
  UserUpn?: string;
  BodyId?: string;
  Role?: string;
}

export interface profileInformation {
  FirstName?: string;
  LastName?: string;
  DisplayName?: string;
  Email?: string;
  CellPhone?: string;
  JobTitle?: string;
  Office?: string;
  PrincipalMail?: string;
  CompanyName?: string;
  Department?: string;
  EmployeeType?: string;
  BusinessPhone?: string;
}