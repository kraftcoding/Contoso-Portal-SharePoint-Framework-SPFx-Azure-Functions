export interface BodyUserInfo{
    UserBodyRole?:UserBodyRole;
    UserInfo:profileInfo;
    UserInvitation?: NewUserInvitation;
}

export interface UserBodyRole {
    UserUpn?: string;
    BodyId?: string;
    Role?: string;
}

export interface profileInfo {
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

export interface NewUserInvitation {
    InvitationMessage?:string;
    InvitationCCRecipient?:string;
}

export interface NewUserErrorsForm {
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
    Organ?:string;
    Rol?:string;
    Cco?:string;
    Message?:string;
}

export interface NewUserFromExcel {
    Nombre:string;
    Apellidos:string;
    Email: string;
    Cargo?: string;
    ComunidadAutonoma?: string;
    Rol?: string;
    Division_GobiernoAutonomico?:string;
    Unidad_Consejeria?:string;
    TelefonoMovil?:string;
    TelefonoEmtestsa?:string;
    CorreoElectronicoSecundario?:string;
    TipoEmpleado?:string;
    DestinatarioInvitacionCC?:string;
}

export enum excelColumsMap {
    Nombre ='Nombre*',
    Apellidos ='Apellidos*',
    Email = 'Correo electrónico*',
    Cargo = 'Cargo',
    ComunidadAutonoma = 'Comunidad autónoma',
    Rol = 'Rol (*)',
    Division_GobiernoAutonomico = 'Division / Gobierno autonómico',
    Unidad_Consejeria = 'Unidad / Consejería',
    TelefonoMovil = 'Teléfono móvil',
    TelefonoEmtestsa = 'Teléfono de emtestsa',
    CorreoElectronicoSecundario = 'Correo electrónico secundario',
    TipoEmpleado = 'Tipo de empleado',
    DestinatarioInvitacionCC = 'Destinatario en copia de la invitación'
}
export enum componentTypeMode {
    All = 'all',
    Users = 'users',
    Organs = 'organs'
}

export interface IUserPermissionsCheck{
    UserPrincipalName?: string;
    DisplayName?: string;
    BodyId?: string;
    Role?: string;
    HasPermissions?: boolean;
    isLoading?: boolean;
    // §5: Campos de trazabilidad del prodceso de asignacion
    AssignmentStatus?: string;
    NumIntentos?: number;
    IsErrorPermanente?: boolean;
    ErrorMensaje?: string;
}

// §6: Resultado devuelto por la Web API CheckPermissions
export interface IPermissionCheckResult {
    HasPermissions: boolean;
    AssignmentStatus: string;
    NumIntentos: number;
    IsErrorPermanente: boolean;
    ErrorMensaje?: string;
    IdOperation?: string;
}

export interface IMemberOrgan {
    UserPrincipalName: string;
    UserRoles: string[];
    RoleId?:string;
    Ccaa?:string;
}

export interface IUsersOrganRelation {
    UserUpn: string;
    BodyId: string;
    Role?:string;
}

export interface IUserOrganList{
    OrganId:string;
    DepartmentName:string;
    RoleId:string;
    Role:string;
}


export interface IUsersInOrgan{
    Type: SearchType;
    Items:IUserOrgan[];
    hasError?:boolean;
}

export interface IUserOrgan{
    IndexKey?:number;
    OrganId?: string;
    UserId?: string;
    Role?: string;
    Ccaa?:string;
}

export interface IOrganResult {
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
    Users: IMemberOrgan[];
    BusinessArea: string;
    NombreDepartment:string;
    IsEXTERNAL?:boolean;
    TipoMembresia: string;
    IdentificadorConferencia: string;
  }

  export interface IUserResult {
    DisplayName: string;
    PrincipalMail: string;
    AssignedLicenses: string[];
    FirstName: string;
    LastName: string;
    Email: string;
    CellPhone: string;
    JobTitle: string;
    Office: string;
    CompanyName: string;
    Department: string;
    EmployeeType: string;
    BusinessPhone: string;
    UserPrincipalName: string;
    Organ?:IUserOrganList[];
  }


export interface IManagedItem {
    itemId: string;
    itemType: SearchType;
    organDialogContent: IOrganResult | undefined;
    userDialogContent: IUserResult | undefined;
    isNewItem?:boolean;
}

export interface IOrganAdministration {
    Id: string;
    Title: string;
    OrganType: string;
    Nomenclatura: string;
    Division?: string;
    Description: string;
    NuevaVentana: boolean;
    URL: {
        Url: string;
    };
    Users?:[]
}

export interface IObjectDifferences {
    property: string;
    oldValue: string;
    newValue: string;
}

export interface IErrorsForm {
    DisplayName?: string;
    PrincipalMail?: string;
    FirstName?: string;
    LastName?: string;
    Email?: string;
    CellPhone?: string;
    JobTitle?: string;
    Office?: string;
    CompanyName?: string;
    Department?: string;
    EmployeeType?: string;
    BusinessPhone?: string;

    Department?: string;
    TipoDepartment?: string;
    Secretaria?: string;
    Division?: string;
    Description?:string;
    Intersectorial?:boolean;
    dir?: string;
    SIA?: string;
    Materia?: string;
    Abreviatura?: string;
    Period?: string;
    FechaConstitucion?: string;
    FechaExtincion?: string;
    Observaciones?: string;
    DiasAprodbacionMinutes?: string;
    BusinessArea?:string;
}

export enum OrganTaxonomyIds{
    OrganType = 'ba3ff2ef-09c8-4345-9bad-777014ddf636',
    Organ = '6f98455a-a1c9-45e7-8c5d-aaca87690eff',
    Ministry = '4581a675-bfd7-4297-a403-5a7cce1be8e6',
    Secretaria = 'c5df120a-8954-43d7-9f71-fdd0e996ee30',
    Period = '7b4c61b9-c43e-4d38-84f6-f717b2bddcac',
    Community = '172d84e2-ba88-40d6-a872-8fe00d066caa',
    Roles = 'c107d212-3417-472b-a979-b9da999e6adf',
    BusinessArea = 'bda21274-3507-4b59-89e4-64cb5bb1aa0c',
    NullTaxonomy = '00000000-0000-0000-0000-000000000000'
}

export enum RoleTaxonomyIds{
    Convocante = 'fa9dd202-5d9e-48b8-a527-e06fa881b81b',
    GestorConvocante = 'ba768f09-298e-4f59-ac31-97eff41db565',
    members = '9ac2fda6-c5cd-4712-9913-5caaa8f21b66',
    AsistenteMiembro = 'd496453c-b274-4d52-b4b5-9b8156468f7f',
    Invitado = 'fd501b97-9dc5-4d15-9908-de34d26d5b98'
}


export interface IOrganproperties {
    DepartmentName: string;
    Nomenclature: string;
    OrganType: string;
    Active: boolean;
    Secretaria: string;
    Subject: string;
    Legislature: string;
    Abbreviation: string;
    IncorporationDate: Date;
    ExpirationDate: Date;
    Observations: string;
    Description: string;
    Intersectorial: boolean;
    DaysToApprodveTheMinutes: number;
    Dir3: string;
    Sia: string;
    Ministry: string;
    //
    Role?: string;
}

export interface IUserAdministration {
    Id: string;
    DisplayName: string;
    EMail: string;
    Organ: string[];
    CCAA: string;
    Role: string[];
    License: boolean;
}

export enum SearchType {
    None = "None",
    Organ = "Órganos",
    User = "Usuarios"
}

export enum SearchFilterKeys {
    OrganType = 'OrganType',
    User = 'User',
    Organ = 'Organ',
    AutonomousCommunity = 'AutonomousCommunity',
    Rol = 'Rol',
    License = 'License',
    Ministry = 'Ministry',
    BusinessArea = 'BusinessArea'
}

export interface ISearchFilters {
    [SearchFilterKeys.OrganType]: string[];
    [SearchFilterKeys.User]: string;
    [SearchFilterKeys.Organ]: string[];
    [SearchFilterKeys.AutonomousCommunity]: string[];
    [SearchFilterKeys.Rol]: string[];
    [SearchFilterKeys.License]: string[];
    [SearchFilterKeys.Ministry]: string[];
    [SearchFilterKeys.BusinessArea]: string[];
}

export enum LicensesTypes
{
    E3 = "05e9a617-0261-4cee-bb44-138d3ef5d965",
    E5 = "06ebc4ee-1bb5-47dd-8120-11324bc54e06",
    E5Developer = "c42b9cae-ea4f-4ab7-9717-81576235ccac",
    PowerAutomateFree = "f30db892-07e9-47e9-837c-80727f46fd3d",
    PowerBI = "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235",
    Teamstestmium = "36a0f3b3-adb5-49ea-bf66-762134cf063a",
    PowerAppsDeveloper = "5b631642-bd26-49fe-bd20-1daaa972ef80",
}

export enum LicenseNames
{
    "05e9a617-0261-4cee-bb44-138d3ef5d965" = "Microsoft 365 E3",
    "06ebc4ee-1bb5-47dd-8120-11324bc54e06" = "Microsoft 365 E5",
    "c42b9cae-ea4f-4ab7-9717-81576235ccac" = "Microsoft 365 E5"
}

export interface ILicenseInfo {
    SkuId: string;
    Total: string;
    Assigned: string;
    Available: string;
}

export class Administration {
    public static OrganListName: string = '/ConfiguracionDepartments';

    public static eCNTyFilters: ISearchFilters = {
        [SearchFilterKeys.OrganType]: [],
        [SearchFilterKeys.User]: "",
        [SearchFilterKeys.Organ]: [],
        [SearchFilterKeys.AutonomousCommunity]: [],
        [SearchFilterKeys.Rol]: [],
        [SearchFilterKeys.License]: [],
        [SearchFilterKeys.Ministry]: [],
        [SearchFilterKeys.BusinessArea]: []
    };

    public static invitationMessage: string =  `Este es un mensaje enviado por la División de Tecnologías de la Información del Division de Política Territorial y Memoria Democrática para invitarle a participar en la plataforma service-account@contoso.local.

Para ayudarle en el primer acceso a la plataforma ponemos a su disposición el siguiente manual explicativo:  https://CNTmd.gob.es/content/dam/CNT/politica-territorial/autonomica/coop_autonomica/Contoso/PTCAAPP_AccesoContoso.pdf 

Por favor, lea atentamente el manual antes de prodceder a aceptar esta invitación.`
}