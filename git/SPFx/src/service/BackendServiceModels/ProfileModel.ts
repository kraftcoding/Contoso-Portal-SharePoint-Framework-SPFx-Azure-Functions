export interface Iprofile{
    FirstName: string;
    LastName: string;
    Email: string;
    CellPhone: string;
    JobTitle: string;
    Office: string;
    PrincipalMail: string;
    BusinessPhone: string;
}

export enum profileFields {
    CellPhone = 'CellPhone',
    Email = 'Email',
    JobTitle = 'JobTitle',
    Office = 'Office'
}

