import { WebPartContext } from "@microsoft/sp-webpart-base";
import { cacheGetOrPut } from './../utils/Utils';
import { ISPService } from "./Service";

// must be same as backend 
const bodyNamePattern = /(?<bodyId>(?<parent>[a-z]+)-(?<type>[cs|co|gr]+)-(?<body>[a-z]+))(?:-(?<team>[tm]+)?)?(?:_(?<role>[a-z]+))?/i;

export enum BodyRole {
    Undefined = 0, // means that the role is not controlled by the system (other user groups)
    Admin = 1 << 0,
    Scheduler = 1 << 1,
    SchedulerAssistant = 1 << 2,
    Member = 1 << 3,
    MemberAssistant = 1 << 4,
    Guest = 1 << 5
}

export enum BodyType {
    Undefined,
    SectorialComission,
    WorkingGroup,
    Conference
}

export interface IUserBodyRole {
    bodyId: string;
    bodyRelativeUrl: string;
    bodyType: BodyType;
    userRoles: BodyRole[];
}

export enum RoleGroups {
    OficinaConferenciatestsidentes = "contoso-operations"
}

export async function userIs(context: WebPartContext, spService: ISPService, role: BodyRole): promise<boolean> {
    return await userIsAnyOf(context, spService, [role]);
}

export async function userIsAnyOf(context: WebPartContext, spService: ISPService, role: BodyRole[]): promise<boolean> {
    const currentBodyId = getBodyIdFormUrl(context.pageContext.web.absoluteUrl);
    return await userIsAnyOfByBodyId(spService, currentBodyId, role);
}

export async function userIsAnyOfByBodyId(spService: ISPService, bodyId: string, role: BodyRole[]): promise<boolean> {
    const myRoles = await userBodyRoles(spService);
    return myRoles.some(item => item.bodyId === bodyId && role.some(roleItem => item.userRoles.some(userRoleItem => userRoleItem === roleItem)));
}

export async function userBodyRoles(spService: ISPService): promise<IUserBodyRole[]> {
    return await cacheGetOrPut<IUserBodyRole[]>(`roles`, async () => {

        const myGroups = await spService.getMyAadGroups();
        const groupsDisplayName = myGroups.map(group => group.displayName);
        const parsedGroups = parseCollectionGroupNamesUsingNamePattern(groupsDisplayName);

        const result: IUserBodyRole[] = [];
        for (const group of parsedGroups) {
            const bodyIndex = result.findIndex(item => item.bodyId === group.bodyId);
            if (bodyIndex >= 0) {
                result[bodyIndex].userRoles.push(group.userRole);
            } else {
                result.push({
                    bodyId: group.bodyId,
                    bodyRelativeUrl: group.relativeUrl,
                    bodyType: group.type,
                    userRoles: [group.userRole]
                });
            }
        }

        return result;
    });
}

export async function userIsRoleFromGroupName(spService: ISPService, groupName: string): promise<boolean>{
    const myGroups = await spService.getMyAadGroups();
    const groupsDisplayName = myGroups.map(group => group.displayName);
    return groupsDisplayName.some(gr => gr ===groupName);
}

export function urlIsBodyTeamSite(url: string): boolean {
    const bodyInfo = parseUsingNamePattern(url);
    if(!bodyInfo || !bodyInfo.isTeam) return false;
    return true;
}

export function bodyIdToSiteUrl(bodyId: string): string {
    return `/sites/${bodyId}`;
}

export function getBodyIdFormUrl(url: string): string {
    const bodyInfo = parseUsingNamePattern(url);
    return bodyInfo ? bodyInfo.bodyId : '';
}

function typeFromString(type: string): BodyType {
    switch (type.toLowerCase()) {
        case "cs":
            return BodyType.SectorialComission;
        case "co":
            return BodyType.Conference;
        case "gr":
            return BodyType.WorkingGroup;
        default:
            return BodyType.Undefined;
    }
}

function roleFromString(role: string): BodyRole {
    switch (role.toLowerCase()) {
        case "administrators":
            return BodyRole.Admin;
        case "schedulers":
            return BodyRole.Scheduler;
        case "gestorschedulers":
            return BodyRole.SchedulerAssistant;
        case "members":
            return BodyRole.Member;
        case "asistentemembers":
            return BodyRole.MemberAssistant;
        case "guests":
            return BodyRole.Guest;
        default:
            return BodyRole.Undefined;
    }
}

function parseUsingNamePattern(value: string): { bodyId: string, relativeUrl: string, type: BodyType, userRole: BodyRole, isTeam: boolean } | undefined {
    const match = value && value.match(bodyNamePattern);
    if (match) {
        return {
            bodyId: match.groups?.bodyId || '',
            relativeUrl: bodyIdToSiteUrl(match.groups?.bodyId || ''),
            type: typeFromString(match.groups?.type || ''),
            userRole: roleFromString(match.groups?.role || ''),
            isTeam: match.groups?.team === 'tm'
        };
    }
    return undefined;
}

function parseCollectionGroupNamesUsingNamePattern(values: string[]): { bodyId: string, relativeUrl: string, type: BodyType, userRole: BodyRole }[] {
    return values.map(value => parseUsingNamePattern(value)).filter(item => item !== undefined) as { bodyId: string, relativeUrl: string, type: BodyType, userRole: BodyRole }[];
}