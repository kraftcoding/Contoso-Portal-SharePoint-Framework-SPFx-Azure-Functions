import { IEventAgenda } from "../../../models/IEventAgenda";


export interface IAgendaAnnouncementState{
    userEvents: IEventAgenda[];
    isOCprodle: boolean;
    isLoading: boolean;
}