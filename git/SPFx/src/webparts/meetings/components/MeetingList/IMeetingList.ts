import { Term } from "../../../../models/ITag";
import { IEvent } from "../../../../service/BackendServiceModels/EventModels";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMeetingListprops {
    context: WebPartContext;
    events: IEvent[];
    selectedEvent?: IEvent;
    onSelectEvent: (event?: IEvent) => void;
    statusTerms: Term[];
    match: { isExact: boolean, params: { id: string; } };
    loadingChanges: boolean;
    loadingEventId: string;
    isEditor: boolean;
    showArchived: boolean;
}

export interface IMeetingListState {
    nextEvents: IEvent[];
    pastEvents: IEvent[];
    archivedEvents: IEvent[];
    selectedEvents: IEvent[];
    selectedTab?: string;
    isMobile: boolean;
}