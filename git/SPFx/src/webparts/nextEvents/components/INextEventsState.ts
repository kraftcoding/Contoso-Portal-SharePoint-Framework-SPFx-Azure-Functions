import { Term } from '../../../models/ITag';
import { IEvent } from '../../../service/BackendServiceModels/EventModels';

export interface INextEventsState {
    nextEvents: IEvent[];
    loadingNextEvents: boolean;
    bodiesNames: Term[];
    attendancesFormatTypes: Term[];
}