import { IVoting } from "../IMeetingVoting";
import {
    IUserAttendance,
    IVote
} from "../../../../../service/BackendServiceModels/EventModels";
import {
    Term
} from "../../../../../models/ITag";

/* ManageResults */
export interface IManageResultsprops {
    updateVoteByRetestsentation: (termId: string, itemSharedId: string, voteId: string) => promise<void>
}

export interface IManageResultsState {
    isMobile: boolean;
    itemInEditing?: Term[];
}

/* ShowQRCode */
export interface IShowQRCodeprops {
    votingSelected?: IVoting;
}

export interface IShowQRCodeState {
    qrUrl: string;
    qrcodedata: any;
    errorMsg: string;
}

/* StartVoting */
export interface IStartVotingprops {
    startVotation: () => promise<void>;
    buttonValue: string;
}
export interface IStartVotingState {
    attendees: IUserAttendance[];
    isLoading: boolean;
    selectedVotingId: string;
}

/* Vote */
export interface IVoteprops {
    sendVote: (vote: IVote) => void;
    refresh(): promise<void>;
    selectedVotingId: string;
}

/* Voting Status */
export interface IVotingStatusprops {
    clickSave: (itemSharedId: string, newStatusId: string)=> void;
}
export interface IVotingStatusState {
    selectedStatus?: string;
}

export interface IVoteState {
    selectedButton: string;
    secondScreen: boolean;
    errorMsg: boolean;
    isLoading: boolean;
    selectedVotingId: string;
}