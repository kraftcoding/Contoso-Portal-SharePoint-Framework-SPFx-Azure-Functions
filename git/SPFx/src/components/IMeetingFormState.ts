import { FormikValues } from "formik/dist/types";
import { IUser } from "../models/ICall";
import { Term } from "./../models/ITag";

export interface IMeetingFormState {
    onlineTools: Term[];
    attendanceTypes: Term[];
    keepTheFormOpen: boolean;
    eventToolId: string;
    submitting: boolean;
    attendanceTypeSelected: string;
    members: IUser[];
    initialFormFieldValues: FormikValues;
    formDataLoading: boolean;
    membersGroupId: string;
    guestsGroupId: string;
    convoGroupId: string;
    gestorGroupId: string;
    asistenteGroupId: string
}
