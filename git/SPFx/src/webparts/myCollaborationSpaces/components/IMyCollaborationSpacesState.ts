import { Teams } from "../../../models/ITeams";

export interface IMyCollaborationSpacesState {
    teams: Teams[];
    loading: boolean;
}