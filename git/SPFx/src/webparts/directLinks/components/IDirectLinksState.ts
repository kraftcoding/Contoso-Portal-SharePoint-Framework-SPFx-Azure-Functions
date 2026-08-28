import { UsefulLink } from "../../../models/IUsefulLink";

export interface IDirectLinksState {
    usefulLinks: UsefulLink[],
    loadingUsefulLinks: boolean
}