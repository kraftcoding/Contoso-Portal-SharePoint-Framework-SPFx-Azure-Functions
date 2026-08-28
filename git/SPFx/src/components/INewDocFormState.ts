import { IFileInfo } from "@pnp/sp/files";

export interface INewDocFormState {
    isLoading: boolean;
    isSaving: boolean;
    files: IFileInfo[];
    selectedFile?: any;
    fileError: string;
}