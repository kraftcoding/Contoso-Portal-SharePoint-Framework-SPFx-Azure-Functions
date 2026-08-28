import { Term } from "../../../models/ITag";
import { IBody } from "../../../service/BackendServiceModels/BodyModel";

export interface ImyDepartmentsState {
    myDepartments: IBody[],
    loadingmyDepartments: boolean,
    bodiesTypes: Term[]
}