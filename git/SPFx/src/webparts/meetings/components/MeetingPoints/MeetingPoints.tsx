import * as React from 'react';
import * as strings from 'MeetingsWebPartStrings';
import styles from './MeetingPoints.module.scss';
import {
    IAgendaItemErrors,
    IMeetingPointsprops,
    IMeetingPointsState,
    MultiDialogType,
    ItemActions,
    ITooltipConfig,
    LineTypes
} from './IMeetingPoints';
import {
    AddRegular,
    ArrowDownload16Regular,
    ArrowDownload24Regular,
    ArrowDownloadRegular,
    ArrowHookDownRightRegular,
    ArrowHookUpRightRegular,
    AttachRegular,
    ChevronDownRegular,
    ChevronUtestgular,
    Clock16Regular, DeleteRegular, Dismiss16Regular,
    DocumentRibbonRegular,
    EditRegular,
    Info20Regular, ListRegular,
    SettingsRegular,
    TextBulletListAddRegular
} from '@fluentui/react-icons';
import {
    Badge,
    Button,
    Dropdown,
    Field,
    Input,
    Link,
    Menu,
    MenuItem,
    MenuList,
    MenuPopover,
    MenuTrigger,
    Option,
    Spinner,
    Textarea,
    Tooltip,
} from '@fluentui/react-components';
import {
    AgendaItemTypes,
    EventStatus,
    IEventAgendaItem,
    OrderTypes
} from '../../../../service/BackendServiceModels/EventModels';
import {
    IsNullOrECNTy,
    getOrderFormatted,
    getTermLabel,
    groupBy
} from '../../../../utils/Utils';
import { Term } from '../../../../models/ITag';
import { ITermInfo } from '@pnp/sp/taxonomy/types';
import { IFileInfo } from '@pnp/sp/files';
import { File } from '@microsoft/mgt-react/dist/es6/spfx';
import { ViewType } from '@microsoft/mgt-spfx';
import { Logger } from '../../../../utils/Logger';
import { find } from '@microsoft/sp-lodash-subset/lib/index';
import NewDocForm from '../../../../components/NewDocForm';
import ConfirmAction from '../../../../components/ConfirmAction';
import { IConfirmActionprops } from '../../../../components/IConfirmActionprops';

const typeSetId: string = '133993de-eefc-4465-b6c0-b1391c04c7dc'; // Taxonomía 'Tipo orden del día'
const orderTypeSetId: string = '18e98841-f04e-42cc-ae7e-3e77554ffb25'; // Taxonomía 'Tipo ordenación orden del día'
const decisiveTermId: string = "c83ab3b1-00a1-4023-9a2f-1697b23c13d5" // Término 'Decisorio'
const coordinateTermId: string = "264eb3ec-23d4-4cc9-8ace-8a643483db0c" // Término 'Decisorio'
const ROOT_NODE: string = "root";

export default class MeetingPoints extends React.Component<IMeetingPointsprops, IMeetingPointsState> {

    constructor(props: IMeetingPointsprops) {
        super(props);
        this.state = {
            isLoading: true,
            isLoadingItemId: undefined,
            agendaItems: [],
            agendaItemsDictionary: {},
            agendaItemTypes: [],
            agendaItemOrderTypes: [],
            openNewDoc: false,
            fileNames: [],
            errors: {},
            validate: true,
            downloadingTheOrderOfTheDay: false,
            isInReOrderMode: false,
            multiDialogState: { isOpen: false, props: {} },
            aNewAgendaItemIsBeingEdited: false
        };
    }

    public async componentDidMount(): promise<void> {
        await this.onInit(true);
    }

    private async onInit(refresh?: boolean): promise<void> {
        const { spService, context, bkService, event } = this.props;
        const { agendaItems } = this.state;
        let { agendaItemTypes, agendaItemOrderTypes } = this.state;
        const localLanguage: string = context.pageContext.cultureInfo.currentUICultureName;

        this.setState({
            isLoading: true,
            editAgendaItem: undefined,
            validate: false
        });

        /* Se obtienen las taxonomías y los puntos/subpuntos del orden del día */
        let [itemTypes, orderTypes, agendaItemsRes]: [ITermInfo[] | Term[], ITermInfo[] | Term[], IEventAgendaItem[]] = [[], [], []];
        if (refresh) {
            try {
                const isArchived: boolean = event.StatusId === EventStatus.Archived;
                [itemTypes, orderTypes, agendaItemsRes] = await promise.all([
                    spService.getTaxonomy(typeSetId),
                    spService.getTaxonomy(orderTypeSetId),
                    bkService.getEventAgendaItems(event.BodyId, event.Id, isArchived)
                ]);
                agendaItemTypes = itemTypes.map((tag: ITermInfo): Term => Term.mapSPToTerm(tag, localLanguage));
                agendaItemOrderTypes = orderTypes.map((tag: ITermInfo): Term => Term.mapSPToTerm(tag, localLanguage));
            }
            catch (error) {
                this.setState({ isLoading: false });
                Logger.error("Error MeetingsPoints - onInit - taxonomy and agenenda initialization", error, this.context);
            }
        }
        else {
            agendaItemsRes = agendaItems;
        }

        /* Se almacenan los IDs de los Documents relacionados */
        console.log(agendaItemsRes);
        const docIds: string[] = [];
        agendaItemsRes.forEach((agendaItem: IEventAgendaItem): void => 
            {
                agendaItem.RelatedDocumentsIds.forEach((docId: string): number => docIds.push(docId));
                if(agendaItem.AgreementRelatedCertificateId && !docIds.includes(agendaItem.AgreementRelatedCertificateId)){
                    docIds.push(agendaItem.AgreementRelatedCertificateId)
                }
            });

        /* Se obtiene toda la información de los Documents relacionados */
        let fileNames: IFileInfo[] = [];
        try {
            fileNames = await promise.all(docIds.map((docId: string): promise<IFileInfo> => spService.getFileInfoById(docId)));
        }
        catch (error) {
            Logger.error("Error MeetingsPoints - OnInit - file info initialization", error, context);
        }

        /* Los puntos/subpuntos del orden del día se almacenan en un diccionario */
        const agendaItemsDictionary: Record<string, IEventAgendaItem[]> = this.agendaItemsFromArrayToDictionary([...agendaItemsRes]);

        this.setState({
            isLoading: false,
            isLoadingItemId: undefined,
            agendaItems: agendaItemsRes,
            agendaItemsDictionary,
            agendaItemTypes,
            agendaItemOrderTypes,
            fileNames
        });
    }

    // private async refreshAgendaItems(): promise<void> {
    //     const { spService, bkService, event } = this.props;
    //     const isArchived: boolean = event.StatusId === EventStatus.Archived;

    //     /* Se obtienen los puntos/subpuntos del orden del día */
    //     let agendaItemsRes: IEventAgendaItem[] = [];
    //     try {
    //         agendaItemsRes = await bkService.getEventAgendaItems(event.BodyId, event.Id, isArchived);
    //     }
    //     catch (error) {
    //         Logger.error(error);
    //     }
    //     /* Se almacenan los IDs de los Documents relacionados */
    //     const docIds: string[] = [];
    //     agendaItemsRes.forEach((agendaItem: IEventAgendaItem): void => agendaItem.RelatedDocumentsIds.forEach((docId: string): number => docIds.push(docId)));

    //     /* Se obtiene toda la información de los Documents relacionados */
    //     let fileNames: IFileInfo[] = [];
    //     try {
    //         fileNames = await promise.all(docIds.map((docId: string): promise<IFileInfo> => spService.getFileInfoById(docId)));
    //     }
    //     catch (error) {
    //         Logger.error(error);
    //     }

    //     /* Los puntos/subpuntos del orden del día se almacenan en un diccionario */
    //     const agendaItemsDictionary: Record<string, IEventAgendaItem[]> = this.agendaItemsFromArrayToDictionary([...agendaItemsRes]);

    //     this.setState({
    //         isLoading: false,
    //         isLoadingItemId: undefined,
    //         editAgendaItem: undefined,
    //         agendaItems: agendaItemsRes,
    //         agendaItemsDictionary,
    //         fileNames
    //     });
    // }

    private agendaItemsFromArrayToDictionary(originalItems: IEventAgendaItem[]): Record<string, IEventAgendaItem[]> {
        const sortedAgendaItems: IEventAgendaItem[] = originalItems?.sort((a: IEventAgendaItem, b: IEventAgendaItem): number => a.Order - b.Order);
        const groupedAgendaItems: Record<string, IEventAgendaItem[]> = groupBy(sortedAgendaItems, (o: IEventAgendaItem): string => this.checkParentId(o.ParentId));

        return groupedAgendaItems;
    }

    private agendaItemsFromDictionaryToArray(): IEventAgendaItem[] {
        const { agendaItemsDictionary } = this.state;
        return Object.keys(agendaItemsDictionary).reduce((acc: IEventAgendaItem[], key: string) => acc.concat(agendaItemsDictionary[key]), []);
    }

    private insertInDictionary(siblings: IEventAgendaItem[], parentId: string, dict?: Record<string, IEventAgendaItem[]>): Record<string, IEventAgendaItem[]> {
        let agendaItemsDictionary = dict ?? this.state.agendaItemsDictionary;
        agendaItemsDictionary[parentId] = siblings.sort((a: IEventAgendaItem, b: IEventAgendaItem): number => a.Order - b.Order);
        return agendaItemsDictionary;
    }

    private addAgendaItemToDictionary(parentId: string = ROOT_NODE): void {
        const { agendaItemsDictionary } = this.state;

        /* Se comprueba y se obtiene el parentId correspondiente */
        const checkedParentId: string = this.checkParentId(parentId);

        /* Se obtienen los demás puntos/subpuntos que hay en el nivel/subnivel correspondiente */
        const siblings: IEventAgendaItem[] = agendaItemsDictionary[checkedParentId] || [];

        /* Se establece el orden que va a tener el nuevo punto/subpunto */
        const highestOrder: number = siblings?.length + 1;

        /* Se inicializa un nuevo punto/subpunto del orden del día */
        const newAgendaItem: IEventAgendaItem = {
            Id: '',
            Title: '',
            Description: '',
            Duration: 10,
            Order: highestOrder,
            AgendaItemType: AgendaItemTypes.Decisive,
            RelatedDocumentsIds: [],
            ParentId: checkedParentId,
            OrderType: OrderTypes.Numeric
        }

        /* Se añade el nuevo punto/subpunto al diccionario y se le pone en modo edición */
        agendaItemsDictionary[checkedParentId] = [...siblings, newAgendaItem]

        this.setState({
            agendaItemsDictionary,
            editAgendaItem: newAgendaItem,
            aNewAgendaItemIsBeingEdited: true
        });
    }

    private putOrRemoveAgendaItemInDictionary(agendaItem: IEventAgendaItem, remove: boolean = false): [Record<string, IEventAgendaItem[]>, IEventAgendaItem[]] {
        const { agendaItemsDictionary } = this.state;

        /* Se comprueba y se obtiene el parentId correspondiente */
        const checkedParentId: string = this.checkParentId(agendaItem.ParentId);

        /* Se obtienen los puntos/subpuntos que hay en el nivel/subnivel correspondiente */
        let siblings: IEventAgendaItem[] = agendaItemsDictionary[checkedParentId] || [];

        /* Se obtiene el índice del punto/subpunto prodporcionado */
        const index: number = siblings.findIndex((sibling: IEventAgendaItem): boolean => sibling?.Id === agendaItem?.Id);

        /* Se añade o se elimina un punto/subpunto en el diccionario */
        if (index >= 0) {
            siblings.splice(index, 1);

            if (!remove) {
                siblings.splice(agendaItem?.Order - 1, 0, agendaItem);
            }

            /* Se actualiza el orden de los puntos/subpuntos correspondientes */
            siblings = this.updateListOrder(siblings);
            agendaItemsDictionary[checkedParentId] = siblings;
        }

        return [agendaItemsDictionary, siblings];
    }

    private updateAgendaItemInDictionary(agendaItem: IEventAgendaItem): IEventAgendaItem[] {
        const [agendaItemsDictionary, itemToUpdate]: [Record<string, IEventAgendaItem[]>, IEventAgendaItem[]] = this.putOrRemoveAgendaItemInDictionary(agendaItem);
        this.setState({ agendaItemsDictionary });
        return itemToUpdate;
    }

    private removeAgendaItemFromDictionary(agendaItem: IEventAgendaItem): IEventAgendaItem[] {
        const [agendaItemsDictionary, itemToUpdate]: [Record<string, IEventAgendaItem[]>, IEventAgendaItem[]] = this.putOrRemoveAgendaItemInDictionary(agendaItem, true);
        this.setState({ agendaItemsDictionary, editAgendaItem: undefined })
        return itemToUpdate;
    }

    private onChangeForm(prodp: string, value: string | number): void {
        const { validate, errors } = this.state;
        let { editAgendaItem } = this.state;

        /* Se actualiza el punto/subpunto que se está editando */
        if (editAgendaItem) {
            editAgendaItem = {
                ...editAgendaItem,
                [prodp]: value
            };

            /* Se valida el formulario de edición */
            const newErrors: IAgendaItemErrors = validate ? this.validateEditForm(editAgendaItem) : errors;

            this.setState({ editAgendaItem, errors: newErrors });
        }
    }

    private validateEditForm(editItem: IEventAgendaItem): IAgendaItemErrors {
        const { errors } = this.state;

        /* Se valida el título */
        if (editItem.Title) {
            errors.title = undefined;
        }
        else {
            errors.title = strings.ErrorTitleRequired;
        }

        /* Se valida la duración */
        if (editItem.Duration) {
            if (editItem.Duration <= 0) {
                errors.duration = strings.ErrorDuration;
            }
            else {
                errors.duration = undefined;
            }
        }
        else {
            errors.duration = strings.ErrorDurationRequired;
        }

        return errors;
    }

    private onCancelEditAgendaItem = (): void => {
        const { editAgendaItem, aNewAgendaItemIsBeingEdited } = this.state;
        /* Se elimina el nuevo punto/subpunto del diccionario si se cancela su edición */
        if (editAgendaItem && aNewAgendaItemIsBeingEdited) {
            this.removeAgendaItemFromDictionary(editAgendaItem);
        }

        this.setState({
            editAgendaItem: undefined,
            aNewAgendaItemIsBeingEdited: false,
            errors: {},
            validate: false
        });
    };

    private async saveAgendaItems(items?: IEventAgendaItem[], siblingsParent: string = ""): promise<void> {
        const { event, bkService } = this.props;

        /* Se almacenan en un array los puntos/subpuntos que hay que actualizar */
        const itemsToUpdate: IEventAgendaItem[] = items || this.agendaItemsFromDictionaryToArray();
        const refreshAllItems = IsNullOrECNTy(siblingsParent)

        if (refreshAllItems) {
            this.setState({
                isLoading: true,
                agendaItemsDictionary: {}
            });
        }

        let agendaItems: IEventAgendaItem[] = []
        /* Se actualizan los puntos/subpuntos del orden del día en el backend */
        try {
            agendaItems = await bkService.addOrUpdateEventAgendaItems(event.BodyId, event.Id, itemsToUpdate) || [];
        }
        catch (error) {
            Logger.error("Error MeetingsPoints - saveAgendaItems", error, this.context);
        }

        /* Si solo se actualiza un elemento no se vuelve a renderizar todo, en caso de guardar el reorden se actualizan todos los elementos */
        let state = {}
        if (refreshAllItems) {
            const agendaItemsDictionary = this.agendaItemsFromArrayToDictionary(agendaItems);
            state = {
                agendaItemsDictionary,
                agendaItems
            }
        } else {
            const agendaItemsDictionary = this.insertInDictionary(agendaItems, siblingsParent);
            state = {
                agendaItemsDictionary
            }
        }
        this.setState({
            ...state,
            isLoading: false,
            isLoadingItemId: undefined,
            editAgendaItem: undefined,
            aNewAgendaItemIsBeingEdited: false,
            validate: false
        })

    }

    private onSaveEditAgendaItem = async (): promise<void> => {
        const { editAgendaItem } = this.state;

        if (editAgendaItem) {
            /* Se valida el formulario de edición */
            const errors: IAgendaItemErrors = this.validateEditForm(editAgendaItem);

            if (!errors.title && !errors.duration) {
                /* Se actualiza el punto/subpunto en el diccionario */
                const newAgendaItems: IEventAgendaItem[] = this.updateAgendaItemInDictionary(editAgendaItem);
                const parentId = this.checkParentId(editAgendaItem.ParentId);
                /* Se almacenan en un array los puntos/subpuntos que hay que actualizar */
                // const newAgendaItems: IEventAgendaItem[] = this.agendaItemsFromDictionaryToArray();

                this.setState({
                    isLoadingItemId: editAgendaItem.Id,
                    editAgendaItem: (IsNullOrECNTy(editAgendaItem.Id)) ? undefined : editAgendaItem
                });

                /* Se actualizan los puntos/subpuntos del orden del día en el backend */
                await this.saveAgendaItems(newAgendaItems, parentId);
            }
            else {
                this.setState({
                    errors,
                    validate: true
                });
            }
        }
    };

    private async onDeleteAgendaItem(agendaItem: IEventAgendaItem): promise<void> {
        const { event, bkService } = this.props;
        /* Se elimina el punto/supunto del diccionario */
        const newAgendaItems: IEventAgendaItem[] = this.removeAgendaItemFromDictionary(agendaItem);
        const parentId = this.checkParentId(agendaItem?.ParentId);
        /* Se almacenan en un array los puntos/subpuntos que hay que actualizar */
        // const newAgendaItems: IEventAgendaItem[] = this.agendaItemsFromDictionaryToArray();

        this.setState({ agendaItems: newAgendaItems });

        /* Se elimina el punto/subpunto en el backend */
        try {
            await bkService.deleteEventAgendaItem(event.BodyId, event.Id, agendaItem.Id.toString());
        }
        catch (error) {
            Logger.error("Error MeetingsPoints - onDeleteAgendaItem", error, this.context);
        } finally {

            await this.saveAgendaItems(newAgendaItems, parentId)

            /* Se actualizan los puntos/subpuntos del orden del día en el backend */



            // /* Se vuelven a obtener desde el backend los puntos/subpuntos del orden del día */
            // await this.refreshAgendaItems();
        }
    }

    private enableReorderMode = (): void => {
        this.setState({ isInReOrderMode: true, agendaItems: this.agendaItemsFromDictionaryToArray() });
    }

    private updateListOrder(agendaItemList: IEventAgendaItem[]): IEventAgendaItem[] {
        return agendaItemList?.map((agendaItem: IEventAgendaItem, index: number) => {
            return {
                ...agendaItem,
                Order: (index + 1)
            };
        }) || [];
    }

    private extrMinutesgendaItemAndOrderedList(agendaItemList: IEventAgendaItem[], index: number): [IEventAgendaItem, IEventAgendaItem[]] {
        const [item] = agendaItemList.splice(index, 1);
        return [item, this.updateListOrder(agendaItemList)];
    }

    private moveBetweenLevels(oldParentId: string, oldIndex: number, newParentId: string, newIndex: number): [IEventAgendaItem[], IEventAgendaItem[]] {
        const { agendaItemsDictionary } = this.state;
        /* Se obtiene el nivel de destino y el nivel de origen desde el diccionario */
        const targetLevel: IEventAgendaItem[] = agendaItemsDictionary[newParentId] || [];
        const originLevel: IEventAgendaItem[] = agendaItemsDictionary[oldParentId] || [];

        /* Se extrae el punto/subpunto del nivel de origen y se actualiza el orden */
        const [item, newOriginLvl] = this.extrMinutesgendaItemAndOrderedList(originLevel, oldIndex);

        /* Se coloca el punto/subpunto en la posición indicada en el nivel de destino */
        const newTargetLvl: IEventAgendaItem[] = this.putInPosition(targetLevel, { ...item, ParentId: newParentId }, newIndex);

        return [newOriginLvl, newTargetLvl];
    }

    private updatePosition(agendaItemList: IEventAgendaItem[], index: number, up: boolean): IEventAgendaItem[] {
        const [item] = agendaItemList.splice(index, 1);
        return this.putInPosition(agendaItemList, item, up ? (index - 1) : (index + 1));
    }

    private putInPosition(agendaItemList: IEventAgendaItem[], agendaItem: IEventAgendaItem, index: number): IEventAgendaItem[] {
        agendaItemList.splice(index, 0, agendaItem);
        return this.updateListOrder(agendaItemList);
    }

    private getPosition(parentId: string, agendaItemId: string): number {
        const { agendaItemsDictionary } = this.state;
        const siblings = agendaItemsDictionary[parentId] || [];

        if (!(siblings?.length > 0)) return 0;
        return siblings.findIndex((agendaItem: IEventAgendaItem): boolean => agendaItem.Id === agendaItemId);
    }

    private updateReorderPosition = (agendaItem: IEventAgendaItem, action: ItemActions): void => {
        const { agendaItemsDictionary } = this.state;

        /* Se comprueba y se obtiene el parentId correspondiente */
        const checkedParentId: string = this.checkParentId(agendaItem.ParentId);

        /* Se obtienen los puntos/subpuntos que hay en el nivel/subnivel correspondiente */
        const siblings: IEventAgendaItem[] = agendaItemsDictionary[checkedParentId] || [];

        /* Se obtiene el índice del punto/subpunto prodporcionado */
        const index: number = siblings.findIndex((sibling: IEventAgendaItem): boolean => sibling.Id === agendaItem.Id);
        switch (action) {
            /* Subir o bajar en el mismo nivel */
            case ItemActions.ArrowUp:
            case ItemActions.ArrowDown: {
                const newSiblings: IEventAgendaItem[] = this.updatePosition([...siblings,], index, action === ItemActions.ArrowUp);
                agendaItemsDictionary[checkedParentId] = newSiblings;
                break;
            }
            /* Ascender de nivel */
            case ItemActions.LvlUp: {
                let grandParentId: string = ROOT_NODE;
                if (checkedParentId !== ROOT_NODE) {
                    const parentItem: IEventAgendaItem | undefined = this.agendaItemsFromDictionaryToArray()?.find((item: IEventAgendaItem): boolean => item.Id === checkedParentId);
                    grandParentId = this.checkParentId(parentItem?.ParentId || ROOT_NODE);
                }
                const newIndex: number = (this.getPosition(grandParentId, checkedParentId) + 1);
                const [newOriginLvl, newTargetLvl] = this.moveBetweenLevels(checkedParentId, index, grandParentId, newIndex);
                agendaItemsDictionary[checkedParentId] = newOriginLvl;
                agendaItemsDictionary[grandParentId] = newTargetLvl;
                break;
            }
            /* Descender de nivel */
            case ItemActions.LvlDown: {
                const targetParentId: string = siblings[index - 1].Id;
                const [newOriginLvl, newTargetLvl] = this.moveBetweenLevels(checkedParentId, index, targetParentId, agendaItemsDictionary[targetParentId]?.length || 0);
                agendaItemsDictionary[checkedParentId] = newOriginLvl;
                agendaItemsDictionary[targetParentId] = newTargetLvl;
                break;
            }
        }

        this.setState({ agendaItemsDictionary });
    }

    private checkParentId(id: string | null): string {
        return IsNullOrECNTy(id) ? ROOT_NODE : id as string;
    }

    private getOrderOptions(agendaItem: IEventAgendaItem, orderType: OrderTypes): JSX.Element[] {
        const { agendaItemsDictionary } = this.state;

        /* Se obtienen todos los hijos que tienen el mismo 'ParentId' prodporcionado */
        const children: IEventAgendaItem[] = agendaItemsDictionary[this.checkParentId(agendaItem?.ParentId)] || [];

        /* Se determina el número máximo de opciones a mostrar en el desplegable */
        const orderMaxNumber: number = (IsNullOrECNTy(agendaItem.Id)) ? (children.length + 1) : children.length;

        const orderOptions: JSX.Element[] = [];
        // TODO: Cambiar por un switch y añadir el caso para los números romanos

        for (let i = 1; i <= orderMaxNumber; i++) {
            orderOptions.push(<Option key={i} value={i.toString()}>{getOrderFormatted(i, orderType)}</Option>);
        }

        return orderOptions;
    }

    private checkDisabledTooltips(agendaItem: IEventAgendaItem, maxLevel: boolean, reorder: boolean = false): ITooltipConfig {
        const { isLoadingItemId, isLoading, editAgendaItem, agendaItemsDictionary, downloadingTheOrderOfTheDay } = this.state;
        const isDisabled: boolean = isLoading || isLoadingItemId !== undefined || editAgendaItem !== undefined || downloadingTheOrderOfTheDay;

        const response: ITooltipConfig = Object.keys(ItemActions).reduce((acc: {}, key: string): {} => ({ ...acc, [key]: { isDisabled, message: '' } }), {});

        if (isDisabled) return response;

        const configs = {
            First: {
                isDisabled: true,
                // "No disponible sobre el primer elemento de un nivel."
                message: strings.TooltipMessageFirstElement
            },
            Last: {
                isDisabled: true,
                // "No disponible sobre el último elemento de un nivel."
                message: strings.TooltipMessageLastElement
            },
            DecisorySibling: {
                isDisabled: true,
                // "No disponible cuando su elemento superior es decisivo."
                message: strings.TooltipMessageTopElementIsDecisive
            },
            RootParent: {
                isDisabled: true,
                // "No disponible cuando no existe un nivel superior."
                message: strings.TooltipMessageWithoutTopElements
            },
            IsDecisory: {
                isDisabled: true,
                // "No disponible sobre ordenes decisorias"
                message: strings.TooltipMessageNoSubordersInDecisiveElements
            },
            HasChildren: {
                isDisabled: true,
                // "No disponible sobre ordenes con subórdenes"
                message: strings.TooltipMessageNoDeleteElementoWithSuborders
            },
            MaxLevel: {
                isDisabled: true,
                // "No disponible cuando la orden se encuentra en el ultimo nivel"
                message: "Se ha alcanzado el límite de subniveles"
            }
        }

        if (reorder) {
            /* Se comprueba y se obtiene el parentId correspondiente */
            const checkedParentId: string = this.checkParentId(agendaItem.ParentId);

            /* Se obtienen los puntos/subpuntos que hay en el nivel/subnivel correspondiente */
            const siblings: IEventAgendaItem[] = agendaItemsDictionary[checkedParentId] || [];

            /* Se obtiene el índice del punto/subpunto prodporcionado */
            const index: number = siblings.findIndex((sibling: IEventAgendaItem): boolean => sibling.Id === agendaItem.Id);

            const isRoot: boolean = checkedParentId === ROOT_NODE;
            const isLast: boolean = index === (siblings?.length - 1);
            const isDecisorySibling: boolean = (index > 0) && (siblings[index - 1]?.AgendaItemType === decisiveTermId || siblings[index - 1]?.AgendaItemType === coordinateTermId);
            const firstElement: boolean = agendaItem?.Order === 1;

            if (firstElement) {
                response[ItemActions.ArrowUp] = { ...configs.First };
                response[ItemActions.LvlDown] = { ...configs.First };
            }
            if (isLast) response[ItemActions.ArrowDown] = { ...configs.Last };
            if (isRoot) response[ItemActions.LvlUp] = { ...configs.RootParent };
            if (isDecisorySibling) response[ItemActions.LvlDown] = { ...configs.DecisorySibling };
            if (maxLevel) response[ItemActions.LvlDown] = { ...configs.MaxLevel };
        }
        else {
            const isDecisory: boolean = (agendaItem?.AgendaItemType !== null) && (agendaItem.AgendaItemType === decisiveTermId || agendaItem.AgendaItemType === coordinateTermId);
            const hasChildren: boolean = agendaItemsDictionary[agendaItem.Id]?.length > 0;

            if (isDecisory) response[ItemActions.Add] = { ...configs.IsDecisory };
            if (hasChildren) response[ItemActions.Remove] = { ...configs.HasChildren };
            if (maxLevel) response[ItemActions.Add] = { ...configs.MaxLevel };
        }

        return response;
    }

    private async downloadTheOrderOfTheDay(): promise<void> {
        const { event, bkService } = this.props;

        this.setState({ downloadingTheOrderOfTheDay: true });
        try {
            /* Se obtiene el archivo con el orden del día */
            const result: ArrayBuffer = await bkService.downloadOrderOfTheDayTemplate(event.BodyId, event.Id);
            if (result) {
                /* Se abre en una pestaña nueva el archivo con el orden del día */
                const objectUrl: string = window.URL.createObjectURL(new Blob([result], { type: "application/pdf" }));
                const url: HTMLAnchorElement = document.createElement('a');
                url.href = objectUrl;
                url.setAttribute('download', `${strings.MeetingPoints} - ${event.Title.substring(0, 50)}.pdf`);
                document.body.appendChild(url);
                url.click();
                document.body.removeChild(url);
                window.URL.revokeObjectURL(objectUrl);
            }
        }
        catch (error) {
            console.log(error);
        }
        finally {
            this.setState({ downloadingTheOrderOfTheDay: false });
        }
    }

    private onSaveRelatedDocument = async (fileUniqueId: string): promise<void> => {
        const { bkService, event } = this.props;
        const { sectionNewDoc } = this.state;
        try {
            /* Se añade el Document relacionado correspondiente en el backend */
            await bkService.addRelatedDocumentByUniqueId(event.BodyId, event.Id, sectionNewDoc ? sectionNewDoc.Id.toString() : '', fileUniqueId);
        }
        catch (error) {
            Logger.error("Error MeetingsPoints - onSaveRelatedDocument", error, this.context);
        }
        finally {
            this.setState({
                isLoading: true,
                isLoadingItemId: sectionNewDoc?.Id
            });

            // eslint-disable-next-line no-void
            void this.onInit(true);
        }
    };

    private onSaveRelatedCertificate = async (fileUniqueId: string): promise<void> => {
            const { bkService, event } = this.props;
            const { selectedCertificateAdd } = this.state;
            let agendaItems = this.state.agendaItems;
            try {
                await bkService.addRelatedDocumentIdToAgreement(event.BodyId, event.Id, selectedCertificateAdd ? selectedCertificateAdd.Id.toString() : '', fileUniqueId, true);
                agendaItems = agendaItems.map(item => {
                    if (item.Id === selectedCertificateAdd?.Id) {
                        return {
                            ...item,
                            RelatedCertificateId: fileUniqueId,
                        };
                    }
                    return item;
                });
                this.setState({ isLoadingItemId: undefined, agendaItems, isLoading: true });
                void this.onInit(true);
            }
            catch (error) {
                Logger.error("Error MeetingsAgreements - onSaveRelatedCertificate - saving agreements", error, this.context);
                this.setState({ isLoading: true });
                void this.onInit(true);
            }
        };

    private onOpenCloseDocForm = (item?: IEventAgendaItem): void => {
        this.setState({ openNewDoc: !this.state.openNewDoc, sectionNewDoc: item || this.state.sectionNewDoc , selectedCertificateAdd: undefined});
    };

    private onOpenCloseCertificateForm = (item?: IEventAgendaItem): void => {
        this.setState({ openNewDoc: !this.state.openNewDoc, sectionNewDoc:undefined , selectedCertificateAdd: item || this.state.selectedCertificateAdd });
    };

    private renderRelatedDocumentation(agendaItem: IEventAgendaItem): JSX.Element[] {
        const { context, isEditor, readOnly } = this.props;
        const { fileNames, isLoading, isLoadingItemId, editAgendaItem: editItem } = this.state;
        const locked: boolean = !isEditor || readOnly;
        const documentsToShow = (agendaItem.AgendaItemType === AgendaItemTypes.Informative && agendaItem.AgreementRelatedCertificateId) ? [agendaItem.AgreementRelatedCertificateId].concat(agendaItem.RelatedDocumentsIds)  : [...agendaItem.RelatedDocumentsIds]
        return (
            documentsToShow.map((fileId: string): JSX.Element => {
                const document: IFileInfo | undefined = find(fileNames, fileName => fileName.UniqueId === fileId);
                const blockDismiss: boolean = isLoading || isLoadingItemId !== undefined || editItem !== undefined;

                return (
                    <div className={`${styles.itemRow}`} style={{ padding: "0 5px 0 0", minHeight: "30px" }}>
                        <div className={`${styles.itemColumn} ${styles.justifyCenter}`}>
                            <File
                                className={styles.file}
                                fileDetails={{ name: document?.Name, webUrl: document?.ServerRelativeUrl }}
                                onClick={() => window.open(document?.ServerRelativeUrl + "?web=1", '_blank')}
                                view={ViewType.oneline}
                            >
                            </File>
                        </div>
                        <div className={`${styles.itemColumn} ${styles.justifyCenter}`}>
                            <ArrowDownload16Regular
                                style={{ cursor: 'pointer' }}
                                onClick={() => window.open(`${context.pageContext.site.serverRelativeUrl}/_layouts/download.aspx?SourceUrl=${document?.ServerRelativeUrl}`, '_self')}
                            />
                        </div>
                        {
                            !locked &&
                            <div className={`${styles.itemColumn} ${styles.justifyCenter}`}>
                                <Dismiss16Regular style={{ flexShrink: 0, cursor: 'pointer', display: blockDismiss ? 'none' : 'block' }} onClick={() => this.openDialog(MultiDialogType.RemoveDocument, { document, agendaItem, fileId })} />

                            </div>
                        }
                    </div>
                );
            })
        );
    }

    private openDialog(type: MultiDialogType, props: any = {}) {
        this.setState({ multiDialogState: { isOpen: true, type, props } })
    }

    private renderMultiDialog() {

        const { isOpen, type, props } = this.state.multiDialogState;
        const { document, agendaItem, fileId } = props;

        const closedState = {
            ...this.state.multiDialogState,
            isOpen: false
        }
        const onSave = (func: () => any) => this.setState({ multiDialogState: closedState }, () => func())

        const cancelButtonprops = {
            appearance: "secondary",
            title: strings.RuleOut,
            onClick: (): void => {
                this.setState({
                    multiDialogState: closedState
                });
            }
        }
        const dialogConfig: { [key: string]: any } = {
            [MultiDialogType.SaveReorder]: {
                dialogTitle: strings.SaveChangesMade,
                dialogContent: strings.SaveChangesMessage,
                acceptButtonprops: {
                    appearance: "primary",
                    title: strings.SaveOrdenation,
                    onClick: (): void => {
                        this.setState({
                            isLoading: true,
                            isInReOrderMode: false,
                            multiDialogState: closedState
                        }, () => this.saveAgendaItems());
                    }
                }
            },
            [MultiDialogType.CancelReorder]: {
                dialogTitle: strings.DiscardChangesTitle,
                dialogContent: strings.CancelReorderModeMessage,
                acceptButtonprops: {
                    appearance: 'primary',
                    title: strings.CancelOrdenation,
                    onClick: (): void => {
                        this.setState({
                            isInReOrderMode: false,
                            multiDialogState: closedState,
                            // Se recupera el orden del día testvio a la ordenación
                            agendaItemsDictionary: this.agendaItemsFromArrayToDictionary(this.state.agendaItems)
                        });
                    }
                },
            },
            [MultiDialogType.RemoveDocument]: {
                dialogTitle: strings.DeleteAssociation,
                dialogContent: <div>{strings.DeleteAssociatedDoc}<strong>{document?.Name}</strong> {strings.FromTheOrderOfTheDay} <strong>{agendaItem?.Title}</strong>?</div>,
                acceptButtonprops: {
                    appearance: 'primary',
                    icon: <DeleteRegular />,
                    title: strings.Delete,
                    onClick: () => onSave(() => this.removeRelatedDocumentation(agendaItem?.Id?.toString() ?? '', fileId ?? ''))
                },
            },
            [MultiDialogType.RemoveOrder]: {
                dialogTitle: strings.DeleteTheOrderOfTheDay,
                dialogContent: <div>{strings.DoYouWantToDeleteTheOrderOfTheDay}<strong>{agendaItem?.Title}</strong>?</div>,
                acceptButtonprops: {
                    appearance: "primary",
                    icon: <DeleteRegular />,
                    title: strings.Delete,
                    onClick: () => onSave(() => this.onDeleteAgendaItem(agendaItem as IEventAgendaItem))
                }

            }
        }

        const open = type !== undefined && isOpen;
        const config: IConfirmActionprops = { ...dialogConfig[type ?? MultiDialogType.CancelReorder], cancelButtonprops }

        return (<ConfirmAction
            dialogprops={{ open }}
            {...config}

        />)
    }

    private async removeRelatedDocumentation(agendaItemId: string, uniqueId: string): promise<void> {
        const { bkService, event } = this.props;
        let { agendaItems } = this.state;

        let isCertificate:boolean = false;
        agendaItems = agendaItems.map((agendaItem: IEventAgendaItem): IEventAgendaItem => {
            if (agendaItem.Id.toString() === agendaItemId) {
                if(agendaItem.AgreementRelatedCertificateId && agendaItem.AgreementRelatedCertificateId === uniqueId){
                    isCertificate = true;
                }
                return {
                    ...agendaItem,
                    RelatedDocumentsIds: agendaItem.RelatedDocumentsIds.filter(relatedDoc => relatedDoc !== uniqueId),
                    AgreementRelatedCertificateId: isCertificate ? '' : agendaItem.AgreementRelatedCertificateId,
                };
                //return { ...agendaItem, RelatedDocumentsIds: agendaItem.RelatedDocumentsIds.filter((relatedDoc: string): boolean => relatedDoc !== uniqueId) };
            }
            return agendaItem;
        });

        this.setState({
            agendaItems,
            isLoading: true,
            isLoadingItemId: agendaItemId
        });

        /* Se quita el Document relacionado correspondiente en el backend */
        try {
            if(isCertificate){
                await bkService.removeRelatedDocumentIdToAgreement(event.BodyId, event.Id, agendaItemId, uniqueId);
            }else{
                await bkService.removeRelatedDocumentByUniqueId(event.BodyId, event.Id, agendaItemId, uniqueId);
            }
        }
        catch (error) {
            Logger.error("Error MeetingsPoints - removeRelatedDocumentation", error, this.context);
        }
        finally {
            // eslint-disable-next-line no-void
            void this.onInit(true);
        }
    }

    private renderRecursiveAgendaItems(node: string = "root", lines: boolean[] = [], orderType: OrderTypes = OrderTypes.Numeric): React.ReactElement[] {
        try {
            const { agendaItemsDictionary: agendaItemsDic, editAgendaItem: editItem } = this.state;
            const nodes: IEventAgendaItem[] = agendaItemsDic[node];


            return (
                nodes?.map((item: IEventAgendaItem, index: number): JSX.Element => {
                    const auxLines = node === "root" ? [] : [...lines, index < nodes?.length - 1]
                    return (
                        <div key={item?.Id} className={`${styles.itemColumn} ${styles.columnBlueBackground} `}>
                            {
                                (editItem && editItem.Id === item.Id) ?
                                    this.renderAgendaItemEdit(item, auxLines, orderType)
                                    :
                                    this.renderAgendaItemView(item, auxLines, orderType)
                            }

                            {!IsNullOrECNTy(item?.Id) && this.renderRecursiveAgendaItems(item?.Id, auxLines, IsNullOrECNTy(item?.OrderType) ? OrderTypes.Numeric : item?.OrderType as OrderTypes)}
                        </div>
                    );
                }) || []
            );
        } catch (ex) {
            Logger.error("Error to render recursive agenda items", ex);
        }
        return [];
    }

    private getLinesByType(type: LineTypes): JSX.Element {
        switch (type) {
            case LineTypes.T:
                /*
                    * * * * *
                    *   |   *
                    *   |---*
                    *   |   *
                    * * * * *
                */
                return (
                    <div className={`${styles.itemColumn} ${styles.treeLines}`}>
                        <div></div>
                        <div className={`${styles.borderBottom} ${styles.borderLeft}`}></div>
                        <div></div>
                        <div className={`${styles.borderLeft}`}></div>
                    </div>
                );
            case LineTypes.L:
                /*
                    * * * * *
                    *   |   *
                    *   |__ *
                    *       *
                    * * * * *
                */
                return (
                    <div className={`${styles.itemColumn} ${styles.treeLines}`}>
                        <div></div>
                        <div className={`${styles.borderBottom} ${styles.borderLeft}`}></div>
                        <div></div>
                        <div></div>
                    </div>
                );
            case LineTypes.I:
                /*
                    * * * * *
                    *   |   *
                    *   |   *
                    *   |   *
                    * * * * *
                */
                return (
                    <div className={`${styles.itemColumn} ${styles.treeLines}`}>
                        <div></div>
                        <div className={styles.borderLeft}></div>
                        <div></div>
                        <div className={styles.borderLeft}></div>
                    </div>
                );
            case LineTypes.S:
            default:
                return (
                    <div className={`${styles.itemColumn} ${styles.treeLines}`}>
                    </div>
                );
        }
    }

    private renderAgendaItemEdit(agendaItem: IEventAgendaItem, infoLines?: boolean[], orderType: OrderTypes = OrderTypes.Numeric): React.ReactElement {
        let treeLines: JSX.Element[] = [];

        if (!IsNullOrECNTy(agendaItem.ParentId) && infoLines) {
            treeLines = infoLines?.map((line: boolean, index: number): JSX.Element => {
                if (index !== (infoLines.length - 1)) {
                    return this.getLinesByType(line ? LineTypes.I : LineTypes.S);

                } else {
                    return this.getLinesByType(line ? LineTypes.T : LineTypes.L);
                }
            })
        }

        return (
            <div
                className={`${styles.itemColumn} ${styles.columnBlueBackground}`}
                ref={(node) => {
                    node?.scrollIntoView({
                        behavior: "smooth",
                        block: "nearest"
                    })
                }}
            >
                <div className={`${styles.itemRow}`} style={{ paddingBottom: "5px", gap: "0px" }}>
                    {/* Líneas de jerarquía */}
                    <div className={`${styles.itemRow}`} style={{ gap: "0px" }}>
                        {treeLines}
                    </div>
                    {/* Punto/subpunto en modo edición */}
                    {this.renderAgendaItemEditCard(agendaItem, orderType)}
                </div>
            </div>
        );
    }

    private renderAgendaItemEditCard(agendaItem: IEventAgendaItem, orderType: OrderTypes): JSX.Element {
        const { agendaItemTypes, editAgendaItem, errors, isLoadingItemId, isLoading, agendaItemOrderTypes, agendaItemsDictionary } = this.state;
        const handleChangeInput = (ev: { target: { name: string; }; }, data: { value: string; }): void => this.onChangeForm(ev.target.name, data.value);
        const disabled: boolean = isLoading || isLoadingItemId !== undefined;
        const checkedParentId: string = this.checkParentId(agendaItem?.ParentId);
        const hasChildren: boolean = agendaItemsDictionary[agendaItem.Id]?.length > 0;


        return (
            /* Tarjeta de un punto/subpunto en modo edición */
            <div className={`${checkedParentId !== ROOT_NODE ? styles.borderAgendaItem : ''}`} style={{ flex: 1, padding: "0 5px" }}>
                {/* Primera fila */}
                <div className={`${styles.itemRow} ${styles.topPaddingAdjust}`}>
                    {/* Índice */}
                    <div className={styles.itemColumn}>
                        {<Dropdown
                            className={styles.autoWidth}
                            disabled={disabled}
                            value={getOrderFormatted(editAgendaItem?.Order ?? 1, orderType)}
                            // selectedOptions={editAgendaItem?.Order ? [editAgendaItem.Order.toString()] : []}
                            onOptionSelect={(ev, data): void => { console.log(data); this.onChangeForm('Order', parseInt(data.optionValue as string)) }}
                        // onOptionSelect={(ev, data): void => }
                        >
                            {this.getOrderOptions(agendaItem, orderType)}
                        </Dropdown>
                        }
                    </div>
                    {/* Tipo (decisorio o informativo) */}
                    <div className={styles.itemColumn}>
                        <Dropdown
                            className={styles.autoWidth}
                            disabled={disabled || hasChildren}
                            value={getTermLabel(agendaItemTypes, editAgendaItem?.AgendaItemType as string)}
                            selectedOptions={(editAgendaItem && editAgendaItem.AgendaItemType) ? [editAgendaItem.AgendaItemType] : []}
                            onOptionSelect={(ev, data): void => this.onChangeForm('AgendaItemType', data.optionValue as string)}
                        >
                            {
                                agendaItemTypes?.map((itemType: Term): JSX.Element =>
                                    <Option key={itemType.key} value={itemType.key}>
                                        {itemType.text}
                                    </Option>
                                )
                            }
                        </Dropdown>
                    </div>
                    {/* Duración */}
                    <div className={styles.itemColumn}>
                        <Field validationMessage={errors.duration}>
                            <Input
                                type="number"
                                name="Duration"
                                contentAfter={"min"}
                                contentBefore={<Clock16Regular />}
                                min={1}
                                maxLength={2}
                                disabled={disabled}
                                className={styles.inputDuration}
                                onChange={handleChangeInput}
                                value={editAgendaItem?.Duration.toString()}
                            />
                        </Field>
                    </div>
                    {/* Tipo de ordenación de los hijos (alfabético, numérico o números romanos) */}
                    <div className={styles.itemColumn}>
                        {
                            hasChildren &&
                            <Dropdown
                                className={styles.autoWidth}
                                disabled={disabled}
                                value={getTermLabel(agendaItemOrderTypes, editAgendaItem?.OrderType as string)}
                                selectedOptions={(editAgendaItem && editAgendaItem.OrderType) ? [editAgendaItem.OrderType] : []}
                                onOptionSelect={(ev, data): void => this.onChangeForm('OrderType', data.optionValue as string)}
                            >
                                {
                                    agendaItemOrderTypes?.map((orderType: Term): JSX.Element =>
                                        <Option key={orderType.key} value={orderType.key}>
                                            {orderType.text}
                                        </Option>
                                    )
                                }
                            </Dropdown>
                        }
                    </div>
                </div>
                {/* Segunda fila */}
                <div className={`${styles.itemRow} ${styles.topPaddingAdjust}`}>
                    {/* Título */}
                    <div className={styles.itemColumnGrow}>
                        <Field validationMessage={errors.title}>
                            <Input
                                disabled={disabled}
                                maxLength={255}
                                name='Title'
                                type='text'
                                value={editAgendaItem?.Title}
                                onChange={handleChangeInput}
                                placeholder={strings.Title}
                            />
                        </Field>
                    </div>
                </div>
                {/* Tercera fila */}
                <div className={`${styles.itemRow} ${styles.topPaddingAdjust}`}>
                    {/* Descripción */}
                    <div className={styles.itemColumnGrow}>
                        <Textarea
                            disabled={disabled}
                            className={styles.inputDesc}
                            name='Description'
                            value={editAgendaItem?.Description || ''}
                            onChange={handleChangeInput}
                            placeholder={strings.Description}
                        />
                    </div>
                </div>
                {/* Cuarta fila */}
                <div className={`${styles.itemRow} ${styles.topPaddingAdjust} ${styles.justifyEnd}`}>
                    {disabled && <Spinner size='tiny' />}
                    {/* Botón para cancelar la edición de un punto/subpunto */}
                    <Button
                        disabled={disabled}
                        appearance="secondary"
                        onClick={this.onCancelEditAgendaItem}
                    >
                        {strings.Cancel}
                    </Button>
                    {/* Botón para guardar la edición de un punto/subpunto */}
                    <Button
                        disabled={!(!errors.title && !errors.duration) || disabled}
                        appearance="primary"
                        onClick={this.onSaveEditAgendaItem}
                    >
                        {IsNullOrECNTy(editAgendaItem?.Id) ? strings.Create : strings.Save}
                    </Button>
                </div>
            </div >
        );
    }

    private renderAgendaItemView(agendaItem: IEventAgendaItem, infoLines: boolean[], orderType: OrderTypes): React.ReactElement {
        let treeLines: JSX.Element[] = [];

        if (!IsNullOrECNTy(agendaItem.ParentId) && infoLines) {
            treeLines = infoLines?.map((line: boolean, index: number): JSX.Element => {
                if (index !== (infoLines.length - 1)) {
                    return this.getLinesByType(line ? LineTypes.I : LineTypes.S);
                }
                else {
                    return this.getLinesByType(line ? LineTypes.T : LineTypes.L);
                }
            });
        }

        return (
            <div style={{ display: "flex", paddingBottom: "5px" }}>
                {/* Líneas de jerarquía */}
                {treeLines}
                {/* Punto/subpunto en modo visualización */}
                {this.renderAgendaItemViewCard(agendaItem, orderType, infoLines?.length >= 4)}
            </div>
        );
    }

    private renderCustomTooltip(props: { defaultMessage: string, icon: any, onClick?: () => void, isDisabled: boolean, message: string, isMobile: boolean }): React.ReactElement {
        const { isDisabled, message, icon, defaultMessage, onClick, isMobile } = props;
        return (isMobile ?
            <MenuItem icon={icon} onClick={onClick} disabled={isDisabled} persistOnClick={false}>{defaultMessage}</MenuItem> :
            <Tooltip
                withArrow
                positioning={'below'}
                content={(isDisabled && !IsNullOrECNTy(message)) ? message : defaultMessage}
                relationship="label"
            ><Button disabled={isDisabled} appearance={"secondary"} icon={icon} onClick={onClick} /></Tooltip>)
    }

    private renderCertificateMenuTooltip(props: { defaultMessage: string, icon: any, onClick?: () => void, isDisabled: boolean, message: string, isMobile: boolean }): React.ReactElement {
        const { isDisabled, icon, defaultMessage, onClick, isMobile } = props;
        return (isMobile ?
            <MenuItem icon={icon} onClick={onClick} disabled={isDisabled} persistOnClick={false}>{defaultMessage}</MenuItem> :
            <Button disabled={isDisabled} appearance={"secondary"} icon={icon} onClick={onClick}>{defaultMessage}</Button>)
    }

    private onUploadCertificate(agendaItem: IEventAgendaItem):void{
        this.onOpenCloseCertificateForm(agendaItem);
    }

    private onDownloadCertificate(agendaItem: IEventAgendaItem):void{
        
        this.downloadCertification(agendaItem).then(res => console.log(res)).catch(ex => console.log(ex));
    }

    private async downloadCertification(agendaItem: IEventAgendaItem): promise<void> {
        const { event, bkService } = this.props;
        this.setState({isLoadingItemId: agendaItem.Id});
        try {
            if (agendaItem) {
                let result: ArrayBuffer = await bkService.downloadAgreementsTemplate(event.BodyId, event.Id, agendaItem.Id.toString());
                if (result) {
                    const url = window.URL.createObjectURL(new Blob([result]));
                    const enlace = document.createElement('a');
                    enlace.href = url;
                    enlace.setAttribute('download', `${strings.Certification} - ${event.Title.substring(0, 50)} - ${agendaItem.Title.substring(0, 25)}.docx`);
                    document.body.appendChild(enlace);
                    enlace.click();
                    document.body.removeChild(enlace);
                    window.URL.revokeObjectURL(url);
                }
            }
            this.setState({isLoadingItemId: undefined});
        }
        catch (error) {
            this.setState({isLoadingItemId: undefined});
            Logger.error("Error MeetingsAgreements - downloadCertification - downloading template", error, this.context);
        }
    }

    private renderTooltipsMenu(agendaItem: IEventAgendaItem, config: ITooltipConfig, reorder: boolean = false) {
        const { isLoadingItemId, isLoading, editAgendaItem, downloadingTheOrderOfTheDay } = this.state;
        const isDisabled: boolean = isLoading || isLoadingItemId !== undefined || editAgendaItem !== undefined || downloadingTheOrderOfTheDay;
        const isMobile: boolean = window.innerWidth <= 640;
        let actions: { [action: string]: { defaultMessage: string, icon: any, onClick: () => void } } = {}
        let certificateActions: { [action: string]: { defaultMessage: string, icon: any, onClick: () => void } } = {}
        let tooltipsMenu: React.ReactElement<any, string | React.JSXElementConstructor<any>>[]= [];

        if (reorder) {
            actions = {
                [ItemActions.LvlUp]: { defaultMessage: strings.prodmoteDifferentLevel, icon: <ArrowHookUpRightRegular />, onClick: () => this.updateReorderPosition(agendaItem, ItemActions.LvlUp) },
                [ItemActions.LvlDown]: { defaultMessage: strings.DecreaseDifferentLevel, icon: <ArrowHookDownRightRegular />, onClick: () => this.updateReorderPosition(agendaItem, ItemActions.LvlDown) },
                [ItemActions.ArrowUp]: { defaultMessage: strings.UpSameLevel, icon: <ChevronUtestgular />, onClick: () => this.updateReorderPosition(agendaItem, ItemActions.ArrowUp) },
                [ItemActions.ArrowDown]: { defaultMessage: strings.DownSameLevel, icon: <ChevronDownRegular />, onClick: () => this.updateReorderPosition(agendaItem, ItemActions.ArrowDown) },
            }
        } else {
            if(agendaItem.AgendaItemType === AgendaItemTypes.Informative){
                if(!agendaItem.AgreementRelatedCertificateId){
                    certificateActions = {
                        [ItemActions.UploadCertificate]:{ defaultMessage: strings.AssociateSignedCertificate, icon: <DocumentRibbonRegular />, onClick: () => this.onUploadCertificate(agendaItem) },
                        [ItemActions.DownloadCertificate]:{ defaultMessage: strings.DownloadCertification, icon: <ArrowDownload24Regular />, onClick: () => this.onDownloadCertificate(agendaItem) },
                    }
                }else{
                    certificateActions = {
                        [ItemActions.DownloadCertificate]:{ defaultMessage: strings.DownloadCertification, icon: <ArrowDownload24Regular />, onClick: () => this.onDownloadCertificate(agendaItem) },   
                    }
                }
                
            }
            actions = {
                [ItemActions.Edit]: { defaultMessage: strings.Edit, icon: <EditRegular />, onClick: () => this.setState({ editAgendaItem: agendaItem }) },
                [ItemActions.Attach]: { defaultMessage: strings.AssociateDocument, icon: <AttachRegular />, onClick: () => this.onOpenCloseDocForm(agendaItem) },
                [ItemActions.Remove]: { defaultMessage: strings.Delete, icon: <DeleteRegular />, onClick: () => this.openDialog(MultiDialogType.RemoveOrder, { agendaItem }) },
            }
        }

        let certificateMenu: React.ReactElement<any, string | React.JSXElementConstructor<any>>[]= [];
        certificateMenu=[...certificateMenu, ...Object.keys(certificateActions).map(key => { return this.renderCertificateMenuTooltip({ ...certificateActions[key], ...config[key], isMobile })})];
        if(certificateMenu && certificateMenu.length > 0){
            const certificationMenu = isMobile? certificateMenu : (
            <Menu>
                <MenuTrigger>
                    <Tooltip
                    withArrow
                    positioning={'below'}
                    content={strings.Certification}
                    relationship="label"
                    >
                        <Button appearance={"secondary"} icon={<DocumentRibbonRegular />} disabled={isDisabled}/>
                    </Tooltip>
                </MenuTrigger>
                <MenuPopover>
                    <MenuList style={{gap:10, padding:10}}>
                        {certificateMenu}
                    </MenuList>
                </MenuPopover>
            </Menu>
            )
            tooltipsMenu = tooltipsMenu.concat(certificationMenu)
        }

        tooltipsMenu=[...tooltipsMenu,...Object.keys(actions).map(key => { return this.renderCustomTooltip({ ...actions[key], ...config[key], isMobile }) })];

        if (!isMobile) return tooltipsMenu;

        return (<Menu>
            <MenuTrigger>
                <Button appearance={"secondary"} icon={<SettingsRegular />} disabled={isDisabled} />
            </MenuTrigger>
            <MenuPopover>
                <MenuList>
                    {tooltipsMenu}
                </MenuList>
            </MenuPopover>
        </Menu>)
    }

    private renderAgendaItemViewCard(agendaItem: IEventAgendaItem, orderType: OrderTypes, maxLevel: boolean = false): JSX.Element {
        const { agendaItemTypes, isLoadingItemId, isInReOrderMode } = this.state;
        const { isEditor, readOnly } = this.props;
        const locked: boolean = !isEditor || readOnly;
        const tooltipsConfig: ITooltipConfig = this.checkDisabledTooltips(agendaItem, maxLevel, isInReOrderMode);
        const checkedParentId: string = this.checkParentId(agendaItem.ParentId);

        
        return (
            /* Tarjeta de un punto/subpunto en modo visualización */
            <div key={agendaItem.Id} className={`${styles.itemColumn} ${styles.columnBlueBackground} ${styles.sidesPaddingAdjust} ${checkedParentId !== ROOT_NODE ? styles.borderAgendaItem : ''}`} style={{ width: '100%' }}>
                {/* Primera fila */}
                <div className={styles.agendaItemViewHeader}>
                    {/* Iconos */}
                    <div className={styles.agendaItemViewHeaderLeft}>
                        <div className={`${styles.itemRow} ${styles.topPaddingAdjust} ${styles.agendaItemViewHeaderLeftElements}`}>
                            {/* Índice */}
                            <div className={`${styles.itemColumn} ${styles.justifyCenter} ${styles.agendaItemViewHeaderLeftElementsIndex}`}>
                                <Badge className={styles.order}>
                                    {getOrderFormatted((agendaItem?.Order), orderType)}
                                </Badge>
                            </div>
                            {/* Tipo */}
                            <div className={`${styles.itemColumn} ${styles.justifyCenter} ${styles.agendaItemViewHeaderLeftElementsType}`}>
                                {
                                    agendaItem.AgendaItemType &&
                                    <Badge appearance='outline'>{getTermLabel(agendaItemTypes, agendaItem.AgendaItemType)}</Badge>
                                }
                            </div>
                            {/* Duración */}
                            <div className={`${styles.itemColumn} ${styles.agendaItemViewHeaderLeftElementsDuration}`} >
                                <Clock16Regular />
                                {`${agendaItem.Duration} min`}
                            </div>
                        </div>
                    </div>
                    {/* Botones */}
                    <div className={styles.agendaItemViewHeaderRight}>
                        {
                            /* Botones principales */
                            <div className={`${styles.itemColumnGrow}`}>
                                <div className={`${styles.itemRow} ${styles.justifyEnd} ${styles.agendaItemViewHeaderRightMainButtons}`}>
                                    {isLoadingItemId === agendaItem.Id && <Spinner size="tiny" />}

                                    {!locked && this.renderTooltipsMenu(agendaItem, tooltipsConfig, isInReOrderMode)}

                                </div>
                            </div>
                        }
                    </div>
                </div>

                {/* Segunda fila */}
                <div className={`${styles.itemRow} ${styles.title} ${styles.topPaddingAdjust}`}>
                    {/* Título */}
                    {agendaItem.Title}
                </div>
                {/* Tercera fila */}
                <div className={`${styles.itemRow}`}>
                    {/* Columna de la izquierda */}
                    <div className={`${styles.itemColumn}`} style={{ flex: "1" }}>
                        {
                            /* Descripción */
                            // En el modo ordenación se oculta la descripción para facilitar la visualización de más elementos
                            (!isInReOrderMode) &&
                            <div
                                className={`${styles.itemRow} ${styles.topPaddingAdjust} ${styles.desc}`}
                                style={{ flex: "1", paddingRight: "5px" }}
                            >
                                {agendaItem.Description}
                            </div>
                        }
                        {
                            /* Documents relacionados */
                            (!isInReOrderMode && (agendaItem.RelatedDocumentsIds.length > 0 || (agendaItem.AgendaItemType === AgendaItemTypes.Informative && agendaItem.AgreementRelatedCertificateId))) &&
                            <div className={`${styles.itemRow}`}>
                                <div className={`${styles.itemColumnGrow}`}>
                                    {this.renderRelatedDocumentation(agendaItem)}
                                </div>
                            </div>
                        }
                    </div>
                    {/* Columna de la derecha */}
                    {(!locked && !isInReOrderMode) && <div
                        className={`${styles.itemColumn}`}
                        style={{ justifyContent: "flex-end", paddingBottom: "5px" }}
                    >
                        {/* Botón para añadir un nuevo subpunto */}
                        <Tooltip
                            withArrow
                            relationship="label"
                            positioning="above"
                            content={
                                (tooltipsConfig[ItemActions.Add].isDisabled && !IsNullOrECNTy(tooltipsConfig[ItemActions.Add].message)) ? tooltipsConfig[ItemActions.Add].message : strings.AddSubPoint
                            }
                        >
                            <Link
                                className={styles.linkAddSubPoint}
                                disabled={tooltipsConfig[ItemActions.Add].isDisabled}
                                onClick={(): void => this.addAgendaItemToDictionary(agendaItem.Id)}
                            >
                                <TextBulletListAddRegular />
                                {window.innerWidth > 640 && <span style={{ marginLeft: "5px", fontSize: "12px" }}>
                                    {strings.AddSubPoint}
                                </span>}
                            </Link>
                        </Tooltip>
                    </div>}
                </div>
            </div >
        );
    }

    public render(): React.ReactElement<IMeetingPointsprops> {
        const {
            event,
            context,
            spService,
            isEditor,
            readOnly
        } = this.props;
        const {
            isLoading,
            isInReOrderMode,
            editAgendaItem,
            openNewDoc,
            sectionNewDoc,
            isLoadingItemId,
            downloadingTheOrderOfTheDay,
            agendaItemsDictionary,
            selectedCertificateAdd
        } = this.state;

        const locked: boolean = (!isEditor || readOnly);
        const isDisabled: boolean = (
            isLoading ||
            (isLoadingItemId !== undefined) ||
            (editAgendaItem !== undefined) ||
            downloadingTheOrderOfTheDay
        );
        const isECNTy: boolean = !(Object.keys(agendaItemsDictionary)?.length > 0);

        return (
            <section className={styles.meetingPoints} >
                <div className={styles.pointsList}>
                    {
                        /* Se muestran todos los puntos/subpuntos del orden del día */
                        this.renderRecursiveAgendaItems()
                    }
                    {
                        /* Si el ID del AgendaItem es 0 (significa que es nuevo) se muestra una caja adicional */
                        // (editAgendaItem?.Id === 0) && this.addAgendaItemToDictionary("")
                    }
                    {
                        (isLoading && isLoadingItemId === undefined) ?
                            /* Mensaje para indicar que se está cargando el orden del día */
                            <div className={`${styles.itemColumnGrow} ${styles.justifyCenter}`}>
                                <Spinner label={`${strings.LoadingPoint}...`} />
                            </div>
                            :
                            /* Mensaje para indicar que no hay un orden del día para mostrar */
                            (isECNTy && editAgendaItem === undefined) &&
                            <div className={`${styles.itemColumnGrow} ${styles.justifyCenter} ${styles.noElementsContainer}`}>
                                {strings.ThereAreNoPointsToShow}
                            </div>
                    }
                </div>
                {
                    (!locked && !isLoading && !isInReOrderMode) &&
                    <div className={`${styles.itemRow} ${styles.bottomButtonRow}`}>
                        {
                            /* Spinner para indicar que se está descargando el orden del día */
                            (downloadingTheOrderOfTheDay) &&
                            <div className={`${styles.itemColumn} ${styles.justifyCenter}`}>
                                <Spinner size="tiny" />
                            </div>
                        }
                        {
                            /* Botón para habilitar el modo reordenar */
                            (isEditor && !readOnly) &&
                            <div className={`${styles.itemColumn}`}>
                                <Button
                                    disabled={isDisabled || isECNTy}
                                    appearance="secondary"
                                    icon={<ListRegular />}
                                    onClick={this.enableReorderMode}
                                >
                                    {strings.ReorderItems}
                                </Button>
                            </div>
                        }
                        {
                            /* Botón para descargar el orden del día */
                            (isEditor && !readOnly) &&
                            <div className={`${styles.itemColumn}`}>
                                <Button
                                    disabled={isDisabled || isECNTy}
                                    appearance='secondary'
                                    icon={<ArrowDownloadRegular />}
                                    onClick={(): promise<void> => this.downloadTheOrderOfTheDay()}
                                >
                                    {strings.DownloadTheOrderOfTheDay}
                                </Button>
                            </div>
                        }
                        {/* Botón para añadir un punto del orden del día */}
                        <div className={`${styles.itemColumn}`}>
                            <Button
                                disabled={isDisabled}
                                appearance="primary"
                                icon={<AddRegular />}
                                onClick={() => this.addAgendaItemToDictionary()}
                            >
                                {strings.AddPointOfTheOrderOfTheDay}
                            </Button>
                        </div>
                    </div>
                }
                {
                    (!locked && !isLoading && isInReOrderMode) &&
                    <div className={`${styles.itemRow} ${styles.bottomButtonRow}`}>
                        {
                            /* Mensaje informativo sobre reordenar */
                            (isEditor && !readOnly) &&
                            <span className={styles.infoText}>
                                <Info20Regular /> {strings.ReorderInfoMessage}
                            </span>
                        }
                        {
                            /* Botón para cancelar el reorden */
                            <div className={`${styles.itemColumn}`}>
                                <Button
                                    disabled={isDisabled}
                                    appearance="secondary"
                                    onClick={(): void => {
                                        this.openDialog(MultiDialogType.CancelReorder)
                                    }}
                                >
                                    {strings.RuleOut}
                                </Button>
                            </div>
                        }
                        {/* Botón para guardar el reorden */}
                        <div className={`${styles.itemColumn}`}>
                            <Button
                                disabled={isDisabled}
                                appearance="primary"
                                onClick={(): void => {
                                    this.openDialog(MultiDialogType.SaveReorder)
                                }}
                            >
                                {strings.Save}
                            </Button>
                        </div>
                    </div>
                }
                {
                    (openNewDoc && sectionNewDoc) &&
                    <NewDocForm
                        open={openNewDoc}
                        openCloseForm={this.onOpenCloseDocForm}
                        context={context}
                        spService={spService}
                        relativeUrlToGetDocs={event.StorageServerRelativeUrl}
                        relatedDocumentsIds={(sectionNewDoc.AgendaItemType === AgendaItemTypes.Informative && sectionNewDoc.AgreementRelatedCertificateId) ? sectionNewDoc.RelatedDocumentsIds.concat([sectionNewDoc.AgreementRelatedCertificateId]) :sectionNewDoc.RelatedDocumentsIds}
                        onSaveData={this.onSaveRelatedDocument}
                        relatedItemTitle={`(${sectionNewDoc?.Order}) ${sectionNewDoc?.Title}`}
                    />
                }{openNewDoc && selectedCertificateAdd &&
                    <NewDocForm
                    open={openNewDoc}
                    openCloseForm={this.onOpenCloseDocForm}
                    context={context}
                    spService={spService}
                    relativeUrlToGetDocs={event.StorageServerRelativeUrl}
                    relatedDocumentsIds={selectedCertificateAdd.RelatedDocumentsIds}
                    onSaveData={this.onSaveRelatedCertificate}
                    relatedItemTitle={`(${selectedCertificateAdd?.Order}) ${selectedCertificateAdd?.Title}`}
                />
                }
                {this.renderMultiDialog()}
            </section >
        );
    }

}