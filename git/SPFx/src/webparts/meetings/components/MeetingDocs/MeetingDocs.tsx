import * as React from 'react';
import * as strings from 'MeetingsWebPartStrings';
import styles from './MeetingDocs.module.scss';
import { IDocLabelColors, IDocumentSensitivityLabel, IMeetingDoc, IMeetingDocsprops, IMeetingDocsState, ITenantSensitivityLabel } from './IMeetingDocs';
import { File, FileList, MgtTemplateprops } from '@microsoft/mgt-react/dist/es6/spfx';
import { ViewType } from '@microsoft/mgt-spfx';
import { DriveItem } from '@microsoft/microsoft-graph-types';
import { Button, Divider, Dropdown, Label, Spinner, Option, Tooltip } from '@fluentui/react-components';
import { EventStatus, IEventStorage } from '../../../../service/BackendServiceModels/EventModels';
import { ArrowClockwiseFilled, DeleteRegular, Signature20Filled, TagMultiple20Filled } from '@fluentui/react-icons';
import ConfirmAction from '../../../../components/ConfirmAction';

const NoSensitivityLabelId = "00000000-0000-0000-0000-000000000000";

export const FileTemplate = (props: MgtTemplateprops & IMeetingDocsprops & { onDeleteClick: any, onUpdateDocs:any}) => {
    const file: IMeetingDoc | undefined = props.dataContext ? props.dataContext.file : undefined;
    //const estadoConBorrado = props.event.StatusId === EventStatus.InConstruction || props.event.StatusId === EventStatus.testBooking || props.event.StatusId === EventStatus.Published || props.event.StatusId === EventStatus.InCelebration || props.event.StatusId === EventStatus.Celebrated;

    if(file?.name?.indexOf("_hidden") == -1)props.onUpdateDocs(file);
    return (<></>)
    /* return (
        file?.name?.indexOf("_hidden") == -1 ?
            <>
                <div style={{ display: 'flex', flexDirection: 'row' }}>
                    <File view={ViewType.twolines} fileDetails={file} onClick={(props) => {
                        if (file && !file.folder) {
                            window.open(file.webUrl?.toString(), "_blank", "noreferrer")
                        }
                    }}></File>
                    {
                        props.isEditor && estadoConBorrado && !file.folder &&
                        <div style={{ flexGrow: 1, display: 'flex', alignSelf: 'center', justifyContent: 'flex-end' }}>
                            <Button title={strings.Delete} appearance='secondary' icon={<DeleteRegular />} onClick={() => { props.onDeleteClick(file) }} />
                        </div>
                    }
                </div>
                <Divider appearance='subtle'></Divider>
            </>
            :
            <></> //no pintamos el item
    );*/
}

export const Loading = (props: MgtTemplateprops) => {
    return (
        <div>
            <Spinner label={strings.Loading + "..."} labelPosition="after" size="small" />
        </div>
    );
}

export default class MeetingDocs extends React.Component<IMeetingDocsprops, IMeetingDocsState> {
    
    private _driveDocsBuffer: DriveItem[] = [];
    private _flushScheduled = false;

    constructor(props: IMeetingDocsprops) {
        super(props);
        this.state = {
            breadCrumbItems: [{ title: strings.Home, id: undefined }],
            currentDriveItemId: undefined,
            eventRootStorage: undefined,
            loadingDocuments: true,
            openNewDocForm: false,
            deleting: false,
            allDocs: [],
            backDocs:[],
            driveDocs: [],
            sensitibityLabels:[],
            updatingDoc: false
        }
    }
    
    private onFileFromTemplate = (file: DriveItem) => {
        // acumulamos sin mutar
        this._driveDocsBuffer.push(file);
        this.scheduleFlush();
    };

    
    private scheduleFlush() {
        if (this._flushScheduled) return;
        this._flushScheduled = true;
        // un microtask / macrotask para agrupar múltiples llamadas
        setTimeout(() => {
        this.flushDriveDocsBuffer();
        this._flushScheduled = false;
        }, 0);
    }

    
    private flushDriveDocsBuffer() {
        if (this._driveDocsBuffer.length === 0) return;
        const batch = this._driveDocsBuffer;
        this._driveDocsBuffer = [];

        // opcional: deduplicar por id (por si el template re-renderiza)
        const dedup = (items: DriveItem[]) => {
        const map = new Map<string, DriveItem>();
        for (const it of items) {
            if (it.id) map.set(it.id, it);
        }
        return Array.from(map.values());
        };

        this.setState(testv => ({
            driveDocs: dedup([...testv.driveDocs, ...batch]),
        }));
    }



    public async componentDidMount(): promise<void> {
        void this.onInit();
    }

    public onFileSelect(file: DriveItem | undefined): void {
        this.setState({ selectedFile: file });
    }

    public async deleteSelectedFile() {
        this.setState({ deleting: true });
        const { bkService, event } = this.props;
        const { selectedFile } = this.state;

        const uniqueId = selectedFile?.cTag?.match(/\{(.*?)\}/g)?.[0] + "";  //workarround para sacar el unique id del cTag (basicamente son lo mismo)
        await bkService.deleteDocumentByUniqueId(event.BodyId, event.Id, uniqueId);
        this.setState({ selectedFile: undefined, deleting: false, allDocs:[], driveDocs: [] });
    }

    public async updatingSelectedFile() {
        this.setState({ updatingDoc: true });
        const { bkService, event, isEditor} = this.props;
        const { fileToUpdate, backDocs } = this.state;
        if(fileToUpdate){
            const isArchived: boolean = event.StatusId === EventStatus.Archived;
            const file = backDocs.find(b => b.FilePath.split("/").pop() == fileToUpdate.name);
            if(file){
                file.ExpectedLabel = fileToUpdate.sensitivityLabelId || "";
                if(file.ExpectedLabel === NoSensitivityLabelId){
                    file.ExpectedLabel = "";
                }
                try{
                    await bkService.updateSensitivityLabel(file, isArchived, isEditor);
                }catch(ex){
                    console.log(ex);
                }
                
            }
            this.setState({ selectedFile: undefined,fileToUpdate:undefined, updatingDoc: false, deleting: false, allDocs:[], driveDocs: [],backDocs:[], loadingDocuments: true});
            const allDocumentsWithSL: IDocumentSensitivityLabel[] = await bkService.getDocumentsSensitivityLabelsByEventSharedId(event.BodyId, event.Id, isArchived, isEditor); 
            this.setState({backDocs:allDocumentsWithSL, loadingDocuments: false});
        }
        
    }

    
    componentDidUpdate(testvprops: IMeetingDocsprops, testvState: IMeetingDocsState) {
        const { driveDocs, backDocs, sensitibityLabels, loadingDocuments, allDocs } = this.state;

        // Si estamos cargando, no derive nada
        if (loadingDocuments) return;

        // Caso sin labels: allDocs = driveDocs con defaults
        const noLabels = !sensitibityLabels || sensitibityLabels.length < 1;

        // Condiciones para recalcular solo cuando cambian las entradas
        const inputsChanged =
            testvState.driveDocs !== driveDocs ||
            testvState.backDocs !== backDocs ||
            testvState.sensitibityLabels !== sensitibityLabels;

        if (!inputsChanged) return;

        let nextAllDocs: IMeetingDoc[] = [];

        if (noLabels) {
            // No hay labels: convierto driveDocs directamente
            nextAllDocs = driveDocs.map<IMeetingDoc>(item => ({
            ...item,
            // Inicializaciones mínimas
            isOtherError: false,
            }));
        } else {
            // Hay labels: combino backDocs + driveDocs
            // (sigue tu lógica actual, pero aquí, fuera del render)
            // Primero copiamos las carpetas tal cual
            nextAllDocs = (driveDocs.filter(el => el?.folder) as IMeetingDoc[]).map(folder => ({
            ...folder,
            isOtherError: false, // carpetas no deberían tener errores de extensión
            }));

            // Luego mapeamos cada backDoc a su driveDoc
            backDocs.forEach(el => {
            const name = el.FilePath.split("/").pop();
            const driveDoc = driveDocs.find(b => b.name === name);
            if (driveDoc) {
                const labelId = el.ExpectedLabel ? el.ExpectedLabel : el.SensitivityLabelId;
                const labelName = labelId
                ? sensitibityLabels.find(s => s.Id === labelId)?.Name
                : undefined;

                const finalDoc: IMeetingDoc = {
                ...driveDoc,
                sensitivityLabelId: labelId,
                sensitivityLabel: labelName,
                isSignedDocument: el.IsSignedDocument,
                isExternalLabel: el.IsExternalLabel,
                isprocessingLabel: el.IsprocessingLabel,
                isOtherError: el.IsOtherError ? el.IsOtherError : !this.isExtensionCompatible(el.FilePath),
                };
                nextAllDocs.push(finalDoc);
            }
            });
        }

        // Evita setState innecesario (comparación por longitud + ids)
        const sameLength = nextAllDocs.length === allDocs.length;
        const sameIds =
            sameLength &&
            nextAllDocs.every((d, i) => d.id === allDocs[i]?.id && d.sensitivityLabelId === allDocs[i]?.sensitivityLabelId);

        if (!sameIds) {
            this.setState({ allDocs: nextAllDocs });
        }
    }

    private async onInit(): promise<void> {
        const { bkService, event, isEditor} = this.props;

        this.setState({ loadingDocuments: true });
        try {
            const isArchived: boolean = event.StatusId === EventStatus.Archived;
            const eventStorage: IEventStorage = await bkService.getEventStorage(event.BodyId, event.Id, isArchived);
            const allSensitivityLabels: ITenantSensitivityLabel[] = await bkService.getSensitivityLabels();
            let allDocumentsWithSL: IDocumentSensitivityLabel[]=[];
            if(allSensitivityLabels && allSensitivityLabels.length> 0){
                allSensitivityLabels.unshift({Id:NoSensitivityLabelId, Name:strings.NoLabel});
                allDocumentsWithSL = await bkService.getDocumentsSensitivityLabelsByEventSharedId(event.BodyId, event.Id, isArchived, isEditor);
            }
            this.setState({ loadingDocuments: false, eventRootStorage: eventStorage, sensitibityLabels: allSensitivityLabels, backDocs:allDocumentsWithSL });
        }
        catch (error) {
            this.setState({ loadingDocuments: false });
        }
    }

    public openCloseFormFromParent() {
        this.setState({ openNewDocForm: !this.state.openNewDocForm });
    }

    public onUpdateDocumentLabel(file: IMeetingDoc, labelId:string){
        if(file.sensitivityLabelId!==labelId){
            this.setState({loadingDocuments:true, fileToUpdate: {...file, sensitivityLabelId: labelId}});
            this.setState(testvState => ({
            allDocs: testvState.allDocs.map(doc =>
                doc.id === file.id ?  {...doc, sensitivityLabelId: labelId} : doc
            )
            , loadingDocuments: false}));
        }
    }

    public onCancelLabelUpdate(){
        const {fileToUpdate, backDocs} = this.state;
        if(fileToUpdate){
            const oldLabel =  backDocs.find(b => b.FilePath.split("/").pop() == fileToUpdate.name)?.SensitivityLabelId;
            this.setState(testvState => ({
                allDocs: testvState.allDocs.map(doc =>
                    doc.id === fileToUpdate.id ?  {...doc, sensitivityLabelId: oldLabel} : doc
                )
                , fileToUpdate: undefined, updatingDoc: false}));
        }
        
    }

    public isExtensionCompatible(filePath:string):boolean{
        const allowedExtensions = [
            '.docx', '.docm',
            '.xlsx', '.xlsm', '.xlsb',
            '.pptx', '.ppsx',
            '.pdf'
        ];

        // Extraer la extensión del archivo
        const fileName = filePath.split('/').pop() ?? '';
        const extension = '.' + fileName.split('.').pop()?.toLowerCase();

        // Comprodbar si la extensión está en la lista
        return allowedExtensions.includes(extension);
    }

    
    handleFolderClick = (folder: IMeetingDoc) => {
        this.itemClick(folder)
    };

    handleOnRefresh =async ()=> {
        const { bkService, event, isEditor} = this.props;
        this.setState({ loadingDocuments: true, allDocs:[], driveDocs:[] });
        try {
            const isArchived: boolean = event.StatusId === EventStatus.Archived;
            const allDocumentsWithSL: IDocumentSensitivityLabel[] = await bkService.getDocumentsSensitivityLabelsByEventSharedId(event.BodyId, event.Id, isArchived, isEditor);
            console.log(allDocumentsWithSL);
            this.setState({ loadingDocuments: false, backDocs:allDocumentsWithSL });
        }
        catch (error) {
            this.setState({ loadingDocuments: false });
        }
    }
    

    public render(): React.ReactElement<IMeetingDocsprops> {
        const { isEditor, readOnly, event } = this.props;
        const { allDocs, breadCrumbItems, currentDriveItemId, eventRootStorage, loadingDocuments, selectedFile, deleting, sensitibityLabels, fileToUpdate, updatingDoc } = this.state;

        const driveRoot = isEditor ? eventRootStorage?.InConstruction : eventRootStorage?.Published;
        const locked = !isEditor || readOnly;
        const estadoConBorrado = event.StatusId === EventStatus.InConstruction ||event.StatusId === EventStatus.testBooking || event.StatusId === EventStatus.Published || event.StatusId === EventStatus.InCelebration || event.StatusId === EventStatus.Celebrated;
        
        return (
            <section className={styles.meetingDocs}>
                <ul className={styles.breadcrumb}>
                    {breadCrumbItems?.map((item, index) => {
                    const isLast = index === breadCrumbItems.length - 1;
                    const onClick = () => {
                        if (isLast) return; // ← no hacemos nada si es el actual
                        this.setState({ currentDriveItemId: item.id !== undefined ? item : undefined, driveDocs:[], allDocs:[] });
                        this.removeBreadcrumbItems(index + 1);
                    };
                    return (
                        <li><a
                        onClick={onClick}
                        aria-current={isLast ? 'page' : undefined}
                        style={{ pointerEvents: isLast ? 'none' : 'auto', opacity: isLast ? 0.7 : 1 }}
                        >
                        {item.title}
                        </a></li>
                    );
                    })}
                </ul>
                {loadingDocuments && 
                    <Spinner/>
                }
                {
                    !loadingDocuments && !deleting && (() => {
                        const isRoot = currentDriveItemId === undefined;
                        const driveId = isRoot ? driveRoot?.DriveId || "" : currentDriveItemId.parent?.driveId || "";
                        const itemId = isRoot ? driveRoot?.DriveItemId : currentDriveItemId.id;
                        const fileListQuery = `/me/drives/${driveId}/items/${itemId}/children`;
                        return (<>
                            <FileList
                                className={styles.fileList}
                                pageSize={100}
                                enableFileUpload={!locked}
                                disableOpenOnClick
                                //itemClick={this.itemClick.bind(this)}
                                fileListQuery={fileListQuery}
                                driveId={driveId}
                                itemId={itemId}
                            >
                                <FileTemplate
                                    template="file"
                                    {...this.props}
                                    onDeleteClick={this.onFileSelect.bind(this)}
                                    onUpdateDocs={(file: IMeetingDoc) =>
                                        this.onFileFromTemplate(file)
                                    }
                                />
                                <Loading template="loading" />
                            </FileList>
                            {!locked &&
                                <div className={styles.renewButton}>
                                    <Button onClick={ev => this.handleOnRefresh()} icon={<ArrowClockwiseFilled />} title={strings.RefreshDocuments}/>
                                </div>
                            }
                            {allDocs.map(doc => (
                                <>
                                <div key={doc.id} style={{ display: 'flex', flexDirection: 'row' }} className={styles.divFile}>
                                    <File view={ViewType.twolines} fileDetails={doc} onClick={() => {
                                        if (!doc.folder) {
                                            window.open(doc.webUrl?.toString(), "_blank", "noreferrer")
                                        }else{
                                            this.handleFolderClick(doc);
                                        }
                                    }} />
                                    {
                                        this.props.isEditor && !doc.folder && sensitibityLabels && sensitibityLabels.length>0 ?
                                            <div style={{ flexGrow: 1, display: 'flex', alignSelf: 'center', justifyContent: 'flex-end', gap:5}}>
                                                
                                                {doc.isSignedDocument || doc.isExternalLabel || doc.isOtherError ?
                                                    <>
                                                        {doc.isSignedDocument ? 
                                                            <>
                                                            <Tooltip
                                                                    withArrow
                                                                    content={strings.SignedLabel}
                                                                    relationship="label"
                                                                >
                                                                    <div className={styles.tagSection}>
                                                                        <Signature20Filled/>
                                                                    </div>
                                                            </Tooltip>    
                                                            <Tooltip
                                                                withArrow
                                                                content={strings.SignedLabel}
                                                                relationship="description"
                                                            >
                                                                <div className={styles.tagSection}>
                                                                    <TagMultiple20Filled color={IDocLabelColors.NoLabel} />
                                                                    <Label className={styles.labeltag}>{strings.NoAvailableToLabel}</Label>
                                                                </div>
                                                            </Tooltip>
                                                            </>
                                                        :
                                                            doc.isOtherError ? 
                                                                <Tooltip
                                                                        withArrow
                                                                        content={strings.DocHasOtherError}
                                                                        relationship="description"
                                                                    >
                                                                    <div className={styles.tagSection}>
                                                                        <TagMultiple20Filled color={IDocLabelColors.NoLabel} />
                                                                        <Label className={styles.labeltag}>{strings.NoAvailableToLabel}</Label>
                                                                    </div>
                                                                </Tooltip>
                                                                :
                                                            <Tooltip
                                                                    withArrow
                                                                    content={strings.ExternalLabelTooltip}
                                                                    relationship="description"
                                                                >
                                                                <div className={styles.tagSection}>
                                                                    <TagMultiple20Filled color={IDocLabelColors.External} />
                                                                    <Label className={styles.labeltag}>{strings.ExternalLabel}</Label>
                                                                </div>
                                                            </Tooltip>
                                                        }
                                                    
                                                    </>
                                                    :
                                                    <div className={styles.dropdownZoneFile}>
                                                        <Tooltip
                                                            withArrow
                                                            content= {doc.isprocessingLabel ? strings.LabelOnprocess : (doc.sensitivityLabelId && doc.sensitivityLabelId !== NoSensitivityLabelId && sensitibityLabels.some(s => s.Id === doc.sensitivityLabelId))? strings.LabelApplied: strings.NoLabelApplied}
                                                            relationship="description"
                                                        >
                                                            <TagMultiple20Filled color={doc.isprocessingLabel ? IDocLabelColors.processing : (doc.sensitivityLabelId && doc.sensitivityLabelId !== NoSensitivityLabelId && sensitibityLabels.some(s => s.Id === doc.sensitivityLabelId))? IDocLabelColors.Applied: IDocLabelColors.NoLabel} />
                                                        </Tooltip>
                                                        <Dropdown
                                                            className={styles.dropDownTag}
                                                            value={doc.sensitivityLabelId  && sensitibityLabels.some(s => s.Id === doc.sensitivityLabelId) ?  sensitibityLabels.find(s => s.Id === doc.sensitivityLabelId)?.Name: strings.NoLabel}
                                                            onOptionSelect={(ev, data) => this.onUpdateDocumentLabel(doc, data.optionValue || "")}
                                                            disabled={doc.isprocessingLabel || locked}
                                                            title={doc.isprocessingLabel ? strings.LabelOnprocess : doc.sensitivityLabelId ?  sensitibityLabels.find(s => s.Id === doc.sensitivityLabelId)?.Name: ""}
                                                        >
                                                            {sensitibityLabels.map((option: ITenantSensitivityLabel) => (
                                                                <Option key={option.Id} value={option.Id}>{option.Name}</Option>
                                                            ))}
                                                        </Dropdown>
                                                    </div>
                                                }
                                                {isEditor && estadoConBorrado && !doc.folder &&
                                                    <Button title={strings.Delete} style={{marginLeft:5}} appearance='secondary' icon={<DeleteRegular />} onClick={() => { this.onFileSelect(doc) }} />
                                                }
                                            </div>:
                                            <>
                                                {isEditor && estadoConBorrado && !doc.folder &&
                                                    <div style={{ flexGrow: 1, display: 'flex', alignSelf: 'center', justifyContent: 'flex-end', gap:5}}>
                                                        <Button title={strings.Delete} style={{marginLeft:5}} appearance='secondary' icon={<DeleteRegular />} onClick={() => { this.onFileSelect(doc) }} />
                                                    </div>
                                                }
                                            </>
                                    }
                                </div>
                                <Divider appearance='subtle' style={{flexGrow:0}}></Divider>
                                </>
                            ))}
                            </>
                        );
                    })()
                }
                {<ConfirmAction
                    dialogprops={{ open: fileToUpdate != undefined }}
                    dialogTitle={strings.UpdatingDocumentLabel}
                    dialogContent={
                    <div>
                        {strings.DocumentLabel}: <strong>{fileToUpdate?.name}</strong>
                        <br/>
                        {strings.TagLabel}: <strong>{fileToUpdate?.sensitivityLabelId ? sensitibityLabels.find(s => s.Id === fileToUpdate.sensitivityLabelId)?.Name: ""}</strong>
                        <br/>
                        <br/>
                        {strings.AreYouSure}
                    </div>
                
                    }
                    loadingActionContent={updatingDoc && <Spinner size="extra-tiny" label={"Actualizando" + "..."} />}
                    acceptButtonprops={{
                        appearance: 'primary',
                        disabled: updatingDoc,
                        title: strings.Update,
                        onClick: () => this.updatingSelectedFile()
                    }}
                    cancelButtonprops={{
                        appearance: "secondary",
                        disabled: updatingDoc,
                        title: strings.Cancel,
                        onClick:() => this.onCancelLabelUpdate()
                    }} />}
                 {<ConfirmAction
                    dialogprops={{ open: selectedFile != undefined }}
                    dialogTitle={"Eliminar Document"}
                    dialogContent={<div>¿Quieres eliminar este Document: <strong>{selectedFile?.name}</strong> ?</div>}
                    loadingActionContent={deleting && <Spinner size="extra-tiny" label={strings.Deleting + "..."} />}
                    acceptButtonprops={{
                        appearance: 'primary',
                        icon: <DeleteRegular />,
                        disabled: deleting,
                        title: strings.Delete,
                        onClick: () => this.deleteSelectedFile()
                    }}
                    cancelButtonprops={{
                        appearance: "secondary",
                        disabled: deleting,
                        title: strings.Cancel,
                        onClick: () => this.setState({ selectedFile: undefined })
                    }} />}
            </section>
        );
    }

    public removeBreadcrumbItems(index: number) {
        const { breadCrumbItems } = this.state;

        this.setState({ breadCrumbItems: breadCrumbItems.slice(0, index) });
    }

    public itemClick(e:IMeetingDoc) {
        if (e.folder) {
            let id = e.id || "";
            let name = e.name || "";
            let parentReference = e.parentReference || undefined;

            // render new file list
            const { breadCrumbItems } = this.state;

            breadCrumbItems?.push({ title: name, id, parent: parentReference });

            this.setState({ breadCrumbItems, currentDriveItemId: { title: name, id, parent: parentReference }, driveDocs:[], allDocs:[] });
        }
    }

}