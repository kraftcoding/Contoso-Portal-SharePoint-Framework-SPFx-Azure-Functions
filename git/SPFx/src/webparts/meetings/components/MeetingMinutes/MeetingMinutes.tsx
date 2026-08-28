import * as React from 'react';
import * as strings from 'MeetingsWebPartStrings';
import styles from './MeetingMinutes.module.scss';
import { IApprodvalStatusOfTheMinutes, IBackMinutes, IMeetingMinutesprops, IMeetingMinutesState, IMinutesSection, MinutesSectionsIds } from './IMeetingMinutes';
import { ArrowDownloadRegular, Calendar24Regular, CircleSmall20Regular, EditRegular, EyeRegular, Folder24Regular, Handshake24Regular, Notepad24Regular, NotepadRegular, PeopleCommunity24Regular } from '@fluentui/react-icons';

import {
    Button,
    Dialog,
    DialogActions,
    DialogBody,
    DialogContent,
    DialogSurface,
    DialogTitle,
    Divider,
    Field,
    Label,
    Popover,
    PopoverSurface,
    PopoverTrigger,
    Spinner,
    Textarea
} from '@fluentui/react-components';
import { NavLink } from 'react-router-dom';
import NewDocForm from '../../../../components/NewDocForm';
import { Logger } from '../../../../utils/Logger';
import { File, Person } from '@microsoft/mgt-react/dist/es6/spfx';
import { IFileInfo } from "@pnp/sp/files";
import { ViewType } from '@microsoft/mgt-spfx';
import { getDateFormatted } from '../../../../utils/Utils';
import { isECNTy } from '@microsoft/sp-lodash-subset';
import { EventStatus, IUserAttendance } from '../../../../service/BackendServiceModels/EventModels';

export default class MeetingMinutes extends React.Component<IMeetingMinutesprops, IMeetingMinutesState> {

    constructor(props: IMeetingMinutesprops) {
        super(props);
        this.state = {
            isLoading: false,
            selectedMinuteSection: undefined,
            selectedMinuteSectionFormErrorMessage: "",
            buttonToSaveTheSelectedMinuteSectionDisabled: false,
            openDocDialog: false,
            openAttendanceDialog: false,
            minutesSections: [],
            attendances: [],
            dialogLoading: false,
            downloadingTemplate: false,
            showApprodvalDetailModal: false,
        };
    }

    public componentDidMount(): void {
        void this.onInit();
    }

    private async onInit(): promise<void> {
        const { bkService, event, spService } = this.props;
        const isArchived: boolean = event.StatusId === EventStatus.Archived;

        this.setState({ isLoading: true });
        /* Primero se generan la Sections del Minutes */
        let minutesSections: IMinutesSection[] = this.generateMinutesSections();
        /* Luego se obtiene la información del Minutes */
        try {
            if (minutesSections && minutesSections.length > 0) {
                const minutesBack: IBackMinutes = await bkService.getMinutesInfo(event.BodyId, event.Id, isArchived);
                if (minutesBack) {
                    minutesSections = minutesSections.map((section: IMinutesSection) => {
                        return { ...section, Description: minutesBack[section.Id as keyof IBackMinutes] as string };
                    });
                    const attendances: IUserAttendance[] = minutesBack?.Attendances || []
                    let minutesFile: IFileInfo | undefined;
                    let approdvalStatus: IApprodvalStatusOfTheMinutes | undefined;
                    if (minutesBack.IdMinutes) {
                        [minutesFile, approdvalStatus] = await promise.all([
                            spService.getFileInfoById(minutesBack.IdMinutes),
                            bkService.getApprodvalStatusMinutes(event.BodyId, event.Id)
                        ]);
                    }
                    this.setState({ isLoading: false, minutesSections, minutesFile, approdvalStatus, attendances });
                    console.log("****");
                    console.log(approdvalStatus);
                }
            }
        }
        catch (error) {
            console.log(error);
            this.setState({ isLoading: false, minutesSections });
        }
    }

    private generateMinutesSections(): IMinutesSection[] {
        const { event } = this.props;
        const buildUrl: (section: string) => string = (section: string): string => `/${event.Id}/${section}`;

        return [
            {
                Id: MinutesSectionsIds.AgendaInformation,
                Title: strings.MeetingPoints,
                Description: '',
                Icon: <Calendar24Regular className={styles.minuteSectionIcon} />,
                Url: buildUrl("ordenes")
            },
            {
                Id: MinutesSectionsIds.AgreementsInformation,
                Title: strings.Agreements,
                Description: '',
                Icon: <Handshake24Regular className={styles.minuteSectionIcon} />,
                Url: buildUrl("acuerdos")
            },
            {
                Id: MinutesSectionsIds.AttendanceInformation,
                Title: strings.Assistants,
                Description: '',
                Icon: <PeopleCommunity24Regular className={styles.minuteSectionIcon} />,
                Url: buildUrl("asistentes")
            },
            {
                Id: MinutesSectionsIds.DocumentationInformation,
                Title: strings.Documents,
                Description: '',
                Icon: <Folder24Regular className={styles.minuteSectionIcon} />,
                Url: buildUrl("docs")
            },
        ];
    }

    private showTheEditingSection = (minutesSection: IMinutesSection): void => {
        this.setState({ selectedMinuteSection: { ...minutesSection } });
    };

    private handleTheTextareaOfTheSelectedMinuteSection = (textareaValue: string): void => {
        const { selectedMinuteSection } = this.state;

        if (selectedMinuteSection) {
            this.setState({
                selectedMinuteSection: {
                    ...selectedMinuteSection,
                    Description: textareaValue
                }
            });
        }
    };

    private handleTheButtonToCancelEditing = (): void => {
        this.setState({ selectedMinuteSection: undefined, selectedMinuteSectionFormErrorMessage: "", buttonToSaveTheSelectedMinuteSectionDisabled: false });
    };

    private handleTheButtonToSaveEditing = async (): promise<void> => {
        const { bkService, event } = this.props
        const { selectedMinuteSection } = this.state;
        let minutesSections: IMinutesSection[] = this.state.minutesSections;

        minutesSections = minutesSections.map((minutesSection: IMinutesSection): IMinutesSection => {
            if (minutesSection.Title === selectedMinuteSection?.Title) {
                return {
                    ...minutesSection,
                    Description: (selectedMinuteSection?.Description === undefined || selectedMinuteSection?.Description.trim() === "") ? "" : selectedMinuteSection?.Description
                };
            }
            return minutesSection;
        });

        this.setState({ buttonToSaveTheSelectedMinuteSectionDisabled: true });

        const updatedMinutes: IBackMinutes = {
            AgendaInformation: minutesSections.find(section => section.Id === MinutesSectionsIds.AgendaInformation)?.Description || '',
            AgreementsInformation: minutesSections.find(section => section.Id === MinutesSectionsIds.AgreementsInformation)?.Description || '',
            AttendanceInformation: minutesSections.find(section => section.Id === MinutesSectionsIds.AttendanceInformation)?.Description || '',
            DocumentationInformation: minutesSections.find(section => section.Id === MinutesSectionsIds.DocumentationInformation)?.Description || '',
            Attendances: []
        };

        try {
            // Llamada al back-end para actualizar las descripciones
            void bkService.updateMinutesInformation(event.BodyId, event.Id, updatedMinutes);

            // Se actualizan las minuteSections del estado para no tener que volver a traerlas del back
            this.setState({
                minutesSections,
                selectedMinuteSection: undefined,
                buttonToSaveTheSelectedMinuteSectionDisabled: false
            });
        } catch {
            this.setState({ selectedMinuteSectionFormErrorMessage: strings.BackendErrorMessage });
        }
    };

    /*
    private handleAcceptMinutesApprodval = async (): promise<void> => {
        const { bkService, event } = this.props;
        this.setState({ dialogLoading: true });

        try {
            await bkService.startTheMinuteApprodvalprocess(event.BodyId, event.Id);

            this.setState({ dialogLoading: false });
            void this.onInit();
        } catch (error) {
            this.setState({ dialogLoading: false });
            console.error(error);
        }
    }
    */

    public render(): React.ReactElement<IMeetingMinutesprops> {
        const { isEditor, context, event, spService, readOnly } = this.props;
        const {
            isLoading,
            minutesSections,
            selectedMinuteSection,
            selectedMinuteSectionFormErrorMessage,
            buttonToSaveTheSelectedMinuteSectionDisabled,
            openDocDialog,
            minutesFile,
            approdvalStatus,
            downloadingTemplate,
            openAttendanceDialog,
            showApprodvalDetailModal
        } = this.state;
        const locked: boolean = !isEditor || readOnly;
        const showApprodvalStatus: boolean = approdvalStatus ?
            (approdvalStatus.ApprodvalsAccepted.length > 0) ||
            (approdvalStatus.ApprodvalsPendingModification.length > 0) ||
            (approdvalStatus.ApprodvalsRejected.length > 0) ||
            (approdvalStatus.ApprodvalsPendings.length > 0)
            :
            false;
        return (
            <section className={styles.meetingMinutes} >
                {
                    isLoading ?
                        <Spinner style={{ flex: 1 }} label={`${strings.LoadingTheMinuteSections}...`} />
                        :
                        <>
                            {
                                (minutesSections && minutesSections.length > 0) ?
                                    <>
                                        <div className={styles.body}>
                                            {
                                                minutesSections.map((minutesSection: IMinutesSection, i: number): JSX.Element => {
                                                    return (
                                                        <div className={styles.minuteSectionContainer} key={i}>
                                                            <div className={styles.minuteSectionFirstRow}>
                                                                {minutesSection.Icon}
                                                                <div className={styles.minuteSectionNameContainer}>
                                                                    <span className={styles.minuteSectionName} title={minutesSection.Title}>{minutesSection.Title}</span>
                                                                </div>
                                                                {(minutesSection.Title !== selectedMinuteSection?.Title) &&
                                                                    <div className={styles.minuteSectionEditVisualizeButtons}>
                                                                        {!locked &&
                                                                            <Button
                                                                                appearance="secondary"
                                                                                title={strings.Edit}
                                                                                icon={<EditRegular />}
                                                                                disabled={selectedMinuteSection !== undefined}
                                                                                onClick={(): void => this.showTheEditingSection(minutesSection)}
                                                                            />
                                                                        }
                                                                        <NavLink exact to={minutesSection.Url} title={`Ir a ${minutesSection.Title}`}>
                                                                            <Button
                                                                                appearance="secondary"
                                                                                title={strings.Visualize}
                                                                                icon={<EyeRegular />}
                                                                                disabled={selectedMinuteSection !== undefined}
                                                                            />
                                                                        </NavLink>
                                                                    </div>
                                                                }
                                                            </div>
                                                            {
                                                                (minutesSection.Title !== selectedMinuteSection?.Title) ?
                                                                    (minutesSection.Description !== undefined && minutesSection.Description.trim() !== "") &&
                                                                    <div className={styles.minuteSectionSecondRow}>
                                                                        <div className={styles.minuteSectionNoteContainer}>
                                                                            <span className={styles.minuteSectionNote} title={minutesSection.Description}>{minutesSection.Description}</span>
                                                                        </div>
                                                                    </div>
                                                                    :
                                                                    <div className={styles.minuteSectionSecondRow}>
                                                                        <Field>
                                                                            <Textarea
                                                                                style={{ height: "100px" }}
                                                                                placeholder={(selectedMinuteSection.Description === undefined || selectedMinuteSection.Description.trim() === "") ? strings.WriteANote : selectedMinuteSection.Description}
                                                                                value={selectedMinuteSection?.Description}
                                                                                onChange={(e: React.ChangeEvent<HTMLTextAreaElement>): void => this.handleTheTextareaOfTheSelectedMinuteSection(e.target.value)}
                                                                            />
                                                                        </Field>
                                                                        {(selectedMinuteSectionFormErrorMessage !== "") &&
                                                                            <Field style={{ color: "red", margin: "-5px 0", paddingLeft: "3px" }}>
                                                                                <span>{selectedMinuteSectionFormErrorMessage}</span>
                                                                            </Field>
                                                                        }
                                                                        <div className={styles.minuteSectionCancelSaveButtons}>
                                                                            <Button appearance="secondary" title={strings.Cancel} onClick={(): void => this.handleTheButtonToCancelEditing()}>
                                                                                {strings.Cancel}
                                                                            </Button>
                                                                            <Button appearance="primary" title={strings.Save} onClick={(): promise<void> => this.handleTheButtonToSaveEditing()} disabled={buttonToSaveTheSelectedMinuteSectionDisabled}>
                                                                                {strings.Save}
                                                                            </Button>
                                                                        </div>
                                                                    </div>
                                                            }
                                                        </div>
                                                    );
                                                })
                                            }
                                            {
                                                minutesFile &&
                                                <div className={styles.minuteSectionContainer}>
                                                    <div className={styles.minuteSectionFirstRow}>
                                                        <Notepad24Regular className={styles.minuteSectionIcon} />
                                                        <div className={styles.minuteSectionNameContainer}>
                                                            <span className={styles.minuteSectionName} title={strings.StatusApprodvalMinutes}>{strings.StatusApprodvalMinutes}</span>
                                                        </div>
                                                    </div>
                                                    {
                                                        /* Estado de aprodbación del Minutes */
                                                        (approdvalStatus && showApprodvalStatus) &&
                                                        <>
                                                        <div className={styles.approdvalStatus}>
                                                            {
                                                                isEditor &&
                                                                <span>
                                                                    {strings.ApprodvalsAccepted}:
                                                                    <strong> {approdvalStatus.ApprodvalsAccepted?.length ?? 0} </strong>
                                                                </span>
                                                            }
                                                            {
                                                                isEditor &&
                                                                <span>
                                                                    {strings.ApprodvalsPendingModification}:
                                                                    <strong> {approdvalStatus.ApprodvalsPendingModification?.length ?? 0} </strong>
                                                                </span>
                                                            }
                                                            {/* 
                                                                isEditor &&
                                                                <span>
                                                                    {strings.ApprodvalsRejected}:
                                                                    <strong> {approdvalStatus.ApprodvalsRejected ?? 0} </strong>
                                                                </span>
                                                            */}
                                                            <span>
                                                                {strings.ApprodvalsPendings}:
                                                                <strong> {approdvalStatus.ApprodvalsPendings?.length ?? 0} </strong>
                                                            </span>
                                                            {
                                                                (!isECNTy(approdvalStatus.ApprodvalEndDate) && isEditor) &&
                                                                <div className={styles.footerMinutes}>
                                                                    <span>
                                                                        {strings.ValidationPeriod}
                                                                        <strong>
                                                                            {getDateFormatted(new Date(approdvalStatus.ApprodvalEndDate))}
                                                                        </strong>.
                                                                    </span>
                                                                </div>
                                                            }
                                                        </div>
                                                        <div>
                                                        {isEditor &&
                                                            <Popover open={showApprodvalDetailModal} withArrow>
                                                                <PopoverTrigger disableButtonEnhancement>
                                                                    <Label className={styles.popOverLink} onClick={() => this.setState({ showApprodvalDetailModal: true })}>{strings.MoreDetailMinutes}</Label>
                                                                </PopoverTrigger>
                                                                <PopoverSurface tabIndex={-1} className={styles.popOverDetail}>
                                                                  <div className={styles.popOverTitle}>
                                                                    {strings.DetailMinutesTitle}
                                                                  </div>
                                                                  <Divider />
                                                                    <div>
                                                                        {approdvalStatus.ApprodvalsPendings && approdvalStatus.ApprodvalsPendings.length > 0 &&
                                                                            <div className={styles.popOverRow}>
                                                                                <Label className={styles.popOverRowValue}>{strings.ApprodvalsPendings}</Label>
                                                                                <ul className={styles.popOverList}>
                                                                                    {approdvalStatus.ApprodvalsPendings.map(item => 
                                                                                        <li>{item.FullName}</li>
                                                                                    )}
                                                                                </ul>
                                                                            </div>
                                                                        }
                                                                        {approdvalStatus.ApprodvalsAccepted && approdvalStatus.ApprodvalsAccepted.length > 0 &&
                                                                            <div className={styles.popOverRow}>
                                                                                <Label className={styles.popOverRowValue}>{strings.Accepted}</Label>
                                                                                <ul className={styles.popOverList}>
                                                                                    {approdvalStatus.ApprodvalsAccepted.map(item => 
                                                                                        <li>{item.FullName}</li>
                                                                                    )}
                                                                                </ul>
                                                                            </div>
                                                                        }
                                                                        {approdvalStatus.ApprodvalsPendingModification && approdvalStatus.ApprodvalsPendingModification.length > 0 &&
                                                                            <div className={styles.popOverRow}>
                                                                                <Label className={styles.popOverRowValue}>{strings.ApprodvalsPendingModification}</Label>
                                                                                <ul className={styles.popOverList}>
                                                                                    {approdvalStatus.ApprodvalsPendingModification.map(item => 
                                                                                        <li>{item.FullName}</li>
                                                                                    )}
                                                                                </ul>
                                                                            </div>
                                                                        }
                                                                        <div className={styles.popOverButton}>
                                                                            <Button appearance="secondary" onClick={() => this.setState({ showApprodvalDetailModal: false })}>{strings.Accept}</Button>
                                                                        </div>
                                                                    </div>
                                                                </PopoverSurface>
                                                            </Popover>
                                                        }
                                                        </div>
                                                        <Divider style={{padding:10}} />
                                                        </>
                                                    }
                                                    <div className={styles.secondLineMinutesSection}>
                                                        <File
                                                            className={styles.file}
                                                            fileDetails={{ name: minutesFile.Name, webUrl: minutesFile.ServerRelativeUrl, lastModifiedDateTime: minutesFile.TimeLastModified }}
                                                            onClick={() => window.open(minutesFile.ServerRelativeUrl + "?web=1", '_blank')}
                                                            line2property='lastModifiedDateTime'
                                                            view={ViewType.twolines}
                                                        />
                                                        {/* 
                                                        <div className={styles.startTheMinuteApprodvalprocessButton}>
                                                            {!locked && this.renderStartTheMinuteApprodvalprocess()}
                                                        </div>
                                                        */}
                                                    </div>
                                                </div>
                                            }
                                        </div>
                                        <div className={styles.footer}>
                                            {!locked &&
                                                <>
                                                    {downloadingTemplate && <Spinner size='tiny' />}
                                                    <Button
                                                        appearance="secondary"
                                                        disabled={downloadingTemplate}
                                                        title={strings.ViewMinutes}
                                                        icon={<ArrowDownloadRegular />}
                                                        onClick={(): promise<void> => this.downloadCertification()}
                                                    >
                                                        {strings.DownloadMinutes}
                                                    </Button>
                                                    <Button
                                                        appearance="secondary"
                                                        title={strings.AssociateMinutes}
                                                        icon={<NotepadRegular />}
                                                        onClick={(): void => this.setState({ openAttendanceDialog: true })}
                                                    >
                                                        {strings.AssociateMinutes}
                                                    </Button>
                                                </>
                                            }
                                        </div>
                                    </>
                                    :
                                    <span className={styles.noMinuteSectionsAvailable}>{strings.NoMinuteSectionsAvailable}</span>
                            }
                        </>
                }
                {
                    openDocDialog &&
                    <NewDocForm
                        open={openDocDialog}
                        openCloseForm={this.onOpenCloseDocForm}
                        context={context}
                        spService={spService}
                        relativeUrlToGetDocs={event.StorageServerRelativeUrl}
                        relatedDocumentsIds={[]}
                        onSaveData={this.onSaveRelatedDocument}
                        relatedItemTitle={strings.ApprodveMinutes}
                    />
                }
                {
                    openAttendanceDialog && this.renderAttendanceDialog()

                }
            </section >
        );
    }

    private async downloadCertification(): promise<void> {
        const { event, bkService } = this.props;

        this.setState({ downloadingTemplate: true });

        try {
            const result: ArrayBuffer = await bkService.downloadMinutesTemplate(event.BodyId, event.Id);

            if (result) {
                const url = window.URL.createObjectURL(new Blob([result]));
                const enlace = document.createElement('a');
                enlace.href = url;
                enlace.setAttribute('download', `${strings.Minutes} - ${event.Title.substring(0, 50)}.docx`);
                document.body.appendChild(enlace);
                enlace.click();
                document.body.removeChild(enlace);
                window.URL.revokeObjectURL(url);
            }

            this.setState({ downloadingTemplate: false });
        } catch (error) {
            Logger.error("Error MeetingMinutes - downloadCertification", error, this.context);
        }
    }

    private onSaveRelatedDocument = async (fileUniqueId: string): promise<void> => {
        const { bkService, event } = this.props;
        try {
            console.log(fileUniqueId);
            await bkService.addRelatedDocumentIdToMinutes(event.BodyId, event.Id, fileUniqueId);
        } catch (error) {
            Logger.error("Error MeetingMinutes - onSaveRelatedDocument", error, this.context);
        } finally {
            void this.onInit();
        }
    };

    private onOpenCloseDocForm = (): void => {
        this.setState({ openDocDialog: !this.state.openDocDialog });
    }

    // private onOpenCloseAttendanceForm = (): void => {
    //     this.setState({ openDocDialog: false, openAttendanceDialog: !this.state.openAttendanceDialog });
    // }


    private renderAttendanceDialog(): JSX.Element {
        const { dialogLoading, attendances } = this.state;

        const onDismiss = () => this.setState({ openAttendanceDialog: false })
        console.log(attendances);

        return (
            <>
                {dialogLoading && <Spinner size='tiny' />}
                <Dialog open={this.state.openAttendanceDialog} onOpenChange={onDismiss}>
                    <DialogSurface>
                        <DialogBody>
                            <DialogTitle >{strings.AttendanceDialogTitle}</DialogTitle>
                            <DialogContent>
                                {attendances?.length > 0 ?
                                    <div>
                                        <Field className={styles.dialogAttendancesSection}>
                                            <span>{strings.AttendanceDialogDescription}</span>
                                        </Field>
                                        <Field
                                            className={styles.dialogAttendancesSection}
                                        >
                                            {attendances?.map((at: IUserAttendance): JSX.Element => (
                                                <Field style={{ display: "flex", alignItems: "center", paddingBottom: "2px" }}>
                                                    <CircleSmall20Regular />
                                                    <Person userId={at.UserPrincipalName} view={ViewType.oneline} />
                                                </Field>
                                            ))
                                            }
                                        </Field>
                                    </div>
                                    :
                                    <Field className={styles.dialogAttendancesSection}>
                                        <span>{strings.AttendanceDialogNoItems}</span>
                                    </Field>
                                }
                                 <Field className={styles.dialogAttendancesSection}>
                                <span>{strings.AttendanceDialogClarification}</span>
                                </Field>
                            </DialogContent>
                            <DialogActions>
                                <Button appearance='secondary' onClick={onDismiss}>{strings.Cancel}</Button>
                                <Button disabled={!(attendances?.length > 0)} appearance='primary' onClick={() => this.setState({ openDocDialog: true, openAttendanceDialog: false })}>{strings.AssociateMinutes}</Button>
                            </DialogActions>
                        </DialogBody>
                    </DialogSurface>
                </Dialog>

            </>
        );
    }

}