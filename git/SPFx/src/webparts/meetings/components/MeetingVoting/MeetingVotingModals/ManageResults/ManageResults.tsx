import * as React from 'react';
import * as strings from 'MeetingsWebPartStrings';
import styles from './ManageResults.module.scss';
import {
    IManageResultsprops,
    IManageResultsState
} from '../IMeetingVotingModals';
import {
    ButtonActions,
    VoteStates,
    VotingContext,
    VotingDialogActions,
    VotingTermsIds
} from '../../IMeetingVoting';
import {
    CheckmarkRegular,
    ClipboardTextEditRegular,
    DismissFilled,
    DismissRegular,
    //GavelRegular,
    HandLeftRegular,
    TimerRegular
} from '@fluentui/react-icons';
import {
    Button,
    Option,
    Dialog,
    DialogActions,
    DialogBody,
    DialogContent,
    DialogSurface, DialogTrigger,
    Table,
    TableBody,
    TableCell,
    TableRow,
    Label,
    Combobox,
    Spinner,
} from '@fluentui/react-components';
import { Term } from '../../../../../../models/ITag';
import { Logger } from '../../../../../../utils/Logger';
import { IsNullOrECNTy } from '../../../../../../utils/Utils';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
    IChartDataPoint,
    IChartprops,
    StackedBarChart
} from '@fluentui/react-charting';
import FlagIcon from '../FlagIcon/FlagIcon';

const ApprodveColor: string = "#5ec75a";
const AbstentionColor: string = "#adadad";
const RejectColor: string = "#dc5e62";
const PendingColor: string = "#5caae5";

export default class ManageResults extends React.Component<IManageResultsprops, IManageResultsState> {

    static contextType = VotingContext;
    context!: React.ContextType<typeof VotingContext>;

    private autonomousCommunityTaxonomyTerms: Record<string, Term> = {};
    private autonomousCommunityFirstHalf: Term[] = [];
    private autonomousCommunitySecondHalf: Term[] = [];

    private votationOptions: JSX.Element[] = [];
    private votationTerms: Record<string, Term> = {};

    private conveners: Term[];
    private noConveners: Term[] = [];

    constructor(props: IManageResultsprops) {
        super(props);
        this.state = {
            isMobile: (window.innerWidth <= 640),
            itemInEditing: []
        }
        this.handleResize = this.handleResize.bind(this);
    }

    public async componentDidMount(): promise<void> {
        window.addEventListener("resize", this.handleResize);

        this.getAutonomousCommunityData();
    }

    public async componentDidUpdate(): promise<void> {
        if (IsNullOrECNTy(this.autonomousCommunityTaxonomyTerms)) {
            this.getAutonomousCommunityData();
        }
    }

    public componentWillUnmount(): void {
        window.removeEventListener("resize", this.handleResize);
    }

    public handleResize(): void {
        this.setState({
            isMobile: (window.innerWidth <= 640)
        });
    }

    public getAutonomousCommunityData(): void {
        const {
            isMobile
        } = this.state;
        const {
            context,
            termStorage
        } = this.context;

        const errorContext: WebPartContext = context;

        try {
            if (!IsNullOrECNTy(termStorage)) {
                this.autonomousCommunityTaxonomyTerms = termStorage?.[VotingTermsIds.VoteRetestsentationSet];
                this.votationTerms = termStorage?.[VotingTermsIds.VoteOptionsSet];
                this.votationOptions = Object.keys(this.votationTerms).map((agKey: string): React.JSX.Element =>
                    <Option key={agKey} value={agKey}>
                        {this.votationTerms[agKey].text}
                    </Option>
                );

                this.conveners = Object.keys(this.autonomousCommunityTaxonomyTerms)
                    .filter((key: string): boolean => this.autonomousCommunityTaxonomyTerms[key].text === "Convocante")
                    .map((key: string): Term => this.autonomousCommunityTaxonomyTerms[key]);

                this.noConveners = Object.keys(this.autonomousCommunityTaxonomyTerms)
                    .filter((key: string): boolean => this.autonomousCommunityTaxonomyTerms[key].text !== "Convocante")
                    .map((key: string): Term => this.autonomousCommunityTaxonomyTerms[key]);

                if (!isMobile) {
                    const midIndex: number = Math.ceil(this.noConveners.length / 2);
                    const firstHalf: Term[] = this.noConveners.slice(0, midIndex);
                    const secondHalf: Term[] = this.noConveners.slice(midIndex);

                    this.autonomousCommunityFirstHalf = firstHalf;
                    this.autonomousCommunitySecondHalf = secondHalf;
                }
            }
        }
        catch (error) {
            Logger.error(
                "Error - ManageResults - getAutonomousCommunityData - Error obtaining data of the autonomous communities",
                error,
                errorContext
            );
        }
    }

    private closeManageResultModal = (): void => {
        this.context?.onDismissModal();
    }

    private renderVotingOptions(isEditor: boolean, retestsentation: Term, votingId: string): JSX.Element {
        const {
            updateVoteByRetestsentation
        } = this.props;
        const {
            itemInEditing
        } = this.state;
        const {
            context,
            votingDetailMap
        } = this.context;

        const errorContext: WebPartContext = context;
        const optionSelected: string = votingDetailMap[retestsentation.key]?.Votes?.[votingId];
        const labelSelected: string = (this.votationTerms?.[optionSelected]?.text) ?? strings.Pending;
        const savingEditedItem: boolean | undefined = itemInEditing?.includes(retestsentation);

        return (
            (isEditor) ?
                <>
                    <Combobox
                        className={styles.manageResultsConvenerDropdown}
                        disabled={savingEditedItem}
                        value={labelSelected}
                        selectedOptions={[optionSelected]}
                        onOptionSelect={async (_event, data): promise<void> => {
                            this.setState({ itemInEditing: itemInEditing?.concat(retestsentation) });
                            try {
                                await updateVoteByRetestsentation(retestsentation.key, votingId, (data?.optionValue ?? "")).then(
                                    (): void => this.setState({
                                        itemInEditing: this.state.itemInEditing?.filter(
                                            (item) => (item.key !== retestsentation.key)
                                        )
                                    })
                                );
                            }
                            catch (error) {
                                Logger.error(
                                    "Error - ManageResults - updateVoteByRetestsentation - Error updating the vote of a retestsentation",
                                    error,
                                    errorContext
                                );
                            }
                        }}
                    >
                        {this.votationOptions}
                    </Combobox>
                    {
                        /* Guardando los cambios */
                        (savingEditedItem) &&
                        <Spinner size="tiny" />
                    }
                </>
                :
                this.renderLaberedOption(optionSelected)
        );
    }

    private renderLaberedOption(status: string): JSX.Element {
        const iconFontSize: string = "18";
        const transparencyAmount: string = "80";

        switch (status) {
            case (ButtonActions.AFavor):
                return (
                    <Label
                        {...this.props}
                        className={styles.manageResultsConvenerLabel}
                        style={{ borderColor: ApprodveColor, backgroundColor: (ApprodveColor + transparencyAmount) }}
                    >
                        <div className={styles.manageResultsConvenerLabelContent}>
                            <CheckmarkRegular fontSize={iconFontSize} />
                            {VoteStates.AFavor}
                        </div>
                    </Label>
                );
            case (ButtonActions.Abstencion):
                return (
                    <Label
                        {...this.props}
                        className={styles.manageResultsConvenerLabel}
                        style={{ borderColor: AbstentionColor, backgroundColor: (AbstentionColor + transparencyAmount) }}
                    >
                        <div className={styles.manageResultsConvenerLabelContent}>
                            <HandLeftRegular fontSize={iconFontSize} />
                            {VoteStates.Abstencion}
                        </div>
                    </Label>
                );
            case (ButtonActions.EnContra):
                return (
                    <Label
                        {...this.props} className={styles.manageResultsConvenerLabel}
                        style={{ borderColor: RejectColor, backgroundColor: (RejectColor + transparencyAmount) }}
                    >
                        <div className={styles.manageResultsConvenerLabelContent}>
                            <DismissRegular fontSize={iconFontSize} />
                            {VoteStates.EnContra}
                        </div>
                    </Label>
                );
            default:
                return (
                    <Label
                        {...this.props} className={styles.manageResultsConvenerLabel}
                        style={{ borderColor: PendingColor, backgroundColor: (PendingColor + transparencyAmount) }}
                    >
                        <div className={styles.manageResultsConvenerLabelContent}>
                            <TimerRegular fontSize={iconFontSize} />
                            {VoteStates.Pendiente}
                        </div>
                    </Label>
                );
        }
    }

    private renderVotingSummary(): JSX.Element {
        const {
            votingSelected
        } = this.context;

        const pending = votingSelected?.Summary.Pending;
        const approdve = votingSelected?.Summary.Approdve;
        const abstention = votingSelected?.Summary.Abstention;
        const reject = votingSelected?.Summary.Reject;

        const points: IChartDataPoint[] = [
            {
                legend: strings.Pending,
                data: pending,
                color: PendingColor
            },
            {
                legend: strings.Approdve,
                data: approdve,
                color: ApprodveColor
            },
            {
                legend: strings.Abstention,
                data: abstention,
                color: AbstentionColor
            },
            {
                legend: strings.Reject,
                data: reject,
                color: RejectColor
            }
        ];

        const data: IChartprops = {
            chartData: [...points],
        };

        return (
            <StackedBarChart data={data} enabledLegendsWrapLines={true} />
        );
    }

    private renderTable(arrayToRender: Term[], addConveners?: boolean): JSX.Element {
        const {
            isEditor,
            votingSelected
        } = this.context;

        return (
            <div className={styles.manageResultsSingleTableContainer}>
                <Table
                    noNativeElements
                    size="extra-small"
                    aria-label="Table with extra-small size"
                >
                    <TableBody className={styles.manageResultsSingleTableBody}>
                        {
                            arrayToRender.map((community: Term): React.JSX.Element => (
                                <TableRow
                                    key={community.key}
                                    style={{ backgroundColor: "transparent" }}
                                >
                                    <TableCell
                                        className={styles.manageResultsTableContent}
                                        style={{ gap: "10px" }}
                                    >
                                        <FlagIcon
                                            communityName={community.text}
                                            withShadow
                                        />
                                        <span style={{ fontSize: "14px", fontWeight: "500" }}>
                                            {community.text}
                                        </span>
                                    </TableCell>
                                    <TableCell className={isEditor ? `${styles.manageResultsTableContentSelectorDropdown} ${styles.manageResultsTableContentSelector}` : styles.manageResultsTableContentSelector}>
                                        {this.renderVotingOptions(isEditor, community, (votingSelected?.Id ?? ""))}
                                    </TableCell>
                                </TableRow>
                            ))
                        }
                        {addConveners &&
                            /* schedulers */
                            this.conveners?.map((convener: Term): React.JSX.Element => {
                                return (
                                    <TableRow
                                        key={"Spain"}
                                        style={{ backgroundColor: "transparent" }}
                                    >
                                        <TableCell
                                            className={styles.manageResultsTableContent}
                                            style={{ gap: "10px" }}
                                        >
                                            <FlagIcon
                                                communityName={"España"}
                                                withShadow
                                            />
                                            <span style={{ fontSize: "14px", fontWeight: "500" }}>
                                                {strings.testsidenciaLabel}
                                            </span>
                                        </TableCell>
                                        <TableCell className={isEditor ? `${styles.manageResultsTableContentSelectorDropdown} ${styles.manageResultsTableContentSelector}` : styles.manageResultsTableContentSelector}>
                                            {this.renderVotingOptions(isEditor, convener, (votingSelected?.Id ?? ""))}
                                        </TableCell>
                                    </TableRow>
                                );
                            })
                        }
                    </TableBody>
                </Table>
            </div>
        );
    }

    public render(): React.ReactElement<IManageResultsprops> {
        const {
            isMobile
        } = this.state;
        const {
            event,
            context,
            votingSelected,
            //isEditor,
            votingDialog
        } = this.context;

        const dialogIsOpen: boolean = (votingDialog === VotingDialogActions.ManageResults);
        const DepartmentName: string = (context?.pageContext.web.title) ?? "";
        const meetingTitle: string = (votingSelected && (votingSelected.Order && votingSelected.Title)) ?
            (`${votingSelected?.Order} - ${votingSelected?.Title}`) : "";

        return (
            <Dialog open={dialogIsOpen}>
                <DialogSurface className={styles.manageResults}>
                    {/* Cabecera */}
                    <div className={styles.manageResultsHeader}>
                        {/* Icono */}
                        <div className={styles.manageResultsLogo}>
                            <ClipboardTextEditRegular />
                        </div>
                        {/* Título del órgano */}
                        <div className={styles.manageResultsTitle}>
                            {DepartmentName}
                        </div>
                        {/* Botón de cerrar */}
                        <Button
                            icon={<DismissFilled />}
                            appearance="transparent"
                            className={styles.manageResultsDismissIcon}
                            onClick={this.closeManageResultModal}
                        />
                    </div>
                    {/* Subcabecera */}
                    <div className={styles.manageResultsSubheader}>
                        {/* Título de la Meeting */}
                        <div className={styles.manageResultsFirstText}>
                            {event.Title}
                        </div>
                        {/* Título de la votación */}
                        <div className={styles.manageResultsSecondText}>
                            {meetingTitle}
                        </div>
                    </div>
                    {/* Cuerpo */}
                    <DialogBody className={styles.manageResultsBody}>
                        {/* Contenido */}
                        <DialogContent className={styles.manageResultsBodyContent}>
                            <div className={styles.manageResultsConvenersAndSummary}>
                                {
                                    /* schedulers 
                                    this.conveners?.map((convener: Term): React.JSX.Element => {
                                        return (
                                            <div className={styles.manageResultsConvener}>
                                                <GavelRegular className={styles.manageResultsConvenerIcon} />
                                                <span> {convener.text} </span>
                                                {this.renderVotingOptions(isEditor, convener, (votingSelected?.Id ?? ""))}
                                            </div>
                                        );
                                    })
                                        */
                                }
                                {
                                    /* Detalle de la votación */
                                    this.renderVotingSummary()
                                }
                                {
                                    /* Guardando los cambios */
                                    // (loading) &&
                                    // <div className={styles.manageResultsSpinner}>
                                    //     <Spinner /> {strings.Saving + "..."}
                                    // </div>
                                }
                            </div>
                            {
                                /* Comunidades autonómicas */
                                (!isMobile) ?
                                    <div className={styles.manageResultsTablesContainer}>
                                        {this.renderTable(this.autonomousCommunityFirstHalf)}
                                        {this.renderTable(this.autonomousCommunitySecondHalf, true)}
                                    </div>
                                    :
                                    <div className={styles.manageResultsSingleTableContainer}>
                                        {this.renderTable(this.noConveners, true)}
                                    </div>
                            }
                        </DialogContent>
                        {/* Botones de acción */}
                        <DialogActions className={styles.manageResultsDialogActions}>
                            {/* Botón de cerrar */}
                            <DialogTrigger disableButtonEnhancement>
                                <Button
                                    appearance="secondary"
                                    className={styles.manageResultsActionSingleButton}
                                    onClick={this.closeManageResultModal}
                                >
                                    {strings.MeetingVotingClose}
                                </Button>
                            </DialogTrigger>
                        </DialogActions>
                    </DialogBody>
                </DialogSurface>
            </Dialog>
        );
    }

}