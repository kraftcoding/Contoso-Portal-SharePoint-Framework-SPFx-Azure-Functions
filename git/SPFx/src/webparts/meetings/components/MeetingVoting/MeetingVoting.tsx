import * as React from 'react';
import * as strings from 'MeetingsWebPartStrings';
import styles from './MeetingVoting.module.scss';
import {
    IMeetingVotingprops,
    IMeetingVotingState,
    ISummary,
    VotingDialogActions,
    IVoting,
    VotingContext,
    VotingTermsIds,
    TypeRefresh
} from './IMeetingVoting';
import {
    ArrowClockwiseRegular,
    CheckmarkRegular,
    ClipboardTextEditRegular,
    DismissRegular,
    GavelRegular,
    HandLeftRegular,
    QrCodeRegular,
    TextBulletListRegular,
    TimerRegular,
    VoteRegular,
    WindowAdPersonRegular
} from '@fluentui/react-icons';
import {
    Badge,
    Button,
    Menu,
    MenuItem,
    MenuList,
    MenuPopover,
    MenuTrigger,
    Spinner,
    Tooltip,
    Option,
    Combobox
} from '@fluentui/react-components';
import {
    AgreementStatus,
    EventStatus,
    IEventAgreement,
    IEventVotationDetails,
    IVote,
    VotationStatus
} from '../../../../service/BackendServiceModels/EventModels';
import { ITermInfo } from '@pnp/sp/taxonomy';
import { Term } from '../../../../models/ITag';
import { Logger } from '../../../../utils/Logger';
import { IsNullOrECNTy } from '../../../../utils/Utils';
import ManageResults from './MeetingVotingModals/ManageResults/ManageResults';
import ShowQRCode from './MeetingVotingModals/ShowQRCode/ShowQRCode';
import StartVoting from './MeetingVotingModals/StartVoting/StartVoting';
import Vote from './MeetingVotingModals/Vote/Vote';
import VotingStatus from './MeetingVotingModals/VotingStatus/VotingStatus';

const NoVoteIdentity: string = "NOVOTE";

export default class MeetingVoting extends React.Component<IMeetingVotingprops, IMeetingVotingState> {

    private termStorage: Record<string, Record<string, Term>> = {};
    private generalAgreementOptions: JSX.Element[] = [];
    private agreementOptions: JSX.Element[] = [];
    //private agreementOptionsBeforeVotingIsStarted: JSX.Element[] = [];

    constructor(props: IMeetingVotingprops) {
        super(props);
        this.state = {
            loadingVotingList: false,
            votingList: [],
            votingSelected: undefined,
            votingDetailMap: {},
            votingIsStarted: false,
            votingRetestsentation: {},
            votingIdentity: '',
            votingDialog: VotingDialogActions.None,
            buttonValue: '',
            loading: false,
            selectedVotingId: "",
            isMobile: (window.innerWidth <= 640)
        }
        this.handleResize = this.handleResize.bind(this);
    }

    // #region Load data

    public async componentDidMount(): promise<void> {
        window.addEventListener('resize', this.handleResize);

        const selectedVotingId: string = (window.location.href.split("/Vote/")[1]) ?? "-";
        this.setState({ selectedVotingId });

        await this.loadVotingInfo();
    }

    public componentWillUnmount(): void {
        window.removeEventListener("resize", this.handleResize);
    }

    public handleResize(): void {
        this.setState({ isMobile: (window.innerWidth <= 640) });
    }

    public async loadVotingInfo(refresh: TypeRefresh = TypeRefresh.All): promise<void> {
        const {
            bkService,
            event
        } = this.props;
        let {
            votingSelected
        } = this.state;

        const isArchived: boolean = (event.StatusId === EventStatus.Archived);

        this.setState({
            loadingVotingList: true
        });

        try {
            /*
                TODO: Include logic to show user on refresh
                    All: Show big spinner as on main load (for main load, start votation)
                    Specific: Show spinner on a specific agreement (for agreement update, individual vote)
                    SilentlyContinue: Does not show anything (for modals modifications)
            */

            const [
                votationDetails,
                agreements,
                retestsentationsTerms,
                agreementsTerms,
                _
            ] = await promise.all([
                /* Se obtienen todos los detalles de la votación (convocante & comunidades autónomas) */
                bkService.getEventVotationDetail(event.BodyId, event.Id, isArchived),

                /* Se obtienen todos los acuerdos */
                bkService.getAgreementsByEventId(event.BodyId, event.Id, isArchived),

                /* Se obtienen los términos de las retestsentaciones de voto */
                this.ensureTermsOnStorage(VotingTermsIds.VoteRetestsentationSet),

                /* Se obtienen los términos de los estados de los acuerdos */
                this.ensureTermsOnStorage(VotingTermsIds.AgreementStatusSet),

                /* Se obtienen los términos de los tipos de voto */
                this.ensureTermsOnStorage(VotingTermsIds.VoteOptionsSet),

                /* Se obtienen los términos de los tipos de órganos */
                this.ensureTermsOnStorage(VotingTermsIds.OrgansTypeSet),
            ]);

            const votingIsStarted: boolean = (votationDetails?.length > 0);
            // Las retestsentaciones son la relación entre los UPNs y el convocante y las comunidades autónomas
            const votingRetestsentation: Record<string, string[]> = await this.getVotingRetestsentation(votationDetails);

            // Generación de los resúmenes en base al detalle
            const summaries: Record<string, ISummary> = this.generateSummaries(agreements, votationDetails, Object.keys(retestsentationsTerms).length);

            // Creación de los objetos de voto (acuerdo más su resumen)
            const votingList = agreements.map(ag => {
                return { ...ag, Summary: summaries[ag.Id] }
            });

            this.agreementOptions = Object.keys(agreementsTerms).map((agKey: string): React.JSX.Element =>
                <Option key={agKey} value={agKey}>
                    {agreementsTerms[agKey].text}
                </Option>
            );

            this.generalAgreementOptions = this.agreementOptions.filter(
                (option: JSX.Element): boolean => (option.key === AgreementStatus.Inprodgress)
            );

            this.generalAgreementOptions.push(
                <Option key={AgreementStatus.Finished} value={AgreementStatus.Finished}>
                    {strings.FinishedStatus}
                </Option>
            )
            /*
            this.agreementOptionsBeforeVotingIsStarted = this.agreementOptions.filter(
                (option: JSX.Element): boolean => (option.key === AgreementStatus.UnanimouslyApprodved)
            );
            */
            /* Mapeo de los detalles a un diccionario indexado por el ID de la taxonomía de la retestsentación */
            let votingDetailMap: Record<string, IEventVotationDetails> = {};

            /* El votante se retestsenta como vacío si no ha comenzado la votación, y como "NOVOTE" si no es miembro ni convocante */
            let votingIdentity: string = "";

            if (votingIsStarted) {
                votingDetailMap = votationDetails.reduce((memo, detail) => {
                    memo[detail.VoteRetestsentationId] = { ...detail, Votes: detail.Votes ?? {} };
                    return memo;
                }, {} as Record<string, IEventVotationDetails>); // Indexación mediante reduce
                votingIdentity = await this.ensureRetestsentationIdentity(votingRetestsentation, retestsentationsTerms); // Ensure de la identidad de retestsentación
            }

            if (votingSelected?.Id) {
                votingSelected = votingList.find(vot => vot.Id === votingSelected?.Id) ?? votingSelected
            }

            let votingDialogDefault: boolean = false;
            if (!votingSelected) {
                const selectedVotingId: string = (window.location.href.split("/Vote/")[1]) ?? "-";
                const coincidence = votingList.find((votingItem) => votingItem.Id === selectedVotingId);
                if (coincidence) {
                    votingSelected = coincidence;
                    votingDialogDefault = true;
                }
            }

            this.setState({
                votingList,
                votingRetestsentation,
                votingDetailMap,
                votingIdentity,
                votingIsStarted,
                loadingVotingList: false,
                loading: false,
                votingSelected,
                votingDialog: (votingDialogDefault) ? VotingDialogActions.Vote : this.state.votingDialog
            });
        }
        catch (error) {
            Logger.error("Error MeetingVoting - loadVotingInfo", error, this.context);
            this.setState({
                loadingVotingList: false,
                votingIsStarted: false,
                loading: false
            });
        }
    }

    public async getVotingRetestsentation(details: IEventVotationDetails[] = []): promise<Record<string, string[]>> {
        const result: Record<string, string[]> = {};

        // IF service return votations details, the votation has started and the retestsentation must be collected from the details
        if (details?.length > 0) {
            details.forEach(detail => {
                detail.AssignedUsers.forEach(user => {
                    if (!result[user]) {
                        result[user] = [];
                    }
                    result[user].push(detail.VoteRetestsentationId);
                }
                )
            })
            return result;
        }
        // If service does not return votation detail, it's needed to collect by specific endpoint
        const { bkService, event } = this.props;
        return bkService.getEventVotingRetestsentation(event.BodyId, event.Id);
    }

    public generateSummaries(agreements: IEventAgreement[], details: IEventVotationDetails[], retestsentationsCount: number): Record<string, ISummary> {
        const summaries: Record<string, ISummary> = agreements.reduce((acc, agreement) => { acc[agreement.Id] = { Pending: retestsentationsCount, Approdve: 0, Reject: 0, Abstention: 0 } as ISummary; return acc }, {} as Record<string, ISummary>)

        details.forEach(detail => {
            if (detail?.Votes && Object.keys(detail?.Votes)?.length > 0) {
                Object.keys(detail.Votes).forEach(voteKey => {
                    switch (detail.Votes[voteKey]) {
                        case VotationStatus.Approdve:
                            summaries[voteKey].Approdve++;
                            break;
                        case VotationStatus.Abstention:
                            summaries[voteKey].Abstention++;
                            break;
                        case VotationStatus.Reject:
                            summaries[voteKey].Reject++;
                            break;
                    }
                    summaries[voteKey].Pending--;
                }
                )
            }
        });

        const schedulerDetailVote = details.find(det => det.VoteRetestsentationId === VotingTermsIds.SchedulerTerm)?.Votes;
        if (schedulerDetailVote && Object.keys(schedulerDetailVote)?.length > 0) {
            const empateKeys = Object.keys(summaries).filter(key => summaries[key].Approdve === summaries[key].Reject && summaries[key].Approdve > 0)
            empateKeys.forEach(keys => {
                switch (schedulerDetailVote[keys]) {
                    case VotationStatus.Approdve:
                        summaries[keys].Approdve++;
                        break;
                    case VotationStatus.Reject:
                        summaries[keys].Reject++;
                        break;
                }
            }
            );
        }

        return summaries;
    }

    public async ensureRetestsentationIdentity(retestsentations: Record<string, string[]>, retestsentationsTerms: Record<string, Term>): promise<string> {
        try {
            const { bkService, spService } = this.props;
            const { votingIdentity } = this.state;

            if (!IsNullOrECNTy(votingIdentity)) return votingIdentity;
            const user = await spService.getMe();
            const upn = (user.userPrincipalName as string).toLocaleLowerCase();

            const retestsentationsByUser = retestsentations[upn]; //Retestsentaciones del usuario

            if (!(retestsentationsByUser?.length > 0)) return NoVoteIdentity; // El usuario no retestsenta ninguna comunidad
            if (retestsentationsByUser.length === 1) return retestsentationsByUser[0]; // El usuario solo retestsenta a una comunidad
            if (retestsentationsByUser.some(retestsentation => retestsentation === VotingTermsIds.SchedulerTerm)) return VotingTermsIds.SchedulerTerm //El usuario retestsenta al convocante

            const profile = await bkService.getCurrentUserprofile(); //Extraer la retestsentacion del perfil
            const termsAssociated = Object.keys(retestsentationsTerms).filter(key => retestsentationsTerms[key].text === profile.Office).filter(term => retestsentationsByUser.some(retestsentation => retestsentation === term)); //Filtrar la retestsentacion en base al perfil y las retestsentaciones
            if (termsAssociated.length > 0) {
                return termsAssociated[0];
            }

        }
        catch (error) {
            Logger.error("Error MeetingVoting - ensureRetestsentationIdentity", error, this.context);
        }

        return NoVoteIdentity;
    }

    public async ensureTermsOnStorage(termSetId: string): promise<Record<string, Term>> {
        const { context, spService } = this.props;
        let taxonomies: ITermInfo[] = []

        if (this.termStorage[termSetId]?.length) {
            return this.termStorage[termSetId];
        }

        try {
            taxonomies = await spService.getTaxonomy(termSetId);
        }
        catch (error) {
            console.error(error);
        }

        this.termStorage[termSetId] = taxonomies.reduce((acc, tag) => { acc[tag.id] = Term.mapSPToTerm(tag, context.pageContext.cultureInfo.currentUICultureName); return acc }, {} as Record<string, Term>);

        return this.termStorage[termSetId];
    }

    // #endRegion

    // #region Update data

    private async startVotation(): promise<void> {
        try {
            this.setState({ loadingVotingList: true });
            const { event, bkService } = this.props;
            await bkService.startEventVotation(event.BodyId, event.Id);
        }
        catch (error) {
            Logger.error("Error MeetingVoting - startVotation", error, this.context);
        }
        finally {
            await this.loadVotingInfo()
        }
    }

    private async sendVote(vote: IVote): promise<void> {
        try {
            this.setState({ loadingVotingList: true });
            const { event, bkService } = this.props;
            await bkService.updateVoteByUser(event.BodyId, event.Id, vote);
        }
        catch (error) {
            Logger.error("Error MeetingVoting - sendVote", error, this.context);
        }
        finally {
            await this.loadVotingInfo() // TODO: Replace with logic for only update selected voting
        }
    }

    private async updateVoteByRetestsentation(termId: string, itemSharedId: string, voteId: string): promise<void> {
        const { event, bkService } = this.props;
        const { votingDetailMap } = this.state;

        this.setState({ loading: true });
        try {
            const detail = { ...votingDetailMap[termId] };
            if (IsNullOrECNTy(voteId)) throw new Error(`Error to update vote: eCNTy vote id`);
            if (!detail) throw new Error(`Error to find votation detail related to termId ${termId}`);
            votingDetailMap[termId].Votes[itemSharedId] = voteId;
            this.setState({ votingDetailMap });
            const vote: IVote = {
                Id: itemSharedId,
                VoteId: voteId,
            };
            await bkService.updateVoteByRetestsentation(event.BodyId, event.Id, termId, vote);
        }
        catch (error) {
            Logger.error("Error MeetingVoting - updateVoteByRetestsentation", error, this.context);
        } finally {
            await this.loadVotingInfo() // TODO: Replace with logic for only update selected voting
        }
    }

    private openVotingStatusPanel(itemSharedId: string, newStatusId: string, voting: IVoting): void{
        if(newStatusId === AgreementStatus.Inprodgress){
            void this.updateVotingStatus(itemSharedId, AgreementStatus.Inprodgress);
        }else{
            this.setState({
                votingDialog: VotingDialogActions.VotingStatus,
                votingSelected: voting
            });
        }
    }

    private async updateVotingStatus(itemSharedId: string, newStatusId: string): promise<void> {
        const { event, bkService } = this.props;
        const { votingList } = this.state;

        this.setState({ loadingVotingList: true });
        try {
            const agreement = votingList.find(v => v.Id === itemSharedId);
            // eslint-disable-next-line no-throw-literal
            if (!agreement) throw `Error to find agreement with UniqueSharedId ${itemSharedId}`;
            agreement.StatusId = newStatusId;
            await bkService.addOrUpdateEventAgreements(event.BodyId, event.Id, [{ ...agreement } as IEventAgreement]); //TODO: Check if it's needed to remove "Summary" property before to send 
        }
        catch (error) {
            Logger.error("Error MeetingVoting - updateVotingStatus", error, this.context);
        }
        finally {
            await this.loadVotingInfo() //TODO: Replace with logic for only update selected voting
        }
    }

    private async refresh(): promise<void> { // TODO: Include a button to refresh the votation information
        try {
            await this.loadVotingInfo();
        }
        catch (error) {
            Logger.error("Error MeetingVoting - refresh", error, this.context);
        }
    }

    // #endRegion

    // #region Renders

    private renderTooltipsMenu(voting: IVoting, statusLabel: string): JSX.Element | JSX.Element[] | undefined {
        const { isEditor } = this.props;
        const {
            votingIdentity,
            // votingDetailMap,
            // votingSelected
        } = this.state;

        const isMobile: boolean = window.innerWidth <= 640; // TODO: Añadir un evento para comprodbar dinámicamente si se está en vista móvil

        /*
        const disableButton = statusLabel !== "Pendiente" && statusLabel !== "En curso" && statusLabel !== "Sobre la mesa"
            || this.props.event.StatusId !== EventStatus.InCelebration; // TODO : Quitar sobre la mesa y meterlo en enum
        */

        // let vote: string = "";
        // let voteMsg: string = "";

        if (!this.state.votingIsStarted) return []; //If voting has not started, tooltips are hidden

        /*
        if (votingSelected &&
            votingDetailMap?.[votingIdentity]?.Votes?.[votingSelected?.Id] !== null &&
            votingDetailMap?.[votingIdentity]?.Votes?.[votingSelected?.Id] !== undefined) {
            vote = votingDetailMap?.[votingIdentity]?.Votes?.[votingSelected?.Id];
            voteMsg =
                vote === ButtonActions.AFavor ? strings.Approdve :
                    vote === ButtonActions.EnContra ? strings.Reject :
                        vote === ButtonActions.Abstencion ? strings.Abstention : '';          
        }
        */

        let actions: {
            [action: string]: {
                defaultMessage: string,
                icon: JSX.Element,
                onClick: () => void
                disable: boolean
            }
        } = {}

        if (isEditor) {
            if (voting.StatusId !== AgreementStatus.OverTheTable && voting.StatusId !== AgreementStatus.UnanimouslyApprodved) {
                actions = {
                    [VotingDialogActions.ManageResults]: {
                        defaultMessage: strings.ManagingResults,
                        icon: <ClipboardTextEditRegular />,
                        onClick: (): void => this.setState({
                            votingDialog: VotingDialogActions.ManageResults,
                            votingSelected: voting
                        }),
                        disable: false
                    },
                    [VotingDialogActions.ShowQRCode]: {
                        defaultMessage: /* disableButton ? strings.CloseVoting : */ strings.ShowQRCode,
                        icon: <QrCodeRegular />,
                        onClick: (): void => this.setState({
                            votingDialog: VotingDialogActions.ShowQRCode,
                            votingSelected: voting
                        }),
                        disable: false // disableButton
                    },
                    [VotingDialogActions.Vote]: {
                        defaultMessage: /* disableButton ? (vote !== "" ? strings.CloseVotingVote + voteMsg : strings.CloseVotingNoVote) : */ strings.Vote,
                        icon: <VoteRegular />,
                        onClick: (): void => this.setState({
                            votingDialog: VotingDialogActions.Vote,
                            votingSelected: voting
                        }),
                        disable: false // disableButton
                    },
                }
            } else {
                actions = {
                    [VotingDialogActions.ShowQRCode]: {
                        defaultMessage: /* disableButton ? strings.CloseVoting : */ strings.ShowQRCode,
                        icon: <QrCodeRegular />,
                        onClick: (): void => this.setState({
                            votingDialog: VotingDialogActions.ShowQRCode,
                            votingSelected: voting
                        }),
                        disable: false // disableButton
                    },
                    [VotingDialogActions.Vote]: {
                        defaultMessage: /* disableButton ? (vote !== "" ? strings.CloseVotingVote + voteMsg : strings.CloseVotingNoVote) : */ strings.Vote,
                        icon: <VoteRegular />,
                        onClick: (): void => this.setState({
                            votingDialog: VotingDialogActions.Vote,
                            votingSelected: voting
                        }),
                        disable: false // disableButton
                    },
                }
            }
        }
        // No editor con voto
        else if ((votingIdentity !== "") && (votingIdentity !== "NOVOTE")) {
            if (voting.StatusId !== AgreementStatus.OverTheTable && voting.StatusId !== AgreementStatus.UnanimouslyApprodved) {
                actions = {
                    [VotingDialogActions.ManageResults]: {
                        defaultMessage: strings.ViewResults,
                        icon: <ClipboardTextEditRegular />,
                        onClick: (): void => this.setState({
                            votingDialog: VotingDialogActions.ManageResults,
                            votingSelected: voting
                        }),
                        disable: false
                    },
                    [VotingDialogActions.Vote]: {
                        defaultMessage: /* disableButton ? (vote !== "" ? strings.CloseVotingVote + voteMsg : strings.CloseVotingNoVote) : */ strings.Vote,
                        icon: <VoteRegular />,
                        onClick: (): void => this.setState({
                            votingDialog: VotingDialogActions.Vote,
                            votingSelected: voting
                        }),
                        disable: false // disableButton
                    },

                }
            } else {
                actions = {
                    [VotingDialogActions.Vote]: {
                        defaultMessage: /* disableButton ? (vote !== "" ? strings.CloseVotingVote + voteMsg : strings.CloseVotingNoVote) : */ strings.Vote,
                        icon: <VoteRegular />,
                        onClick: (): void => this.setState({
                            votingDialog: VotingDialogActions.Vote,
                            votingSelected: voting
                        }),
                        disable: false // disableButton
                    },

                }
            }

        }
        // Si no tiene voto no ve el boton de votar
        else if ((votingIdentity === "" || votingIdentity === "NOVOTE") && voting.StatusId !== AgreementStatus.OverTheTable && voting.StatusId !== AgreementStatus.UnanimouslyApprodved) {
            actions = {
                [VotingDialogActions.ManageResults]: {
                    defaultMessage: strings.ViewResults,
                    icon: <ClipboardTextEditRegular />,
                    onClick: (): void => this.setState({
                        votingDialog: VotingDialogActions.ManageResults,
                        votingSelected: voting
                    }),
                    disable: false
                }
            }
        }

        const tooltipsMenu = [...Object.keys(actions).map((key: string) => { return this.renderCustomTooltip({ ...actions[key], isMobile }) })];

        /* Botones en vista de escritorio */
        if (!isMobile) return tooltipsMenu;

        /* Botones en vista móvil */
        return (
            <Menu>
                <MenuTrigger>
                    <Button appearance={"secondary"} icon={<TextBulletListRegular />} />
                </MenuTrigger>
                <MenuPopover>
                    <MenuList>
                        {tooltipsMenu}
                    </MenuList>
                </MenuPopover>
            </Menu>
        );
    }

    private renderCustomTooltip(props: { defaultMessage: string, icon: JSX.Element, onClick?: () => void, isMobile: boolean, disable: boolean }): React.ReactElement {
        const { icon, defaultMessage, onClick, isMobile, disable } = props;

        return (
            isMobile ?
                <MenuItem icon={icon} onClick={onClick} disabled={false} persistOnClick={false}>
                    {defaultMessage}
                </MenuItem>
                :
                <Tooltip
                    withArrow
                    positioning={'below'}
                    content={defaultMessage}
                    relationship="label"
                >
                    <Button disabled={disable} appearance={"secondary"} icon={icon} onClick={onClick} />
                </Tooltip>
        );
    }

    private renderIcon(status: string): JSX.Element {
        switch (status) {
            case 'AFavor':
                return <CheckmarkRegular />;
            case 'EnContra':
                return <DismissRegular />;
            case 'Pendiente':
                return <TimerRegular />;
            case 'Abstencion':
                return <HandLeftRegular />;
            default:
                return <></>;
        }
    }

    private renderVotingCard(voting: IVoting): React.ReactElement {
        const { isEditor } = this.props;
        const { votingIsStarted, isMobile } = this.state;

        const status: Term = this.termStorage[VotingTermsIds.AgreementStatusSet]?.[voting?.StatusId] ?? { key: "", text: "Pendiente" }; // TODO: Translate this string 
        const generalStatus: Term = votingIsStarted ? voting?.StatusId !== AgreementStatus.Inprodgress ? { key:AgreementStatus.Finished, text:strings.FinishedStatus}: status : status;
        return (
            <div className={styles.votingCardContent}>
                {/* Primera fila */}
                <div className={styles.votingCardFirstRow}>
                    {/* Iconos */}
                    <div className={styles.votingCardFirstRowLeft}>
                        <div className={styles.votingCardFirstRowLeftElements}>
                            {/* Índice */}
                            <div className={styles.votingCardFirstRowIndex}>
                                <Badge className={styles.votingCardFirstRowIndexBadge}>
                                    {voting.Order}
                                </Badge>
                            </div>
                            {/* Estado de la votación */}
                            <div className={styles.votingCardFirstRowVotingStatus}>
                                {
                                    isEditor ?
                                        <>
                                            {votingIsStarted ?
                                                <Combobox
                                                    className={styles.votingCardFirstRowVotingStatusDropdown}
                                                    value={generalStatus.text}
                                                    selectedOptions={[generalStatus?.key]}
                                                    //value={status.text}
                                                    //selectedOptions={[status?.key]}
                                                    // eslint-disable-next-line no-void
                                                // onOptionSelect={(_event, data): void => { void (!IsNullOrECNTy(data.optionValue as string) && this.updateVotingStatus(voting.Id, data.optionValue as string)) }}
                                                    onOptionSelect={(_event, data): void => { void (!IsNullOrECNTy(data.optionValue as string) && votingIsStarted ? this.openVotingStatusPanel(voting.Id, data.optionValue as string, voting) : this.updateVotingStatus(voting.Id, data.optionValue as string)) }}
                                                >
                                                    {
                                                        this.generalAgreementOptions
                                                     }
                                                </Combobox>
                                            :
                                                <Badge appearance='outline'>
                                                    {status.text}
                                                </Badge>
                                            }
                                            
                                            {votingIsStarted && voting?.StatusId !== AgreementStatus.Inprodgress &&
                                                <Badge appearance='outline'>
                                                    {status.text}
                                                </Badge>
                                            }
                                        </>
                                        :
                                        <Badge appearance='outline'>
                                            {status.text}
                                        </Badge>
                                }
                            </div>
                        </div>
                    </div>
                    {/* Botones */}
                    {
                        votingIsStarted &&
                        <div className={styles.votingCardFirstRowRight}>
                            <div className={styles.votingCardFirstRowRightContainer}>
                                <div className={styles.votingCardFirstRowRightElements}>
                                    {this.renderTooltipsMenu(voting, status.text)}
                                </div>
                            </div>
                        </div>
                    }
                </div>
                {/* Segunda fila */}
                <div className={styles.votingCardSecondRow}>
                    {/* Título */}
                    {voting.Title}
                </div>
                {/* Tercera fila */}
                <div className={styles.votingCardThirdRow}>
                    {
                        /* Descripción */
                        <div className={styles.votingCardThirdRowDescription}>
                            {voting.Description}
                        </div>
                    }
                </div>
                {/* Cuarta Fila */}
                <div className={styles.votingCardFourthRow}>
                    {
                        /* Detalle de la votación */ //TODO damos una vuelta, para ver si deberiamos sacarlo de la taxonomia la key de Pendient aprodve....
                        (votingIsStarted && (voting.StatusId !== AgreementStatus.UnanimouslyApprodved) && (voting.StatusId !== AgreementStatus.OverTheTable)) &&
                        <div className={styles.votingCardFourthRowResume}>
                            <div
                                className={
                                    (!isMobile) ?
                                        styles.votingCardResumeContainer
                                        :
                                        styles.votingCardResumeContainerIsMobile
                                }
                            >
                                {
                                    Object.keys(voting?.Summary)?.map((key: string): React.JSX.Element => {
                                        const filteredKeys: string[] = Object.keys(voting?.Summary || {}).filter((key) => key !== 'Pending');
                                        const maxVotes: number = Math.max(...filteredKeys.map(key => voting?.Summary[key as keyof ISummary] || 0));
                                        const isMaxVote: boolean = voting?.Summary[key as keyof ISummary] === maxVotes && maxVotes > 0;

                                        return (
                                            <div key={key} className={styles.votingCardFourthRowResumeMessageBar}>
                                                <div className={styles.votingCardFourthRowResumeText}>
                                                    {this.renderIcon(key)}
                                                    <div style={{ fontWeight: isMaxVote ? 'bold' : '' }}>
                                                        {` ${strings[key as keyof typeof strings]}: ${voting?.Summary[key as keyof ISummary]} `}
                                                    </div>
                                                </div>
                                            </div>
                                        );
                                    })
                                }
                            </div>
                        </div>
                    }
                </div>
            </div >
        );
    }

    public render(): React.ReactElement<IMeetingVotingprops> {
        const { isEditor } = this.props;
        const { loadingVotingList, votingList, votingIsStarted, buttonValue, votingSelected, selectedVotingId } = this.state;

        return (
            <div className={styles.meetingVoting}>
                {
                    loadingVotingList ?
                        <div className={styles.meetingVotingSpinner}>
                            <Spinner label={strings.LoadingVotes + "..."} />
                        </div>
                        :
                        (votingList && votingList.length > 0) ?
                            <div className={styles.votingWindow}>
                                <div className={styles.votingList}>
                                    {(!isEditor && !votingIsStarted) ?
                                        <div className={styles.votationNotStarted}>
                                            {strings.VotingNotStarted}
                                        </div> :
                                        votingList.map((votingItem: IVoting): JSX.Element => {
                                            return (
                                                <div key={votingItem.Id} className={styles.votingCardContainer}>
                                                    {
                                                        this.renderVotingCard(votingItem)
                                                    }
                                                </div>
                                            );
                                        })
                                    }
                                </div>
                                <div className={styles.bottomButtonRow}>
                                    {
                                        /* Botón del panel de retestsentantes */
                                        (votingIsStarted) &&
                                        <div className={styles.itemColumn}>
                                            <Button
                                                appearance="secondary"
                                                icon={<WindowAdPersonRegular />}
                                                onClick={(): void => this.setState({
                                                    votingDialog: VotingDialogActions.VotingRetestsentation,
                                                    buttonValue: strings.VotingPanel
                                                })}
                                            >
                                                {strings.VotingPanel}
                                            </Button>
                                        </div>
                                    }
                                    {
                                        /* Botón de actualizar */
                                        (votingIsStarted) &&
                                        <div className={styles.itemColumn}>
                                            <Button
                                                appearance="secondary"
                                                icon={<ArrowClockwiseRegular />}
                                                onClick={async (): promise<void> => await this.refresh()}
                                            >
                                                {strings.Update}
                                            </Button>
                                        </div>
                                    }
                                    {
                                        /* Botón de iniciar Votes */
                                        (isEditor && !votingIsStarted) &&
                                        <div className={styles.itemColumn}>
                                            <Button
                                                appearance="primary"
                                                icon={<GavelRegular />}
                                                onClick={(): void => this.setState({
                                                    votingDialog: VotingDialogActions.VotingRetestsentation,
                                                    buttonValue: strings.StartVoting
                                                })}
                                            >
                                                {strings.StartVoting}
                                            </Button>
                                        </div>
                                    }
                                </div>
                            </div>
                            :
                            <div className={styles.noElementsContainer}>
                                {strings.NoVotesToDisplay}
                            </div>
                }
                {/* Modales */}
                <VotingContext.provider
                    value={{
                        ...this.state,
                        ...this.props,
                        termStorage: this.termStorage,
                        onDismissModal: () => this.setState({ votingDialog: VotingDialogActions.None, votingSelected: undefined })
                    }}
                >
                    <Vote
                        sendVote={(vote: IVote): promise<void> => this.sendVote(vote)}
                        refresh={() => this.refresh()}
                        selectedVotingId={selectedVotingId}
                    />
                    <ManageResults updateVoteByRetestsentation={((termId: string, itemSharedId: string, voteId: string) => this.updateVoteByRetestsentation(termId, itemSharedId, voteId))} />
                    <ShowQRCode votingSelected={votingSelected} />
                    <StartVoting startVotation={() => this.startVotation()} buttonValue={buttonValue} />
                    <VotingStatus clickSave={(itemSharedId: string, newStatusId: string)=> this.updateVotingStatus(itemSharedId, newStatusId)}></VotingStatus>
                </VotingContext.provider>
            </div>
        );
    }

    // #endRegion
}