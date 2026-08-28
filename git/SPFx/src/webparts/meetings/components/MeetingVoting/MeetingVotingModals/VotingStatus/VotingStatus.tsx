import * as React from 'react';
import * as strings from 'MeetingsWebPartStrings';
import styles from './VotingStatus.module.scss';
import {
    IVotingStatusState,
    IVotingStatusprops
} from '../IMeetingVotingModals';
import {
    VotingContext,
    VotingDialogActions,
    VotingTermsIds
} from '../../IMeetingVoting';
import {
    DismissFilled,
    WindowAdPersonRegular
} from '@fluentui/react-icons';
import {
    Button,
    Dialog,
    DialogActions,
    DialogContent,
    DialogSurface,
    DialogTrigger,
    Radio,
    RadioGroup
} from '@fluentui/react-components';
import { AgreementStatus } from '../../../../../../service/BackendServiceModels/EventModels';

export default class VotingStatus extends React.Component<IVotingStatusprops, IVotingStatusState> {

    static contextType = VotingContext;
    context!: React.ContextType<typeof VotingContext>;

    constructor(props: IVotingStatusprops) {
        super(props);
        this.state = {
            selectedStatus: undefined
        }
    }

    private closeVotingStatusModal = (): void => {
        this.context?.onDismissModal();
        setTimeout(() => {
            this.setState({selectedStatus: undefined});
        }, 300);
    }

    private saveAndCloseVotingStatusModal =  (): void => {
        this.context?.onDismissModal();
        if(this.context.votingSelected?.Id && this.state.selectedStatus){
             this.props.clickSave(this.context.votingSelected?.Id, this.state.selectedStatus);
        }
        setTimeout(() => {
            this.setState({selectedStatus: undefined});
        }, 300);    
    }

    private renderStatusOptions(): JSX.Element[]{   
        const allStatusTerms = this.context?.termStorage?.[VotingTermsIds.AgreementStatusSet];
        if(allStatusTerms){
            const finalStatusOptions = Object.keys(allStatusTerms).filter(el => el !== AgreementStatus.Inprodgress).map((agKey: string): React.JSX.Element => {
                return(
                    <Radio key={agKey} value={agKey} label={allStatusTerms[agKey].text}/>
                )
            });
    
            return finalStatusOptions;
        }else{
            return [];
        }
        
    }

    public render(): React.ReactElement<IVotingStatusprops> {
        const {
            isEditor,
            votingDialog,
            votingSelected
        } = this.context;
        const {selectedStatus} = this.state;
        const currentStatus = selectedStatus ? selectedStatus : votingSelected?.StatusId !== AgreementStatus.Inprodgress ? votingSelected?.StatusId : undefined;
        console.log(selectedStatus);
        console.log(votingSelected);
        return (
            <Dialog open={votingDialog === VotingDialogActions.VotingStatus}>
                <DialogSurface className={styles.votingStatus}>
                    {/* Cabecera */}
                    <div className={styles.votingStatusHeader}>
                        {/* Icono */}
                        <div className={styles.votingStatusLogo}>
                            {
                                <WindowAdPersonRegular />
                            }
                        </div>
                        {/* Título */}
                        <div className={styles.votingStatusTitle}>
                            {
                                strings.VotingStatusTitle
                            }
                        </div>
                        {/* Botón de cerrar */}
                        <Button
                            icon={<DismissFilled />}
                            appearance="transparent"
                            className={styles.votingStatusDismissIcon}
                            onClick={this.closeVotingStatusModal}
                            //disabled={isLoading}
                        />
                    </div>
                        <DialogContent className={styles.votingStatusBody}>
                            {/* Mensaje */}
                            <span className={styles.votingStatusMessage}>
                                {
                                    `${strings.FinishVotationChooseResult}`
                                }
                            </span>
                            {/* Usuarios que votan */}
                            <div className={styles.votingStatusOptionsContainer}>
                                <RadioGroup
                                    value={currentStatus}
                                    onChange={(_, data) => this.setState({selectedStatus:data.value})}
                                >
                                    {
                                        this.renderStatusOptions()
                                    }
                                </RadioGroup>
                            </div>
                        </DialogContent>
                        <DialogActions className={styles.votingStatusButtons}>
                            <DialogTrigger disableButtonEnhancement>
                                <Button
                                    appearance="secondary"
                                    className={styles.votingStatusSingleButton}
                                    onClick={this.closeVotingStatusModal}
                                >
                                    {strings.Cancel}
                                </Button>
                            </DialogTrigger>
                            {
                                isEditor &&
                                <Button
                                    disabled={!selectedStatus}
                                    appearance="primary"
                                    className={styles.votingStatusSingleButton}
                                    onClick={this.saveAndCloseVotingStatusModal}
                                >
                                    {strings.Accept}
                                </Button>
                            }
                        </DialogActions>
                </DialogSurface>
            </Dialog>
        );
    }

}