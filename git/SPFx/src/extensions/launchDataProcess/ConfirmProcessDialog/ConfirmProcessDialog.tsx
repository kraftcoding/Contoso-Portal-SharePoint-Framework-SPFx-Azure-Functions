import * as React from 'react';
import { IConfirmprocessDialogprops, IConfirmprocessDialogState } from './IConfirmprocessDialog';
import styles from './ConfirmprocessDialog.module.scss';
import { Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Spinner } from '@fluentui/react-components';
import IQuery, { IQueryOrder } from '../../../models/IQuery';
import strings from 'LaunchDataprocessCommandSetStrings';

const listRequestUrl="/Lists/RequestesAdministracion";
const listRequestQuery:IQuery ={
  viewFields: [
      "OperacionPeticion",
      "EstadoPeticion",
  ],
  expand: [],
  filter: "(OperacionPeticion eq 'powerbireport') and (EstadoPeticion eq 'In prodgress')",
};
class ConfirmprocessDialog extends React.Component<IConfirmprocessDialogprops, IConfirmprocessDialogState> {

  constructor(props: IConfirmprocessDialogprops | Readonly<IConfirmprocessDialogprops>) {
    super(props);
    this.state = {
      running: false,
      loading: false,
      error: false,
      open: true,
      launched: false
    };
  }

    public componentDidMount(): void {
      void this.onInit();
    }
    componentDidUpdate(testvprops: Readonly<IConfirmprocessDialogprops>): void {
      if(!testvprops.isOpen && this.props.isOpen){
        void this.onInit();
      }
    }
  
    private async onInit(): promise<void> {
      this.setState({ loading: true, error:false, launched: false});
      this.props.spService.getItems(`${this.props.context.pageContext.site.serverRelativeUrl}${listRequestUrl}`, listRequestQuery).then(res =>{
        if(res && res.length>0){
          this.setState({loading:false, running: true, launched: false});
        }else{
          this.setState({loading:false, running: false, launched: false});
        }
      }).catch(ex => {
        this.setState({loading:false, error: true, running: false,launched: false});
        console.log(ex);
      })
    }

  public render(): JSX.Element {
    const {error, running, loading, launched} = this.state;
    return (
      <Dialog open={this.props.isOpen} onOpenChange={(ev, data) => {this.setState({ open: data.open })}}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>{strings.ModalTitle}</DialogTitle>
            <DialogContent>
              {loading?
                <Spinner/>
              :
                <div className={styles.confirmprocessDialog}>
                  {error?
                    <div className={styles.errorprocess}>
                      {strings.ErrorMsg}
                    </div>
                  :
                  running ?
                    <div className={styles.runningprocess}>
                      {strings.RunningMsg}
                    </div>
                  :
                  launched?
                    <div className={styles.textprocessDialog}>
                      {strings.LaunchedMsg}
                    </div>
                  :
                    <div className={styles.textprocessDialog}>
                      <div>{strings.processType}<b>{this.props.isComplete ? strings.CompleteLabel : strings.IncrementalLabel}</b></div>
                      <br/>
                      <div>{strings.AreYouSure}</div>
                      
                    </div>
                  }
                </div>
              }
            </DialogContent>
            <DialogActions className="Dialog">
              <Button appearance="secondary" 
                //disabled={this.state.saving} 
                onClick={this.closeDialog}>
                {launched ? "Cerrar": strings.CancelButton}
              </Button>
              <Button
                form="details"
                disabled={loading || running || error || launched}
                appearance="primary"
                onClick={this.sendprocessRequest}
              >
                {strings.AcceptButton}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    );
  }

  private sendprocessRequest = async (): promise<void> => {
    if(this.props.isComplete){
      void this.props.bkService.launchSaveDataprocess();
    }else{
      const lastDate = await this.getLastDateprocess();
       console.log(lastDate);
      if(lastDate){
        void this.props.bkService.launchSaveDataprocess(lastDate);
      }else{
        void this.props.bkService.launchSaveDataprocess();
      }
    }
    this.setState({loading: true});
    setTimeout(() => {
      this.setState({launched: true, loading:false, running: false, error: false});
    }, 300);
  }

  
private getLastDateprocess = async (): promise<string> => {
  const listReportsUrl = "/processReports";
  const listReportsQuery: IQueryOrder = {
    viewFields: ["Title", "Inicioprodceso","EstadoPeticion"],
    orderBy: "Inicioprodceso",
    orderByAscending: false,
    expand: [],
    filter: "EstadoPeticion eq 'OK'",
    top: 1
  };
  try {
    const res = await this.props.spService.getItemsByOrder(
      `${this.props.context.pageContext.site.serverRelativeUrl}${listReportsUrl}`,
      listReportsQuery
    );

    if (res && res.length > 0) {
      const date = (res as any[])[0]["Inicioprodceso"];
      return date ? date : "";
    } else {
      return "";
    }
  } catch (ex) {
    this.setState({ loading: false, error: true, running: false });
    console.log(ex);
    return "";
  }
};


  public closeDialog = (): void => {
    this.props.onClosePanel(this.props.isComplete);
  };
}

export default ConfirmprocessDialog;
