import * as React from "react";
import * as strings from "SendDocToBodyCommandSetStrings";
import {
  Dialog,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Button,
  DialogActions,
  Dropdown,
  Option,
  Table,
  TableRow,
  TableCell,
  TableCellLayout,
  Spinner,
} from "@fluentui/react-components";
import { DocumentRegular } from "@fluentui/react-icons";
import { ISPService } from "../../../service/Service";
import { IBackendService } from "../../../service/BackendService";
import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { IFileToCopy } from "../../../models/IFileToCopy";
import { IEvent } from "../../../service/BackendServiceModels/EventModels";
import { Logger } from "../../../utils/Logger";

export interface ISendDocumentsState {
  loading: boolean;
  showMessage: boolean;
  saving: boolean;
  open: boolean;
  events: IEvent[];
  selectedOption: string;
  error: boolean;
  errorMessage: string;
}

export interface ISendDocumentsprops {
  spService: ISPService;
  context: ListViewCommandSetContext;
  bkService: IBackendService;
  isOpen: boolean;
  onClosePanel: () => void;
  filesToCopy: IFileToCopy[];
  body: string;
}

export default class SendDocumentsContent extends React.Component<ISendDocumentsprops, ISendDocumentsState> {
  constructor(props: ISendDocumentsprops) {
    super(props);

    this.state = {
      loading: false,
      showMessage: false,
      saving: false,
      open: true,
      events: [],
      selectedOption: "",
      error: false,
      errorMessage: ''
    };
  }

  public componentDidMount(): void {
    void this.onInit();
  }

  private async onInit(): promise<void> {
    try {
      this.setState({ loading: true });

      let [eventsRes]: [IEvent[]] = await promise.all([this.getEvents()]);
      
        this.setState({ events: eventsRes, loading: false });
        Logger.debug("SendDocToBodyCommandSet: events obtained", undefined, this.context);

    } catch (ex) {
      Logger.error("SendDocToBodyCommandSet: Error to obtain events. Exception: ", ex, this.context);
      this.setState({ loading: false, events: [], error: true, errorMessage: strings.ErrorTextOnLoad, showMessage: true });
    }
  }

  public closeDialog = (): void => {
    this.setState({ selectedOption: "", saving: false, showMessage: false, error: false });
    this.props.onClosePanel();
  };

  

  public async getEvents(): promise<IEvent[]> {
    const eventos: IEvent[] = await this.props.bkService.getEvents(this.props.body);
    return eventos;
  }

  private sendDocuments = async (): promise<void> => {
    try {
      this.setState({ saving: true });

      const files = this.props.filesToCopy?.map((file) => {
        return `${window.location.origin}${file.Path}`;
      });
      const res = await this.props.spService.copyToDocumentSet(files, `${window.location.origin}${this.state.selectedOption}`);
      if (res.status === 403 || res.status === 401 || res.status === 500){
        Logger.error(`SendDocToBodyCommandSet: Error ${res?.status} trying to send document.`, this.context);
        this.setState({ error: true, errorMessage: strings.ErrorTextOnSave, showMessage: true, saving: false })
      }
       
      else
      Logger.debug("SendDocToBodyCommandSet: document sent", this.context)
        this.setState({ showMessage: true, saving: false });
    } catch (ex) {
      Logger.error("SendDocToBodyCommandSet: Error to send documents. Exception: ", ex, this.context);
      this.closeDialog();
    }
  };

  public render(): React.ReactElement<ISendDocumentsprops> {
    const eventos = this.state.events;

    return (
      <Dialog open={this.props.isOpen} onOpenChange={(ev, data) => this.setState({ open: data.open })}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>{strings.DialogTitle}</DialogTitle>
            {!this.state.showMessage ? (
              <>
                <DialogContent>
                  <Table>
                    {this.props.filesToCopy?.map((file) => {
                      return (
                        <TableRow>
                          <TableCell>
                            <TableCellLayout media={<DocumentRegular />}>{file?.Title}</TableCellLayout>
                          </TableCell>
                        </TableRow>
                      );
                    })}
                  </Table>
                  <br></br>
                  {eventos?.length > 0 ? (
                    <Dropdown
                      disabled={this.state.loading}
                      placeholder={strings.DropDownPlaceholder}
                      selectedOptions={[this.state.selectedOption]}
                      onOptionSelect={(ev, data) => this.setState({ selectedOption: data.optionValue || "" })}
                    >
                      {eventos?.map((event) => {
                        return (
                          <Option key={event.StorageServerRelativeUrl} value={event.StorageServerRelativeUrl}>
                            {event.Title}
                          </Option>
                        );
                      })}
                    </Dropdown>
                  ) : (
                    <DialogContent>{this.state.loading ? <Spinner size={"tiny"}/> : strings.ErrorTextOnLoad}</DialogContent>
                  ) }

                </DialogContent>
                <DialogActions className="Dialog">
                  {this.state.saving ? (
                    <>
                      <Spinner size="small" />
                      <Button appearance="secondary" disabled={this.state.saving} onClick={this.closeDialog}>
                        {strings.ButtonCancelar}
                      </Button>
                      <Button
                        form="details"
                        disabled={this.state.saving}
                        appearance="primary"
                        onClick={this.sendDocuments}
                      >
                        {strings.ButtonAceptar}
                      </Button>
                    </>
                  ) : (
                    <>
                      <Button appearance="secondary" disabled={this.state.saving} onClick={this.closeDialog}>
                        {strings.ButtonCancelar}
                      </Button>
                      <Button
                        form="details"
                        disabled={!this.state.selectedOption}
                        appearance="primary"
                        onClick={this.sendDocuments}
                      >
                        {strings.ButtonAceptar}
                      </Button>
                    </>
                  )}
                </DialogActions>
              </>
            ) : (
              <>
                <DialogContent>{this.state.error ? this.state.errorMessage : strings.MessageOK}</DialogContent>
                <DialogActions className="Dialog">
                  <Button appearance="primary" onClick={this.closeDialog}>
                    {strings.ButtonAceptar}
                  </Button>
                </DialogActions>
              </>
            )}
          </DialogBody>
        </DialogSurface>
      </Dialog>
    );
  }
}
