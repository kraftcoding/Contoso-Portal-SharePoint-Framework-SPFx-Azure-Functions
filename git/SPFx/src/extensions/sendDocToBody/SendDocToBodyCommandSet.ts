import { Logger } from '../../utils/Logger';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs,
} from "@microsoft/sp-listview-extensibility";
import SPService from "../../service/Service";
import { IFileToCopy } from "../../models/IFileToCopy";
import SendDocuments, { ISendDocumentsprops } from "./components/SendDocuments";
import { BackendService } from "../../service/BackendService";
import React from "react";
import ReactDOM from "react-dom";
import { Fluentprovider, Fluentproviderprops, Idprefixprovider, webLightTheme } from "@fluentui/react-components";

const testFIXBAGprodP = "Bagprodp_x005f_";

export default class SendDocToBodyCommandSet extends BaseListViewCommandSet<{}> {
  private dialogPlaceHolder: HTMLDivElement;
  private bodySite: string = "";

  public async onInit(): promise<void> {
    Logger.debug("SendDocToBodyCommandSet: Initialized SendDocToBodyCommandSet", undefined, this.context);
    // initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand("COMMAND_1");
    compareOneCommand.visible = false;
    this.bodySite = await this.getBodyTeamSite();
    if (this.bodySite && this.bodySite !== "") {
      this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
      this.dialogPlaceHolder = document.body.appendChild(document.createElement("div"));
      Logger.info("SendDocToBodyCommandSet: Acceso a carpeta perteneciente a canal de teams", undefined, this.context)
      return promise.resolve();
    }
    Logger.debug("SendDocToBodyCommandSet: Carpeta no perteneciente a un canal de teams", undefined, this.context);

    return promise.reject();
  }

  private async getBodyTeamSite(): promise<string> {
    let bodySite = "";

    try {
      const spService = this.context.serviceScope.consume(SPService.serviceKey);
      const properties = await spService.getpropertiesByRootFolder();
      const teamproperties = this.mappropertiesBagprodp(properties);
      const relativePath = this.getFolderRelativePathFromURL(window.location.href);

      if (Object.keys(teamproperties)?.length > 0 && relativePath) {
        const foldersId = await spService.getFolderIdsFromPath(
          relativePath,
          this.context.pageContext.site.serverRelativeUrl
        );

        foldersId.forEach((id) => {
          if (teamproperties[id]) bodySite = teamproperties[id];
        });
      }
    } catch (ex) {
      Logger.error("SendDocToBodyCommandSet: Error to check team site. Exception: ", ex, this.context);
    }

    return bodySite;
  }

  private getFolderRelativePathFromURL(url: string) {
    const regex = /\/sites\/[^\/]+\//;
    return new URL(url).searchParams.get("id")?.replace(regex, "");
  }

  private mappropertiesBagprodp(properties: { [key: string]: string }): { [key: string]: string } {
    return Object.keys(properties)
      .filter((clave: string) => clave.startsWith(testFIXBAGprodP))
      .reduce((resultado: { [key: string]: string }, clave: string) => {
        let nuevaClave = clave.replace(testFIXBAGprodP, "");
        resultado[nuevaClave] = properties[clave];
        return resultado;
      }, {});
  }

  public async onExecute(event: IListViewCommandSetExecuteEventParameters): promise<void> {
    switch (event.itemId) {
      case "COMMAND_1":
        let filesToCopy: IFileToCopy[] = [];
        if (event.selectedRows.length > 0) {
          event.selectedRows.forEach((row) => {
            let file: IFileToCopy = {
              Id: row.getValueByName("ID"),
              Path: row.getValueByName("FileRef"),
              Title: row.getValueByName("FileLeafRef"),
              Type: row.getValueByName("File_x0020_Type"),
            };
            filesToCopy.push(file);
          });
        }

        this.renderSendDocumentsDialog(true, filesToCopy);
        break;
      default:
        Logger.error("SendDocToBodyCommandSet: executed unknown command ", undefined, this.context);
        throw new Error("Unknown command");
    }
  }

  private renderSendDocumentsDialog(isOpen: boolean, filesToCopy: IFileToCopy[]) {
    const element: React.ReactElement<ISendDocumentsprops> = React.createElement(SendDocuments, {
      isOpen: isOpen,
      onClosePanel: this.onCloseDialog,
      context: this.context,
      spService: this.context.serviceScope.consume(SPService.serviceKey),
      bkService: this.context.serviceScope.consume(BackendService.serviceKey),
      filesToCopy: filesToCopy || [],
      body: this.bodySite
    });
    const fluentElement: React.ReactElement<Fluentproviderprops> = React.createElement(
      Fluentprovider,
      {
        theme: webLightTheme,
      },
      element
    );

    const idprefixprovider: React.ReactElement<Fluentproviderprops> = React.createElement(
      Idprefixprovider,
      {
        value: "SendDoc-",
      },
      fluentElement
    );
    ReactDOM.render(idprefixprovider, this.dialogPlaceHolder);
  }

  private onCloseDialog = (): void => {
    const filestocopy: IFileToCopy[] = [];
    this.renderSendDocumentsDialog(false, filestocopy);
  };

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Logger.debug("SendDocToBodyCommandSet: List view state changed", undefined, this.context);

    const compareOneCommand: Command = this.tryGetCommand("COMMAND_1");
    if (compareOneCommand) {
      compareOneCommand.visible = this.context.listView.selectedRows?.length !== 0;
    }

    this.raiseOnChange();
  };
}
