import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as strings from 'MeetingsWebPartStrings';
import Meetings from './components/Meetings';
import SPService from '../../service/Service';
import { BackendService } from '../../service/BackendService';
import { IMeetingsprops } from './components/IMeetings';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IpropertyPaneConfiguration, propertyPaneTextField, propertyPaneToggle } from '@microsoft/sp-property-pane';
import { Fluentprovider, Fluentproviderprops, Idprefixprovider, webLightTheme } from '@fluentui/react-components';
import { providers, SharePointprovider, LocalizationHelper } from '@microsoft/mgt-spfx';

// eslint-disable-next-line @typescript-eslint/ban-ts-comment     
// @ts-ignore
LocalizationHelper.strings = {
  _components: {
    "people-picker": {
      noResultsFound: strings.NoResultsFound,
      loadingMessage: strings.Loading + "...",
      maxSelectionsPlaceHolder: ""
    },
    "file-list": {
      showMoreSubtitle: strings.ShowMoreSubtitle
    },
    "file-upload": {
      failUploadFile: strings.FailUploadFile,
      cancelUploadFile: strings.CancelUploadFile,
      buttonUploadFile: strings.ButtonUploadFile,
      maximumFilesTitle: strings.MaximumFilesTitle,
      maximumFiles: strings.MaximumFilesMessage,
      maximumFileSizeTitle: strings.MaximumFileSizeTitle,
      maximumFileSize: strings.MaximumFileSizeMessage,
      fileTypeTitle: strings.FileTypeTitle,
      fileType: strings.FileTypeMessage,
      checkAgain: strings.CheckAgain,
      checkApplyAll: strings.CheckApplyAll,
      buttonOk: strings.ButtonOk,
      buttonCancel: strings.Cancel,
      buttonUpload: strings.Upload,
      buttonKeep: strings.ButtonKeep,
      buttonReplace: strings.Replace,
      buttonReselect: strings.Reselect,
      fileReplaceTitle: strings.FileReplaceTitle,
      fileReplace: strings.FileReplaceMessage,
      uploadButtonLabel: strings.UploadButtonLabel
    },
    "file": {
      modifiedSubtitle: 'Modificado'
    }
  }
}

export interface IMeetingsWebPartprops {
  title: string;
  showVoting: boolean;
}

export default class MeetingsWebPart extends BaseClientSideWebPart<IMeetingsWebPartprops> {

  protected async onInit() {
    if (!providers.globalprovider) {
      providers.globalprovider = new SharePointprovider(this.context);
    }
    return promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<IMeetingsprops> = React.createElement(
      Meetings,
      {
        spService: this.context.serviceScope.consume(SPService.serviceKey),
        bkService: this.context.serviceScope.consume(BackendService.serviceKey),
        title: this.properties.title,
        displayMode: this.displayMode,
        updateproperty: (value: string) => {
          this.properties.title = value;
        },
        context: this.context,
        locale: this.context.pageContext.cultureInfo.currentUICultureName,
        showVoting: this.properties.showVoting ? true : false
      }
    );

    const fluentElement: React.ReactElement<Fluentproviderprops> = React.createElement(
      Fluentprovider,
      {
        theme: webLightTheme
      },
      element
    );

    const idprefixprovider: React.ReactElement<Fluentproviderprops> = React.createElement(
      Idprefixprovider,
      {
        value: 'Meetings-'
      },
      fluentElement
    )

    ReactDom.render(idprefixprovider, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getpropertyPaneConfiguration(): IpropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                propertyPaneTextField('title', {
                  label: "Título"
                }),
                propertyPaneToggle("showVoting", {
                  key: "showVoting",
                  label: "Mostrar funcionalidad de voto",
                  checked: this.properties.showVoting,
                })
              ]
            }
          ]
        }
      ]
    };
  }

}