import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IpropertyPaneConfiguration,
  propertyPaneSlider,
  propertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CertificationAppWebPartStrings';
import SPService from '../../service/Service';
import CertificationApp from './components/CertificationApp';
import { ICertificationAppprops } from './components/ICertificationApp';
import { BackendService } from '../../service/BackendService';
import { Fluentproviderprops, Fluentprovider, webLightTheme, Idprefixprovider } from '@fluentui/react-components';
import { providers, SharePointprovider, LocalizationHelper } from '@microsoft/mgt-spfx';

// eslint-disable-next-line @typescript-eslint/ban-ts-comment     
// @ts-ignore
LocalizationHelper.strings = {
  _components: {
    "people-picker": {
      noResultsFound: strings.NoResultsFoundPicker,
      loadingMessage: strings.LoadingPicker,
      maxSelectionsPlaceHolder: ""
    },
  }
}

export interface ICertificationAppWebPartprops {
  componentTitle: string;
  itemsPerPage: number;
}

export default class CertificationAppWebPart extends BaseClientSideWebPart<ICertificationAppWebPartprops> {
  protected async onInit(): promise<void> {
    if (!providers.globalprovider) {
      providers.globalprovider = new SharePointprovider(this.context);
    }
    return promise.resolve();
  }

  public render(): void {
    const { context, properties } = this;
    const certificationAppprops: ICertificationAppprops=
      {
        context,
        spService: context.serviceScope.consume(SPService.serviceKey),
        bkService: context.serviceScope.consume(BackendService.serviceKey),
        localLanguage: context.pageContext.cultureInfo.currentUICultureName,
        componentTitle: properties.componentTitle,
        isMobile: window.innerWidth < 640,
        itemsPerPage: properties.itemsPerPage ? properties.itemsPerPage : 10,
      }

      const reactElement: React.ReactElement<ICertificationAppprops> = React.createElement(
        CertificationApp,
        certificationAppprops
      );
  
      const fluentElement: React.ReactElement<Fluentproviderprops> = React.createElement(
        Fluentprovider,
        { theme: webLightTheme },
        reactElement
      );
  
      const idprefixprovider: React.ReactElement<Fluentproviderprops> = React.createElement(
        Idprefixprovider,
        { value: 'CertificationApp-' },
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
              groupName: strings.BasicGroupName,
              groupFields: 
              [
                propertyPaneTextField("componentTitle", {
                  disabled: false,
                  label: strings.TitlePane,
                  value: this.properties.componentTitle
                }),
                propertyPaneSlider("itemsPerPage",{
                  min: 2,
                  max: 100,
                  label: strings.ItemsPerPage,
                  value: this.properties.itemsPerPage
                })
              ]
            } 
          ]
        }
      ]
    };
  }
}
