import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IpropertyPaneConfiguration,
  propertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'departmentManagementWebPartStrings';
import departmentManagement from './components/departmentManagement';
import { IdepartmentManagementprops } from './components/IdepartmentManagement';
import { providers, SharePointprovider, LocalizationHelper } from '@microsoft/mgt-spfx';
import SPService from '../../service/Service';
import { BackendService } from '../../service/BackendService';
import { Fluentprovider, Fluentproviderprops, Idprefixprovider, webLightTheme } from '@fluentui/react-components';

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

export interface IdepartmentManagementWebPartprops {
  componentTitle: string;
}

export default class departmentManagementWebPart extends BaseClientSideWebPart<IdepartmentManagementWebPartprops> {
  protected async onInit(): promise<void> {
    if (!providers.globalprovider) {
      providers.globalprovider = new SharePointprovider(this.context);
    }
    return promise.resolve();
  }

  public render(): void {
    const { context, properties } = this;

    const departmentManagementprops:IdepartmentManagementprops =  {
      context,
      spService: context.serviceScope.consume(SPService.serviceKey),
      bkService: context.serviceScope.consume(BackendService.serviceKey),
      localLanguage: context.pageContext.cultureInfo.currentUICultureName,
      componentTitle: properties.componentTitle,
      isMobile: window.innerWidth < 640,
    }

    const reactElement: React.ReactElement<IdepartmentManagementprops> = React.createElement(
      departmentManagement,
      departmentManagementprops
    );

    const fluentElement: React.ReactElement<Fluentproviderprops> = React.createElement(
      Fluentprovider,
      { theme: webLightTheme },
      reactElement
    );

    const idprefixprovider: React.ReactElement<Fluentproviderprops> = React.createElement(
      Idprefixprovider,
      { value: 'AdministrationApp-' },
      fluentElement
    )

    ReactDom.render(idprefixprovider, this.domElement);
    /*
    const element: React.ReactElement<IdepartmentManagementprops> = React.createElement(
      departmentManagement,
      {
        context,
      spService: context.serviceScope.consume(SPService.serviceKey),
      bkService: context.serviceScope.consume(BackendService.serviceKey),
      localLanguage: context.pageContext.cultureInfo.currentUICultureName,
      componentTitle: properties.componentTitle,
      isMobile: window.innerWidth < 640,
      }
    );

    ReactDom.render(element, this.domElement);*/
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
          header: {
            description: strings.propertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                propertyPaneTextField("componentTitle", {
                  disabled: false,
                  label: strings.TitlePane,
                  value: this.properties.componentTitle
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
