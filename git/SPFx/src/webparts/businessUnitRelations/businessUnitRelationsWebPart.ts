import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IpropertyPaneConfiguration,
  propertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'BusinessAreaRelationsWebPartStrings';
import BusinessAreaRelations from './components/BusinessAreaRelations';
import { IBusinessAreaRelationsprops } from './components/IBusinessAreaRelations';
import { providers, SharePointprovider } from '@microsoft/mgt-spfx';
import { BackendService } from '../../service/BackendService';
import SPService from '../../service/Service';
import { Fluentprovider, Fluentproviderprops, Idprefixprovider, webLightTheme } from '@fluentui/react-components';
export interface IBusinessAreaRelationsWebPartprops {
  title: string;
}

export default class BusinessAreaRelationsWebPart extends BaseClientSideWebPart<IBusinessAreaRelationsWebPartprops> {

  protected async onInit(): promise<void> {
    if (!providers.globalprovider) {
      providers.globalprovider = new SharePointprovider(this.context);
    }
    return promise.resolve();
  }

  public render(): void {
    const { context, properties } = this;
    const BusinessAreaRelationsprops:IBusinessAreaRelationsprops={
      context,
      spService: context.serviceScope.consume(SPService.serviceKey),
      bkService: context.serviceScope.consume(BackendService.serviceKey),
      localLanguage: context.pageContext.cultureInfo.currentUICultureName,
      title: properties.title,
      isMobile: window.innerWidth < 640,
    }
    const reactElement: React.ReactElement<IBusinessAreaRelationsprops> = React.createElement(
      BusinessAreaRelations,
      BusinessAreaRelationsprops
    );

    const fluentElement: React.ReactElement<Fluentproviderprops> = React.createElement(
      Fluentprovider,
      { theme: webLightTheme },
      reactElement
    );

    const idprefixprovider: React.ReactElement<Fluentproviderprops> = React.createElement(
      Idprefixprovider,
      { value: 'BusinessAreaRelations-' },
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
          header: {
            description: strings.propertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                propertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
