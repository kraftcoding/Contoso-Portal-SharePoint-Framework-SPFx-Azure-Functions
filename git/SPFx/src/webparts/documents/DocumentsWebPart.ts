import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as strings from 'DocumentsWebPartStrings';
import Documents from './components/Documents';
import SPService from '../../service/Service';
import { IDocumentsprops } from './components/IDocumentsprops';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IpropertyPaneConfiguration, propertyPaneTextField, propertyPaneSlider } from '@microsoft/sp-property-pane';
import { Fluentproviderprops, Fluentprovider, webLightTheme, Idprefixprovider } from '@fluentui/react-components';
import { providers, SharePointprovider } from '@microsoft/mgt-spfx';
export interface IDocumentsWebPartprops {
  title: string;
  docsToShow: number;
  moreLink: string;
}

export default class DocumentsWebPart extends BaseClientSideWebPart<IDocumentsWebPartprops> {

  private docsToShow_minValue: number = 1;
  private docsToShow_maxValue: number = 20;



  public render(): void {
    const element: React.ReactElement<IDocumentsprops> = React.createElement(
      Documents,
      {
        spService: this.context.serviceScope.consume(SPService.serviceKey),
        context: this.context,
        title: this.properties.title,
        docsToShow: (this.properties.docsToShow >= this.docsToShow_minValue && this.properties.docsToShow <= this.docsToShow_maxValue) ? this.properties.docsToShow : 0,
        moreLink: this.properties.moreLink,
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
        value: 'Documents-'
      },
      fluentElement
    )
    ReactDom.render(idprefixprovider, this.domElement);
  }

  protected async onInit() {
    if (!providers.globalprovider) {
      providers.globalprovider = new SharePointprovider(this.context);
    }
    return promise.resolve();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getpropertyPaneConfiguration(): IpropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: strings.Settingproperties,
              groupFields: [
                propertyPaneTextField("title", {
                  disabled: false,
                  label: strings.Title,
                  value: this.properties.title
                }),
                propertyPaneTextField("moreLink", {
                  disabled: false,
                  label: strings.MoreLinkLabel,
                  placeholder: 'https://'
                }),
                propertyPaneSlider("docsToShow", {
                  disabled: false,
                  label: strings.DocsToShow,
                  value: this.properties.docsToShow,
                  min: this.docsToShow_minValue,
                  max: this.docsToShow_maxValue,
                  showValue: true
                }),
              ]
            }
          ]
        }
      ]
    };
  }

}