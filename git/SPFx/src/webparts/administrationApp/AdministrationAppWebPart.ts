import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as strings from 'AdministrationAppWebPartStrings';
import SPService from '../../service/Service';
import { BackendService } from '../../service/BackendService';
import AdministrationApp from './components/AdministrationApp';
import { IAdministrationAppprops } from './components/IAdministrationApp';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  type IpropertyPaneConfiguration,
  propertyPaneCheckbox,
  propertyPaneDropdown,
  propertyPaneSlider,
  propertyPaneTextField
} from '@microsoft/sp-property-pane';
import { Fluentproviderprops, Fluentprovider, webLightTheme, Idprefixprovider } from '@fluentui/react-components';
import { providers, SharePointprovider, LocalizationHelper } from '@microsoft/mgt-spfx';
import { componentTypeMode } from './models/AdministrationAppModels';

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
export interface IAdministrationAppWebPartprops {
  componentTitle: string;
  componentType: string;
  itemsPerPage: number;
  showEXTERNAL: boolean;
}

export default class AdministrationAppWebPart extends BaseClientSideWebPart<IAdministrationAppWebPartprops> {
  protected async onInit(): promise<void> {
    if (!providers.globalprovider) {
      providers.globalprovider = new SharePointprovider(this.context);
    }
    return promise.resolve();
  }

  public render(): void {
    const { context, properties } = this;

    const administrationAppprops: IAdministrationAppprops = {
      context,
      spService: context.serviceScope.consume(SPService.serviceKey),
      bkService: context.serviceScope.consume(BackendService.serviceKey),
      localLanguage: context.pageContext.cultureInfo.currentUICultureName,
      componentTitle: properties.componentTitle,
      componentMode: properties.componentType ? properties.componentType : componentTypeMode.All,
      isMobile: window.innerWidth < 640,
      itemsPerPage: properties.itemsPerPage ? properties.itemsPerPage : 10,
      showEXTERNAL: false,
    };

    const reactElement: React.ReactElement<IAdministrationAppprops> = React.createElement(
      AdministrationApp,
      administrationAppprops
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
              this.properties.componentType !== componentTypeMode.Organs ? [
                propertyPaneTextField("componentTitle", {
                  disabled: false,
                  label: strings.TitlePane,
                  value: this.properties.componentTitle
                }),
                propertyPaneDropdown("componentType", {
                  label: strings.ComponentTypePane,
                  options:[{key:componentTypeMode.All, text:"Completo"},{key:componentTypeMode.Organs, text:"Órganos"},{key:componentTypeMode.Users, text:"Usuarios"}],
                  selectedKey: this.properties.componentType ? this.properties.componentType: componentTypeMode.All
                }),
                propertyPaneSlider("itemsPerPage",{
                  min: 2,
                  max: 100,
                  label: strings.ItemsPerPage,
                  value: this.properties.itemsPerPage
                })
              ]
              :
              [
                propertyPaneTextField("componentTitle", {
                  disabled: false,
                  label: strings.TitlePane,
                  value: this.properties.componentTitle
                }),
                propertyPaneDropdown("componentType", {
                  label: strings.ComponentTypePane,
                  options:[{key:componentTypeMode.All, text:"Completo"},{key:componentTypeMode.Organs, text:"Órganos"},{key:componentTypeMode.Users, text:"Usuarios"}],
                  selectedKey: this.properties.componentType ? this.properties.componentType: componentTypeMode.All
                }),
                propertyPaneCheckbox("showEXTERNAL", {
                  text: strings.ComponentTypePane,
                  checked: this.properties.showEXTERNAL ,
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