import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as strings from 'myDepartmentsWebPartStrings';
import myDepartments from './components/myDepartments';
import SPService from '../../service/Service';
import { ImyDepartmentsprops } from './components/ImyDepartmentsprops';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IpropertyPaneConfiguration, propertyPaneLabel, propertyPaneTextField, propertyPaneSlider } from '@microsoft/sp-property-pane';
import { Fluentproviderprops, Fluentprovider, webLightTheme, Idprefixprovider } from '@fluentui/react-components';
import { BackendService } from '../../service/BackendService';

export interface ImyDepartmentsWebPartprops {
  title: string;
  slidesToShow: number;
  slidesToScroll: number;
}

export default class myDepartmentsWebPart extends BaseClientSideWebPart<ImyDepartmentsWebPartprops> {

  private slidesToShow_minValue: number = 1;
  private slidesToShow_maxValue: number = 4;
  private slidesToScroll_minValue: number = 1;
  private slidesToScroll_maxValue: number = 4;

  public render(): void {
    const element: React.ReactElement<ImyDepartmentsprops> = React.createElement(
      myDepartments,
      {
        spService: this.context.serviceScope.consume(SPService.serviceKey),
        bkService: this.context.serviceScope.consume(BackendService.serviceKey),
        context: this.context,
        title: this.properties.title,
        slidesToShow: (this.properties.slidesToShow >= this.slidesToShow_minValue && this.properties.slidesToShow <= this.slidesToShow_maxValue) ? this.properties.slidesToShow : 0,
        slidesToScroll: (this.properties.slidesToScroll >= this.slidesToScroll_minValue && this.properties.slidesToScroll <= this.slidesToScroll_maxValue) ? (this.properties.slidesToScroll <= this.properties.slidesToShow ? this.properties.slidesToScroll : 0) : 0,
        locale: this.context.pageContext.cultureInfo.currentUICultureName,
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
        value: 'myDepartments-'
      },
      fluentElement
    )
    ReactDom.render(idprefixprovider, this.domElement);
  }

  protected onInit(): promise<void> {
    return super.onInit();
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
                propertyPaneSlider("slidesToShow", {
                  disabled: false,
                  label: strings.SlidesToShow,
                  value: this.properties.slidesToShow,
                  min: this.slidesToShow_minValue,
                  max: this.slidesToShow_maxValue,
                  showValue: true
                }),
                propertyPaneSlider("slidesToScroll", {
                  disabled: false,
                  label: strings.SlidesToScroll,
                  value: this.properties.slidesToScroll,
                  min: this.slidesToScroll_minValue,
                  max: this.slidesToScroll_maxValue,
                  showValue: true
                }),
                propertyPaneLabel("", {
                  text: strings.SlideInformation
                })
              ]
            }
          ]
        }
      ]
    };
  }

}