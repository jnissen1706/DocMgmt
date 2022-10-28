import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'DocMgmtWebPartStrings';
import DocMgmt from './components/DocMgmt';
import { IDocMgmtProps } from './components/IDocMgmtProps';
import { getSP } from './pnpjsConfig';
import { MSGraphClientV3 } from "@microsoft/sp-http";

export interface IDocMgmtWebPartProps {
  description: string;
}

export default class DocMgmtWebPart extends BaseClientSideWebPart<IDocMgmtWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _graphClient: MSGraphClientV3 = null;

  public render(): void {
    const element: React.ReactElement<IDocMgmtProps> = React.createElement(
      DocMgmt,
      {
        context: this.context,
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        graphClient: this._graphClient
      }
    );

    let fakeElement = this.startGraphClient(element);

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    await super.onInit();

    //Initialize our _sp object that we can then use in other packages without having to pass around the context.
    //  Check out pnpjsConfig.ts for an example of a project setup file.
    getSP(this.context);
}

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  //For Graph Client initialization. Currently not functioning correctly and returns a null element.
  protected startGraphClient(InElement?:React.ReactElement<IDocMgmtProps>):React.ReactElement<IDocMgmtProps> {
    try {
        const context1 = InElement.props.context;
        let element: React.ReactElement<IDocMgmtProps> = null;
        const props: Promise<React.ReactElement<IDocMgmtProps>> = new Promise<React.ReactElement<IDocMgmtProps>>(
            (resolve, reject) => {
                context1.msGraphClientFactory
            .getClient('3')
            .then((client: MSGraphClientV3): void => {
            const element1: React.ReactElement<IDocMgmtProps> = React.createElement(DocMgmt,
            {
                context: context1,
                description: InElement.props.description,
                isDarkTheme: InElement.props.isDarkTheme,
                environmentMessage: InElement.props.environmentMessage,
                graphClient: client
                });
            element = element1;
            }).catch((errorL) => {
              console.log(``);
        });
    });

    return element;
}
    catch(ex) {
        console.log(`Error`)
    }
}

}
