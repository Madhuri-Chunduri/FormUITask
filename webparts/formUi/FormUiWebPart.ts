import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'FormUiWebPartStrings';
import FormUi from './components/FormUi';
import { IFormUiProps } from './components/IFormUiProps';
import * as pnp from 'sp-pnp-js';
import DropDownComponent from './components/DropDown';

export interface IFormUiWebPartProps {
  redirectionUrl: string,
  context: WebPartContext
}

export default class FormUiWebPart extends BaseClientSideWebPart<IFormUiWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IFormUiProps> = React.createElement(
      FormUi,
      {
        redirectionUrl: this.properties.redirectionUrl,
        context: this.context
      }
    );
    console.log("WebPart Context : ", this.context);
    ReactDom.render(element, this.domElement);
  }

  public async onInit(): Promise<void> {
    const _ = await super.onInit();
    pnp.setup({
      spfxContext: this.context
    });
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
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('redirectionUrl', {
                  label: "Redirection Url"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
