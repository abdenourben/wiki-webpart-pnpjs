import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'WikiWebpartPnpjsWebPartStrings';
import WikiWebpartPnpjs from './components/WikiWebpartPnpjs';
import { IWikiWebpartPnpjsProps } from './components/IWikiWebpartPnpjsProps';

import { sp } from '@pnp/sp'; 

export interface IWikiWebpartPnpjsWebPartProps {
  description: string;
}

export default class WikiWebpartPnpjsWebPart extends BaseClientSideWebPart <IWikiWebpartPnpjsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IWikiWebpartPnpjsProps> = React.createElement(
      WikiWebpartPnpjs,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
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

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
  
      // other init code may be present
      sp.setup({
        spfxContext: this.context
      });
    });
  }
}
