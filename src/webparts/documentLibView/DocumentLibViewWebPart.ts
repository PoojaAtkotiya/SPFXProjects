import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DocumentLibViewWebPartStrings';
import DocumentLibView from './components/DocumentLibView';
import { IDocumentLibViewProps } from './components/IDocumentLibViewProps';

export interface IDocumentLibViewWebPartProps {
  description: string;
  listName:string;
}

export default class DocumentLibViewWebPart extends BaseClientSideWebPart <IDocumentLibViewWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IDocumentLibViewProps> = React.createElement(
      DocumentLibView,
      {
        description: this.properties.description,
        siteUrl:this.context.pageContext.web.absoluteUrl,
        listName:this.properties.listName
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
}
