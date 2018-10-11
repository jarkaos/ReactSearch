import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'PropertyPaneTextDialogWebPartStrings';
import PropertyPaneTextDialog from './components/PropertyPaneTextDialog';
import { IPropertyPaneTextDialogProps } from './components/IPropertyPaneTextDialogProps';

export interface IPropertyPaneTextDialogWebPartProps {
  description: string;
}

export default class PropertyPaneTextDialogWebPart extends BaseClientSideWebPart<IPropertyPaneTextDialogWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPropertyPaneTextDialogProps > = React.createElement(
      PropertyPaneTextDialog,
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
}
