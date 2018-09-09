import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneButton
} from '@microsoft/sp-webpart-base';

require('VisioEmbed');

import * as strings from 'VisioOnlineSpfxStrings';
import { PropertyFieldDocumentPicker } from './PropertyFieldDocumentPicker/PropertyFieldDocumentPicker';

import { IVisioOnlineSpfxWebPartProps } from './VisioOnlineSpfxWebPart/IVisioOnlineSpfxWebPartProps';
import VisioOnlineSpfxWebPart from './VisioOnlineSpfxWebPart/VisioOnlineSpfxWebPart';

export interface IVisioOnlineScriptProps {
  url: string;
  width: string;
  height: string;
}

export default class VisioOnlineScript extends BaseClientSideWebPart<IVisioOnlineScriptProps> {

  public render(): void {

    const element: React.ReactElement<IVisioOnlineSpfxWebPartProps> = React.createElement(
      VisioOnlineSpfxWebPart,
      {
        url: this.properties.url,
        width: this.properties.width,
        height: this.properties.height
      }
    );

    ReactDom.render(element, this.domElement);
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

                PropertyFieldDocumentPicker('url', {
                  label: 'Select a document',
                  initialValue: this.properties.url,
                  allowedFileExtensions: '.vsd,.vsdx,.vst,.vstx',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  render: this.render.bind(this),
                  disableReactivePropertyChanges: this.disableReactivePropertyChanges,
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  readOnly: false,
                  key: 'picturePickerFieldId'
                }),

                PropertyPaneTextField('width', {
                  label: 'Width',
                  value: '100%',
                  disabled: false,
                }),

                PropertyPaneTextField('height', {
                  label: 'Height',
                  value: '500px',
                  disabled: false,
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
