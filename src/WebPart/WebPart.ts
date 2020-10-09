import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from "@microsoft/sp-property-pane";
import { PropertyFieldFilePicker, IFilePickerResult } from "@pnp/spfx-property-controls/lib/propertyFields/filePicker";
import { sp } from '@pnp/sp';

require('VisioEmbed');

import * as strings from 'WebPartStrings';

import { TopFrame } from './TopFrame';
import { PropertyPaneVersionField } from '../PropertyPaneVersionField';

export interface IVisioOnlineScriptProps {
  width: string;
  height: string;
  showToolbars: boolean;
  showBorders: boolean;
  filePickerResult: IFilePickerResult;
  zoom: number;
}

export default class WebPart extends BaseClientSideWebPart<IVisioOnlineScriptProps> {

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
      sp.setup({ spfxContext: this.context });
    });
  }

  public render(): void {

    const element = React.createElement(
      TopFrame,
      {
        width: this.properties.width,
        height: this.properties.height,
        context: this.context,
        filePickerResult: this.properties.filePickerResult,
        showToolbars: this.properties.showToolbars,
        showBorders: this.properties.showBorders,
        zoom: +this.properties.zoom
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
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [

                PropertyFieldFilePicker('filePicker', {
                  context: this.context,
                  filePickerResult: this.properties.filePickerResult,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: (e: IFilePickerResult) => { console.log(`Save: ${e}`); this.properties.filePickerResult = e; },
                  onChanged: (e: IFilePickerResult) => { console.log(`Changed: ${e}`); this.properties.filePickerResult = e; },
                  key: "filePickerId",
                  accepts: [".vsd", ".vsdx", ".vsdm"],
                  buttonLabel: strings.FieldVisioFileBrowse,
                  label: strings.FieldVisioFile,
                }),

                PropertyPaneTextField('width', {
                  label: strings.FieldWidth,
                }),

                PropertyPaneTextField('height', {
                  label: strings.FieldHeight,
                }),

                PropertyPaneTextField('zoom', {
                  label: strings.FieldZoom,
                }),
              ]
            },
            {
              groupName: strings.Toolbars,
              groupFields: [
                PropertyPaneToggle('showToolbars', {
                  label: strings.FieldShowToolbars,
                }),

                PropertyPaneToggle('showBorders', {
                  label: strings.FieldShowBorders,
                }),
              ]
            },
            {
              groupName: "About",
              groupFields: [
                PropertyPaneVersionField(this.context.manifest.version)
              ]
            }
          ]
        }
      ]
    };
  }
}
