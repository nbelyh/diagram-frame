import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IPropertyPaneField, PropertyPaneFieldType, IPropertyPaneCustomFieldProps } from '@microsoft/sp-property-pane';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react';

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/files";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { PropertyPaneUrlFieldComponent } from './PropertyPaneUrlFieldComponent';

export function PropertyPaneUrlField(targetProperty: string, props: {
  url: string;
  context: WebPartContext;
}): IPropertyPaneField<IPropertyPaneCustomFieldProps> {

  return {
    targetProperty: targetProperty,
    type: PropertyPaneFieldType.Custom,
    properties: {
      key: 'url',

      onRender: (parent: HTMLElement, context: any, changeCallback: (targetProperty: string, newValue: any) => void) => {
        return ReactDom.render(<PropertyPaneUrlFieldComponent context={props.context} url={props.url} setUrl={(url) => changeCallback('url', url)} />, parent);
      },

      onDispose(parent: HTMLElement): void {
        ReactDom.unmountComponentAtNode(parent);
      }
    }
  };
}
