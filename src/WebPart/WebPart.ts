import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { PropertyPaneConfiguration } from './properties/PropertyPaneConfiguration';
import { sp } from '@pnp/sp';

require('VisioEmbed');

import { TopFrame } from './TopFrame';
import { IWebPartProps } from './IWebPartProps';

export default class WebPart extends BaseClientSideWebPart<IWebPartProps> {

  public onInit(): Promise<void> {

    return super.onInit().then(() => {
      sp.setup({ spfxContext: this.context as any });
    });
  }

  public render(): void {

    const properties = {
      ...this.properties,
      width: this.properties.width || '100%',
      height: this.properties.height || '50vh',
      openHyperlinksInNewWindow: typeof(this.properties.openHyperlinksInNewWindow) === 'undefined' ? true : this.properties.openHyperlinksInNewWindow,
      forceOpeningOfficeFilesOnline: typeof(this.properties.forceOpeningOfficeFilesOnline) === 'undefined' ? true : this.properties.forceOpeningOfficeFilesOnline,
    };

    const element = React.createElement(TopFrame, {
      ...properties,
      isReadOnly: this.displayMode === DisplayMode.Read,
      context: this.context,
    });

    ReactDom.render(element, this.domElement);
  }

  public onPropertyPaneConfigurationStart() {
    this.render();
  }

  public onPropertyPaneConfigurationComplete() {
    this.render();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return PropertyPaneConfiguration.get(this.context, this.properties)
  }
}
