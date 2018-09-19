/**
 * @file PropertyFieldDocumentPickerHost.tsx
 * Renders the controls for PropertyFieldDocumentPicker component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldDocumentPickerPropsInternal } from './PropertyFieldDocumentPicker';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { IconButton, DefaultButton, PrimaryButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Async } from 'office-ui-fabric-react/lib/Utilities';

import * as strings from 'VisioOnlineSpfxStrings';

/**
 * @interface
 * PropertyFieldDocumentPickerHost properties interface
 *
 */
export interface IPropertyFieldDocumentPickerHostProps extends IPropertyFieldDocumentPickerPropsInternal {
}

export interface IPropertyFieldDocumentPickerHostState {
  openPanel?: boolean;
  selectedDocument: string;
  errorMessage?: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldDocumentPicker component
 */
export default class PropertyFieldDocumentPickerHost extends React.Component<IPropertyFieldDocumentPickerHostProps, IPropertyFieldDocumentPickerHostState> {

  private latestValidateValue: string;
  private async: Async;
  private delayedValidate: (value: string) => void;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldDocumentPickerHostProps) {
    super(props);

    //Bind the current object to the external called onSelectDate method
    this.onTextFieldChanged = this.onTextFieldChanged.bind(this);
    this.onOpenPanel = this.onOpenPanel.bind(this);
    this.onClosePanel = this.onClosePanel.bind(this);
    this.handleIframeData = this.handleIframeData.bind(this);
    this.onEraseButton = this.onEraseButton.bind(this);

    //Inits the state
    this.state = {
      selectedDocument: this.props.initialValue,
      openPanel: false,
      errorMessage: ''
    };

    this.async = new Async(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);
  }


  /**
   * @function
   * Save the document value
   *
   */
  private saveDocumentProperty(documentUrl: string): void {
    this.setState({ selectedDocument: documentUrl });
    this.delayedValidate(documentUrl);
  }

  /**
   * @function
   * Validates the new custom field value
   */
  private validate(value: string): void {
    if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
      this.notifyAfterValidate(this.props.initialValue, value);
      return;
    }

    if (this.latestValidateValue === value)
      return;
    this.latestValidateValue = value;

    var result: string | PromiseLike<string> = this.props.onGetErrorMessage(value || '');
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.props.initialValue, value);
        this.setState({ errorMessage: result });
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.props.initialValue, value);
          this.setState({ errorMessage: errorMessage });
        });
      }
    }
    else {
      this.notifyAfterValidate(this.props.initialValue, value);
    }
  }

  /**
   * @function
   * Notifies the parent Web Part of a property value change
   */
  private notifyAfterValidate(oldValue: string, newValue: string) {
    if (this.props.onPropertyChange && newValue != null) {
      this.props.properties[this.props.targetProperty] = newValue;
      this.props.onPropertyChange(this.props.targetProperty, oldValue, newValue);
      if (!this.props.disableReactivePropertyChanges && this.props.render != null)
        this.props.render();
    }
  }

  /**
  * @function
  * Click on erase button
  *
  */
  private onEraseButton(): void {
    this.saveDocumentProperty('');
  }

  /**
  * @function
  * Open the panel
  *
  */
  private onOpenPanel(element?: any): void {
    this.setState({ openPanel: true });
  }

  /**
  * @function
  * The text field value changed
  *
  */
  private onTextFieldChanged(newValue: string): void {
    this.saveDocumentProperty(newValue);
  }

  /**
  * @function
  * Close the panel
  *
  */
  private onClosePanel(element?: any): void {
    this.setState({ openPanel: false });
  }

  /**
  * @function
  * Intercepts the iframe onedrive messages
  *
  */
  private handleIframeData(element?: any) {
    if (this.state.openPanel != true)
      return;
    var data: string = element.data;
    var indexOfPicker = data.indexOf("[OneDrive-FromPicker]");
    if (indexOfPicker != -1) {
      var message = data.replace("[OneDrive-FromPicker]", "");
      var messageObject = JSON.parse(message);
      if (messageObject.type == "cancel") {
        this.onClosePanel();
      } else if (messageObject.type == "success") {
        var documentUrl: string = this.props.context.pageContext.web.absoluteUrl;
        documentUrl += '/_layouts/15/Doc.aspx?sourcedoc=%7B'+messageObject.items[0].sharePoint.uniqueId+'%7D&action=embedview';
        this.saveDocumentProperty(documentUrl);
        this.onClosePanel();
      }
    }
  }

  /**
  * @function
  * When component is mount, attach the iframe event watcher
  *
  */
  public componentDidMount() {
    window.addEventListener('message', this.handleIframeData, false);
  }

  /**
  * @function
  * Releases the watcher
  *
  */
  public componentWillUnmount() {
    window.removeEventListener('message', this.handleIframeData, false);
    if (this.async !== undefined)
      this.async.dispose();
  }

  /**
   * @function
   * Renders the datepicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {

    var iframeUrl = this.props.context.pageContext.web.absoluteUrl;
    iframeUrl += '/_layouts/15/onedrive.aspx?picker=';
    iframeUrl += '%7B%22sn%22%3Afalse%2C%22v%22%3A%22files%22%2C%22id%22%3A%221%22%2C%22o%22%3A%22';
    iframeUrl += encodeURI(this.props.context.pageContext.web.absoluteUrl.replace(this.props.context.pageContext.web.serverRelativeUrl, ""));
    iframeUrl += "%22%7D&id=";
    iframeUrl += encodeURI(this.props.context.pageContext.web.serverRelativeUrl);
    iframeUrl += '&view=2&typeFilters=';
    iframeUrl += encodeURI('folder,' + this.props.allowedFileExtensions);
    iframeUrl += '&p=2';

    //Renders content
    return (
      <div style={{ marginBottom: '8px' }}>
        <Label>{this.props.label}</Label>
        <table style={{ width: '100%', borderSpacing: 0 }}>
          <tbody>
            <tr>
              <td style={{ width: "*" }}>
                <TextField
                  disabled={this.props.disabled}
                  value={this.state.selectedDocument}
                  style={{ width: '100%' }}
                  onChanged={this.onTextFieldChanged}
                  readOnly={this.props.readOnly}
                />
              </td>
              <td style={{ width: "64px" }}>
                <table style={{ width: '100%', borderSpacing: 0 }}>
                  <tbody>
                    <tr>
                      <td><IconButton
                        disabled={this.props.disabled} iconProps={{ iconName: 'FolderSearch' }} onClick={this.onOpenPanel} /></td>
                      <td><IconButton
                        disabled={this.props.disabled === false && (this.state.selectedDocument != null && this.state.selectedDocument != '') ? false : true}
                        iconProps={{ iconName: 'Delete' }} onClick={this.onEraseButton} /></td>
                    </tr>
                  </tbody>
                </table>
              </td>
            </tr>
          </tbody>
        </table>

        {this.state.errorMessage != null && this.state.errorMessage != '' && this.state.errorMessage != undefined ?
          <div><div aria-live='assertive' className='ms-u-screenReaderOnly' data-automation-id='error-message'>{this.state.errorMessage}</div>
            <span>
              <p className='ms-TextField-errorMessage ms-u-slideDownIn20'>{this.state.errorMessage}</p>
            </span>
          </div>
          : ''}

        {this.state.openPanel === true ?
          <Panel
            isOpen={this.state.openPanel} hasCloseButton={true} onDismiss={this.onClosePanel}
            isLightDismiss={true} type={PanelType.large}
            headerText={strings.DocumentPickerTitle}>

            <iframe ref="filePickerIFrame" style={{ border: "0", width: "100%", height: "80vh"}}
              title="Select files from site picker view. Use toolbaar menu to perform operations, breadcrumbs to navigate between folders and arrow keys to navigate within the list"
              src={iframeUrl}>Loading...</iframe>
          </Panel>
          : ''}

      </div>
    );
  }

}
