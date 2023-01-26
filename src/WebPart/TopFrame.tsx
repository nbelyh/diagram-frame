import * as React from 'react';
import styles from './TopFrame.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IWebPartProps } from './WebPart';
import { SPHttpClient } from '@microsoft/sp-http';
import { Placeholder } from '../min-sp-controls-react/controls/placeholder';
import { MessageBar, MessageBarType, ThemeProvider } from '@fluentui/react';
import * as strings from 'WebPartStrings';

interface ITopFrameProps extends IWebPartProps {
  context: WebPartContext;
  isPropertyPaneOpen: boolean;
  isReadOnly: boolean;
  isTeams: boolean;
  onConfigure: () => void;
}

export function TopFrame(props: ITopFrameProps) {

  const ref = React.useRef(null);
  const [embedUrl, setEmbedUrl] = React.useState(null);

  const enablePropsChanged = React.useRef(false);

  const init = async (session: OfficeExtension.EmbeddedSession, ctx: Visio.RequestContext) => {

    ctx.document.application.showToolbars = !props.hideToolbars;
    ctx.document.application.showBorders = !props.hideBorders;

    ctx.document.view.hideDiagramBoundary = props.hideDiagramBoundary;
    ctx.document.view.disableHyperlinks = props.disableHyperlinks;
    ctx.document.view.disablePan = props.disablePan;
    ctx.document.view.disablePanZoomWindow = props.disablePanZoomWindow;
    ctx.document.view.disableZoom = props.disableZoom;

    const result = await ctx.sync();

    if (props.startPage) {
      setTimeout(() => {
        Visio.run(session, (ctxPage) => {
          ctxPage.document.setActivePage(props.startPage);
          return ctxPage.sync();
        });
      }, 750);
    }

    enablePropsChanged.current = true;

    return result;
  };

  const [propsChanged, setPropsChanged] = React.useState(0);

  React.useEffect(() => {
    if (enablePropsChanged.current) {
      const timer = setTimeout(() => setPropsChanged(propsChanged + 1), 1000);
      return () => clearTimeout(timer);
    }
  }, [
    props.height, props.width,
    props.zoom, props.startPage,
    props.hideToolbars, props.hideBorders, props.hideDiagramBoundary,
    props.disablePan, props.disableZoom, props.disablePanZoomWindow, props.disableHyperlinks
  ]);

  React.useEffect(() => {

    if (embedUrl) {
      const root: HTMLElement = ref.current;

      let url = embedUrl;
      if (props.zoom)
        url = url + `&wdzoom=${props.zoom}`;

      const session: any = new OfficeExtension.EmbeddedSession(url, {
        container: root,
        height: '100%',
        width: '100%'
      });

      session.init().then(() => Visio.run(session, (ctx) => init(session, ctx)));

      return () => {
        root.innerHTML = '';
      };
    } else {
      enablePropsChanged.current = true;
    }

  }, [embedUrl, propsChanged]);

  const resolveUrl = async (fileUrl: string): Promise<string> => {

    if (fileUrl) {
      const apiUrl = `${props.context.pageContext.web.absoluteUrl}/_api/SP.RemoteWeb(@a1)/Web/GetFileByUrl(@a1)/ListItemAllFields/GetWopiFrameUrl(0)?@a1='${encodeURIComponent(fileUrl)}'`;
      const oneDriveWopiFrameResult = await props.context.spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, {
        headers: {
          'accept': 'application/json;odata=nometadata',
          'content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      });

      if (!oneDriveWopiFrameResult || !oneDriveWopiFrameResult.ok) {
        if (oneDriveWopiFrameResult.status === 404) {
          throw new Error(`${strings.messageVisioFileNotFound} ${fileUrl} ${strings.messageWasTheFileDeleted}`);
        }
        if (oneDriveWopiFrameResult.status === 403) {
          throw new Error(`${strings.messageVisioFileCannotAccessed} ${fileUrl} ${strings.messageArePermissionsMissing}`);
        }
        throw new Error(`${strings.messageSomethingWentWrongResolveURL}  ${fileUrl}. Status='${oneDriveWopiFrameResult.status}'`);
      }

      const oneDriveWopiFrameData = await oneDriveWopiFrameResult.json();
      if (!oneDriveWopiFrameData || !oneDriveWopiFrameData.value) {
        throw new Error(`${strings.messageCannotResolveFileURL} ${fileUrl}`);
      }

      const result = oneDriveWopiFrameData.value
        .replace('action=view', 'action=embedview')
        .replace('action=default', 'action=embedview');

      return result;
    }
  };

  const [loadError, setLoadError] = React.useState('');

  React.useEffect(() => {
    setLoadError('');
    props.context.statusRenderer.displayLoadingIndicator(ref.current, 'diagram');
    resolveUrl(props.url).then(val => {
      props.context.statusRenderer.clearLoadingIndicator(ref.current);
      setEmbedUrl(val);
    }, err => {
      props.context.statusRenderer.clearLoadingIndicator(ref.current);
      setLoadError(`${err}`);
      // props.context.statusRenderer.renderError(ref.current, err);
    });
  }, [props.url]);

  const rootStyle = {
    height: props.height,
    width: props.width,
  };

  const showPlaceholder = !props.url || loadError;

  const placeholderIconName = loadError
    ? strings.Error
    : strings.Edit;

  const placeholderIconText = loadError
    ? strings.placeholderIconTextUnableShowVisio
    : strings.placeholderIconTextVisioNotSelected

  const placeholderDescription = props.isPropertyPaneOpen
    ? strings.placeholderIconTextPleaseclickBrowse
    : props.isReadOnly
      ? (props.isTeams
        ? strings.placeholderIconTextPleaseclickSettings
        : strings.placeholderIconTextPleaseclickEdit
      )
      : strings.placeholderIconTextPleaseclickConfigure

  return (
    <ThemeProvider className={styles.root} style={rootStyle} ref={ref}>
      {loadError && <MessageBar messageBarType={MessageBarType.error}>{loadError}</MessageBar>}
      {showPlaceholder && <Placeholder
        iconName={placeholderIconName}
        iconText={placeholderIconText}
        description={placeholderDescription}
        buttonLabel={strings.FieldConfigureLabel}
        onConfigure={() => props.onConfigure()}
        hideButton={props.isReadOnly}
        disableButton={props.isPropertyPaneOpen}
      />}
    </ThemeProvider>
  );
}
