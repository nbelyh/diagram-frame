import * as React from 'react';
import styles from './TopFrame.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IWebPartProps } from './WebPart';

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/files";
import { Placeholder } from '../min-sp-controls-react/controls/placeholder';

interface ITopFrameProps extends IWebPartProps {
  context: WebPartContext;
  isPropertyPaneOpen: boolean;
  isReadOnly: boolean;
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
        root.innerHTML = "";
      };
    } else {
      enablePropsChanged.current = true;
    }

  }, [embedUrl, propsChanged]);

  const resolveUrl = async (url: string) => {
    if (url) {
      const file = sp.web.getFileByUrl(url);
      const item = await file.getItem();

      const wopiFrameUrl = await item.getWopiFrameUrl(0);
      const result = wopiFrameUrl.replace("action=view", "action=embedview");
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
      setLoadError(err);
      // props.context.statusRenderer.renderError(ref.current, err);
    });
  }, [props.url]);

  const rootStyle = {
    height: props.height,
    width: props.width,
  };

  return (
    <div className={styles.root} style={rootStyle} ref={ref}>
      {!props.url && <Placeholder
        iconName="Edit"
        iconText={"Configure Web Part"}
        description={props.isPropertyPaneOpen
          ? "Click 'Browse...' Button on configuration panel to select the diagram"
          : props.isReadOnly
            ? `Click 'Edit' to start page editing to reconfigure this web part`
            : `Click 'Configure' button to configure the web part`}
        buttonLabel={"Configure"}
        onConfigure={() => props.onConfigure()}
        hideButton={props.isReadOnly}
        disableButton={props.isPropertyPaneOpen}
      />}
      {!!loadError && <Placeholder
        iconName="Error"
        iconText={"Unable to show this Visio diagram"}
        description={props.isPropertyPaneOpen
          ? `${loadError} Click 'Browse...' Button on configuration panel to select other diagram. Unable to display: ${props.url}`
          : props.isReadOnly
            ? `${loadError} Click 'Edit' to start page editing to reconfigure this web part. Unable to display: ${props.url}`
            : `${loadError} Click 'Configure' button to reconfigure this web part. Unable to display: ${props.url}`}
        buttonLabel={"Configure"}
        onConfigure={() => props.onConfigure()}
        hideButton={props.isReadOnly}
        disableButton={props.isPropertyPaneOpen}
      />}
    </div>);
}
