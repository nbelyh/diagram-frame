import * as React from 'react';
import styles from './TopFrame.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IWebPartProps } from './WebPart';

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/files";

interface ITopFrameProps extends IWebPartProps {
  context: WebPartContext;
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

  React.useEffect(() => {
    props.context.statusRenderer.displayLoadingIndicator(ref.current, 'diagram');
    resolveUrl(props.url).then(val => {
      props.context.statusRenderer.clearLoadingIndicator(ref.current);
      setEmbedUrl(val);
    }, err => {
      props.context.statusRenderer.renderError(ref.current, err);
    });
  }, [props.url]);

  const defaultWidth = '100%';
  const defaultHeight = (props.context.sdks?.microsoftTeams || !props.context.pageContext?.listItem?.id) ? '100%' : '65vh';

  const rootStyle = {
    height: props.height ? props.height : defaultHeight,
    width: props.width ? props.width : defaultWidth,
    overflow: 'hidden'
  };

  return (
    <div className={styles.root} style={rootStyle} ref={ref} />
  );
}
