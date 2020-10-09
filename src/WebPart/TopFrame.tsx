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

  const init = async (ctx: Visio.RequestContext) => {

    ctx.document.application.showToolbars = !props.hideToolbars;
    ctx.document.application.showBorders = !props.hideBorders;

    ctx.document.view.hideDiagramBoundary = props.hideDiagramBoundary;
    ctx.document.view.disableHyperlinks = props.disableHyperlinks;
    ctx.document.view.disablePan = props.disablePan;
    ctx.document.view.disablePanZoomWindow = props.disablePanZoomWindow;
    ctx.document.view.disableZoom = props.disableZoom;

    const pages = ctx.document.pages.load();

    await ctx.sync();

    for (let i = 0; i < pages.items.length; ++i) {
      if (pages.items[i].name === props.startPage || (i + 1) === +props.startPage)
        ctx.document.setActivePage(pages.items[i].name);
    }

    await ctx.sync();
  };

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

      session.init().then(() => Visio.run(session, ctx => init(ctx)));

      return () => {
        root.innerHTML = "";
      };
    }
  }, [embedUrl,
    props.height, props.width,
    props.zoom, props.startPage,
    props.hideToolbars, props.hideBorders, props.hideDiagramBoundary,
    props.disablePan, props.disableZoom, props.disablePanZoomWindow, props.disableHyperlinks,
  ]);

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
    resolveUrl(props.url).then(val => setEmbedUrl(val));
  }, [props.url]);

  const rootStyle = {
    height: props.height,
    width: props.width,
  };

  return (
    <div className={styles.root} style={rootStyle} ref={ref} />
  );
}
