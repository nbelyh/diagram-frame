import * as React from 'react';
import styles from './TopFrame.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/files";

export function TopFrame(props: {
  url: string;
  width: string;
  height: string;
  showToolbars: boolean;
  showBorders: boolean;
  zoom: number;
  context: WebPartContext
}) {

  const ref = React.useRef(null);
  const [embedUrl, setEmbedUrl] = React.useState(null);

  const init = (ctx: Visio.RequestContext) => {
    ctx.document.application.showToolbars = props.showToolbars;
    ctx.document.application.showBorders = props.showBorders;

    // if (props.zoom) {
    //   var activePage = ctx.document.getActivePage();
    //   activePage.view.zoom = props.zoom;
    // }

    return ctx.sync();
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

      return () => root.innerHTML = "";
    }
  }, [embedUrl, props.height, props.width, props.showToolbars, props.showBorders, props.zoom]);

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
    display: "flex"
  };

  return (
    <div className={styles.root} style={rootStyle} >
      {embedUrl && <div style={{ flex: 1 }} ref={ref} />}
    </div>
  );
}
