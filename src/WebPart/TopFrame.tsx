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

  const init = (ctx: Visio.RequestContext) => {

    ctx.document.application.showToolbars = !props.hideToolbars;
    ctx.document.application.showBorders = !props.hideBorders;

    ctx.document.view.hideDiagramBoundary = props.hideDiagramBoundary;
    ctx.document.view.disableHyperlinks = props.disableHyperlinks;
    ctx.document.view.disablePan = props.disablePan;
    ctx.document.view.disablePanZoomWindow = props.disablePanZoomWindow;
    ctx.document.view.disableZoom = props.disableZoom;

    if (props.startPage)
      ctx.document.setActivePage(props.startPage);

    // sample selection changed event
    ctx.document.onSelectionChanged.add(async (args) => {
      if (args.shapeNames.length === 1) {
        const page = ctx.document.pages.getItem(args.pageName);
        const shape = page.shapes.getItem(args.shapeNames[0]);
        ctx.load(shape, ['hyperlinks']);
        await ctx.sync();
        const link = shape.hyperlinks.items.map(x => `${x.address}`)[0];
        if (link) {
          alert(`navigating to: ${link}`);
          const embedUrl = await resolveUrl(link);
          setEmbedUrl(embedUrl);
        }
      }
    });

    return ctx.sync();
  };

  const [propsChanged, setPropsChanged] = React.useState(0);

  React.useEffect(() => {
    const timer = setTimeout(() => setPropsChanged(propsChanged + 1), 1000);
    return () => clearTimeout(timer);
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

      session.init().then(() => Visio.run(session, ctx => init(ctx)));

      return () => {
        root.innerHTML = "";
      };
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

  const rootStyle = {
    height: props.height ?? "65vh",
    width: props.width ?? "100%",
  };

  return (
    <div className={styles.root} style={rootStyle} ref={ref} />
  );
}
