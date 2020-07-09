import * as React from 'react';

export function VisioOnlineSpfxWebPart(props: {
  url: string;
  width: string;
  height: string;
}) {

  const ref = React.useRef(null);

  React.useEffect(() => {

    const root = ref.current;
    if (props.url && props.url.indexOf("Doc.aspx") >= 0) {
      const session: any = new OfficeExtension.EmbeddedSession(props.url, {
        container: root,
        width: props.width,
        height: props.height,
      });

      session.init().then(() => {
        return Visio.run(session, ctx => {
          ctx.document.application.showToolbars = false;
          return ctx.sync();
        });
      });

      return () => {
        for (var i = 0; i < root.childNodes.length; ++i) {
          root.removeChild(root.childNodes[i]);
        }
      };
    }
  }, [props.url]);

  return (<div style={{ height: props.height, width: props.width }} ref={ref} ></div>);
}
