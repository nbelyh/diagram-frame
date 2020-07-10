import * as React from 'react';
import styles from './TopFrame.module.scss';

export function TopFrame(props: {
  url: string;
  width: string;
  height: string;
}) {

  const ref = React.useRef(null);

  React.useEffect(() => {

    const root: HTMLElement = ref.current;
    if (props.url && props.url.indexOf("Doc.aspx") >= 0) {
      const session: any = new OfficeExtension.EmbeddedSession(props.url, {
        container: root,
        height: '100%',
        width: '100%'
      });

      session.init().then(() => {
        return Visio.run(session, ctx => {
          ctx.document.application.showToolbars = false;
          return ctx.sync();
        });
      });

      return () => {
        root.innerHTML = "";
      };
    }
  }, [props.url, props.height, props.width]);

  const rootStyle = {
    height: props.height,
    width: props.width,
  };

  return (<div className={styles.root} style={rootStyle} ref={ref} ></div>);
}
