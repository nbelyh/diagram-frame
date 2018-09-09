import * as React from 'react';
import { IVisioOnlineSpfxWebPartProps } from './IVisioOnlineSpfxWebPartProps';
import styles from './VisioOnlineSpfxWebPart.module.scss';

export default class VisioOnlineSpfxWebPart extends React.Component<IVisioOnlineSpfxWebPartProps, {}> {

  private refreshVisioFrame() {

    var root = this.refs.frame as HTMLElement;

    for (var i = 0; i < root.childNodes.length; ++i) {
      root.removeChild(root.childNodes[i]);
    }

    const url = this.props.url;
    if (url) {
      if (url.indexOf("https://") >= 0) {
        const session: any = new OfficeExtension.EmbeddedSession(url, {
          container: root,
          width: this.props.width,
          height: this.props.height,
        });

        session.init().then(() => {

          Visio.run(session, ctx => {
            ctx.document.application.showToolbars = false;
            return ctx.sync();
          });

        });
      }
    }
  }

  public componentDidUpdate(prevProps) {
    if (this.props.url != prevProps.url ||
      this.props.width != prevProps.width ||
      this.props.height != prevProps.height) {
      this.refreshVisioFrame();
    }
  }

  public componentDidMount() {
    this.refreshVisioFrame();
  }

  public render(): React.ReactElement<IVisioOnlineSpfxWebPartProps> {
    return (
      <div className={ styles.VisioOnlineSpfxWebPart } ref="frame">
      </div>
    );
  }
}
