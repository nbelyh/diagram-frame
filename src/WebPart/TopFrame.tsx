import * as React from 'react';
import styles from './TopFrame.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IWebPartProps } from "./IWebPartProps";
import { Utils } from './Utils';
import { ErrorPlaceholder } from './components/ErrorPlaceholder';
import { Breadcrumb, IBreadcrumbItem, MessageBar, MessageBarType, ThemeProvider } from '@fluentui/react';
import * as strings from 'WebPartStrings';

interface ITopFrameProps extends IWebPartProps {
  context: WebPartContext;
  isReadOnly: boolean;
}

export function TopFrame(props: ITopFrameProps) {

  const refContainer = React.useRef<HTMLDivElement>(null);
  const refUrl = React.useRef('');

  const refSession = React.useRef<OfficeExtension.EmbeddedSession>(null);

  const getVisioLink = async (args: Visio.SelectionChangedEventArgs) => {
    return await Visio.run(refSession.current, async (ctx) => {
      const [shapeName] = args.shapeNames;
      if (args.pageName && shapeName) {
        try {
          const links = ctx.document.pages.getItem(args.pageName).shapes.getItem(shapeName).hyperlinks.load();
          await ctx.sync();
          return links.items[0];
        } catch (err) {
          console.error(err);
        }
      }
    });
  }

  const deselectVisioShape = async (args: Visio.SelectionChangedEventArgs) => {
    try {
      await Visio.run(refSession.current, async (ctx) => {
        const [shapeName] = args.shapeNames;
        ctx.document.pages.getItem(args.pageName).shapes.getItem(shapeName).set({ select: false });
        await ctx.sync();
      });
    } catch (err) {
      console.error(err);
    }
  }

  const onVisioSelectionChanged = async (args: Visio.SelectionChangedEventArgs) => {
    try {
      const { baseUrl } = Utils.splitPageUrl(refUrl.current);
      const [shapeName] = args.shapeNames;

      const link = await getVisioLink(args);
      if (link) {
        const target = Utils.getVisioLinkTarget(link, baseUrl, shapeName);
        if (target) {
          await deselectVisioShape(args);
          await reloadEmbed({...target, retry: 0 });
        }
      }
    } catch (err) {
      console.error(err);
    }
  };

  const setPage = async (startPage: string) => {
    await Visio.run(refSession.current, async ctx => {
      console.log(`[DiagramFrame] set page ${startPage}`);
      ctx.document.setActivePage(startPage);
      await ctx.sync();
    })
  }

  const getPage = async () => {
    return await Visio.run(refSession.current, async ctx => {
      const page = ctx.document.getActivePage().load('name');
      await ctx.sync();
      console.log(`[DiagramFrame] get page: ${page.name}`);
      return page.name;
    })
  }

  const doInit = async (url: string, startPage: string) => {
    await Visio.run(refSession.current, async (ctx) => {
      ctx.document.application.showToolbars = !props.hideToolbars;
      ctx.document.application.showBorders = !props.hideBorders;

      ctx.document.view.hideDiagramBoundary = props.hideDiagramBoundary;
      ctx.document.view.disableHyperlinks = props.disableHyperlinks;
      ctx.document.view.disablePan = props.disablePan;
      ctx.document.view.disablePanZoomWindow = props.disablePanZoomWindow;
      ctx.document.view.disableZoom = props.disableZoom;

      if (props.enableNavigation) {
        ctx.document.onSelectionChanged.add(onVisioSelectionChanged);
      }

      if (startPage) {
        ctx.document.setActivePage(startPage);
      }

      await ctx.sync();
    });
  };

  const init = async (url: string, pageName: string) => {
    try {
      await Utils.doWithRetry(() => doInit(url, pageName))
    } catch (err) {
      throw new Error(`Error initializing diagram parameters. The view may be not the expected one. ${err.message}`)
    }
  }

  const udpateBreadcrumb = (opts: { url: string, label?: string }) => {
    setBreadcrumb(oldBreadcrumb => {
      const foundIndex = oldBreadcrumb.findIndex(x => x.key === opts.url);
      const newBreadcrumb = [...oldBreadcrumb];
      if (foundIndex >= 0) {
        newBreadcrumb.splice(foundIndex);
      }
      newBreadcrumb.push({ key: opts.url, text: opts.label, onClick: () => reloadEmbed({ url: opts.url, label: opts.label, retry: 0 }) });
      return newBreadcrumb;
    });
  }

  const reloadEmbed = async (opts: { url: string, label: string, retry: number }) => {

    setError('');
    try {

      // if clicked the same URL, force reload
      const force = refUrl.current === opts.url;

      const { baseUrl: oldBaseUrl, pageName: oldPageName } = Utils.splitPageUrl(refUrl.current);
      const { baseUrl: newBaseUrl, pageName: newPageName } = Utils.splitPageUrl(opts.url);

      let reloaded = false;
      if (newBaseUrl && (oldBaseUrl !== newBaseUrl || force)) {

        let resolved = await Utils.resolveUrl(props.context, newBaseUrl);

        if (props.zoom)
          resolved = resolved + `&wdzoom=${props.zoom}`;

        refContainer.current.innerHTML = '';

        refSession.current = null;

        console.log(`[DiagramFrame] open file ${newBaseUrl}`);
        refSession.current = new OfficeExtension.EmbeddedSession(resolved, {
          container: refContainer.current,
          height: '100%',
          width: '100%'
        });

        await refSession.current.init();
        await init(newBaseUrl, newPageName);
        reloaded = true;
      }

      if (newPageName && (oldPageName !== newPageName || force)) {
        if (reloaded) { // Visio bug (hanging) on immediate page change with logo screen, timeout seems to help a bit
          // await new Promise(r => setTimeout(r, 750));
          // await setPage(newPageNameOrDefault);
          await new Promise(r => setTimeout(r, 750));
          const pageName = await getPage();
          if (pageName !== newPageName) {
            if (opts.retry < 3) {
              await reloadEmbed({...opts, retry: opts.retry + 1});
            }
          }
        } else {
          await setPage(newPageName);
        }
      }

      if (opts?.label && props.enableNavigation || force) {
        udpateBreadcrumb(opts);
      }

      refUrl.current = opts.url;

    } catch (err) {
      setError(`${err}`);
    }

  }

  React.useEffect(() => {
    if (refSession.current) {
      const timer = setTimeout(() => {
        reloadEmbed({ url: refUrl.current, label: undefined, retry: 0 })
      }, 750);
      return () => clearTimeout(timer);
    }
  }, [
    props.height, props.width, props.zoom,
    props.hideToolbars, props.hideBorders, props.hideDiagramBoundary,
    props.disablePan, props.disableZoom, props.disablePanZoomWindow, props.disableHyperlinks
  ]);

  React.useEffect(() => {
    const timer = setTimeout(() => {
      setBreadcrumb([]);
      if (props.url) {
        const opts = { url: Utils.joinPageUrl(props.url, props.startPage), label: strings.BreadcrumbStart, retry: 0 };
        udpateBreadcrumb({...opts, label: strings.BreadcrumbLoading });
        reloadEmbed(opts);
      }
    }, 750);
    return () => clearTimeout(timer);
  }, [props.url, props.startPage, props.enableNavigation]);

  const [error, setError] = React.useState('');

  const [breadcrumb, setBreadcrumb] = React.useState<IBreadcrumbItem[]>([]);

  return (
    <ThemeProvider className={styles.root} style={{ height: props.height, width: props.width }} >
      {props.enableNavigation && <Breadcrumb styles={{ root: { margin: 4 } }} items={breadcrumb} />}
      {error && <MessageBar onDismiss={() => setError('')} messageBarType={MessageBarType.severeWarning}>{error}</MessageBar>}
      {!props.url && <ErrorPlaceholder context={props.context} isReadOnly={props.isReadOnly} />}
      <div className={styles.diagram} ref={refContainer} />
    </ThemeProvider>
  );
}
