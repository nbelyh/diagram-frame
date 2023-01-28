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
  isPropertyPaneOpen: boolean;
  isReadOnly: boolean;
  isTeams: boolean;
  onConfigure: () => void;
}

export function TopFrame(props: ITopFrameProps) {

  const refContainer = React.useRef<HTMLDivElement>(null);
  const refUrl = React.useRef('');

  const refDefaultPageName = React.useRef({});

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
          await reloadEmbed(target);
        }
      }
    } catch (err) {
      console.error(err);
    }
  };

  const doSetPage = async (startPage: string) => {
    await Visio.run(refSession.current, async ctx => {
      ctx.document.setActivePage(startPage);
      await ctx.sync();
    })
  }

  const setPage = async (startPage: string) => {
    try {
      await Utils.doWithRetry(() => doSetPage(startPage));
    } catch (err) {
      throw new Error(`Unable to set page to ${startPage}. The current page may not be the expected one. ${err.message}`);
    }
  }

  const doInit = async (url: string) => {
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

      const page = ctx.document.getActivePage().load('name');

      await ctx.sync();

      refDefaultPageName.current[url] = page.name;
    });
  };

  const init = async (url: string) => {
    try {
      await Utils.doWithRetry(() => doInit(url))
    } catch (err) {
      throw new Error(`Error initializing diagram parameters. The view may be not the expected one. ${err.message}`)
    }
  }

  const reloadEmbed = async (opts: { url: string, label?: string }) => {

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

        console.debug(`start new Visio session ${newBaseUrl}`);
        refSession.current = new OfficeExtension.EmbeddedSession(resolved, {
          container: refContainer.current,
          height: '100%',
          width: '100%'
        });

        await refSession.current.init();
        await init(newBaseUrl);
        reloaded = true;
      }

      const oldPageNameOrDefault = oldPageName || refDefaultPageName.current[oldBaseUrl];
      const newPageNameOrDefault = newPageName || refDefaultPageName.current[newBaseUrl];

      if (newPageNameOrDefault && (oldPageNameOrDefault !== newPageNameOrDefault || force)) {
        if (reloaded) { // Visio bug (hanging) on immediate page change with logo screen, timeout seems to help a bit
          setTimeout(() => setPage(newPageNameOrDefault), 750);
        } else {
          await setPage(newPageNameOrDefault);
        }
      }

      if (opts?.label && props.enableNavigation || force) {
        setBreadcrumb(oldBreadcrumb => {
          const foundIndex = oldBreadcrumb.findIndex(x => x.key === opts.url);
          const newBreadcrumb = [...oldBreadcrumb];
          if (foundIndex >= 0) {
            newBreadcrumb.splice(foundIndex);
          }
          newBreadcrumb.push({ key: opts.url, text: opts.label, onClick: () => reloadEmbed({ url: opts.url, label: opts.label }) });
          return newBreadcrumb;
        });
      }

      refUrl.current = opts.url;

    } catch (err) {
      setError(`${err}`);
    }

  }

  const [reloadTrigger, setReloadTrigger] = React.useState(0);

  React.useEffect(() => {
    if (refSession.current) {
      const timer = setTimeout(() => setReloadTrigger(old => old + 1), 1000);
      return () => clearTimeout(timer);
    }
  }, [
    props.height, props.width, props.zoom, props.startPage,
    props.hideToolbars, props.hideBorders, props.hideDiagramBoundary,
    props.disablePan, props.disableZoom, props.disablePanZoomWindow, props.disableHyperlinks
  ]);

  React.useEffect(() => {
    if (refSession.current) {
      reloadEmbed({ url: refUrl.current });
    }
  }, [reloadTrigger]);

  React.useEffect(() => {
    setBreadcrumb([]);
    if (props.url) {
      const url = Utils.joinPageUrl(props.url, props.startPage);
      reloadEmbed({ url, label: strings.NavigationHome });
    }
  }, [props.url, props.enableNavigation]);

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
