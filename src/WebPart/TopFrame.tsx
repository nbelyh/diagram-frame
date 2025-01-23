import * as React from 'react';
import styles from './TopFrame.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IWebPartProps } from "./IWebPartProps";
import { Utils } from './Utils';
import { ErrorPlaceholder } from './components/ErrorPlaceholder';
import { Breadcrumb, IBreadcrumbItem, MessageBar, MessageBarType, Spinner, Stack, ThemeProvider } from '@fluentui/react';
import * as strings from 'WebPartStrings';

interface ITopFrameProps extends IWebPartProps {
  context: WebPartContext;
  isReadOnly: boolean;
}

const sleep = async (ms: number) => await new Promise(r => setTimeout(r, ms));

export function TopFrame(props: ITopFrameProps) {

  const refContainer = React.useRef<HTMLDivElement>(null);
  const refUrl = React.useRef('');

  const refSession = React.useRef<OfficeExtension.EmbeddedSession>(null);

  const refDefaultPageName = React.useRef({});

  const getVisioLink = async (args: Visio.SelectionChangedEventArgs) => {
    return await Visio.run(refSession.current, async (ctx) => {
      const [shapeName] = args.shapeNames;
      if (args.pageName && shapeName) {
        try {
          const links = ctx.document.pages.getItem(args.pageName).shapes.getItem(shapeName).hyperlinks.load();
          await ctx.sync();
          return links.items[0];
        } catch (err) {
          console.warn(`[DiagramFrame] unable to get shape "${shapeName}" on page "${args.pageName}": ${err.message}`);
        }
      }
    });
  }

  const deselectVisioShape = async (args: Visio.SelectionChangedEventArgs) => {
    const [shapeName] = args.shapeNames;
    try {
      await Visio.run(refSession.current, async (ctx) => {
        ctx.document.pages.getItem(args.pageName).shapes.getItem(shapeName).set({ select: false });
        await ctx.sync();
      });
    } catch (err) {
      console.warn(`[DiagramFrame] unable to deselect shape "${shapeName}" ${err.message}`);
    }
  }

  const officeExtensions = new Set([
    'doc', 'docx', 'dot', 'dotx', // Word
    'xls', 'xlsx', 'xlsm', 'xltx', 'xltm',  // Excel
    'ppt', 'pptx', 'pps', 'ppsx', 'pot', 'potx', // PowerPoint
    'pub', // Publisher
    'vsd', 'vsdx', // Visio
    'odt', 'ods', 'odp', // OpenDocument Text/Spreadsheet/Presentation
    'rtf' // Rich Text Format
  ]);

  const isOfficeFileExtension = (url: string) => {
    const extension = url.split('.').pop().toLowerCase().split(/#|\?/)[0];
    return officeExtensions.has(extension);
  }

  const onVisioSelectionChanged = async (args: Visio.SelectionChangedEventArgs) => {
    const [shapeName] = args.shapeNames;
    try {
      const { baseUrl } = Utils.splitPageUrl(refUrl.current);

      const link = await getVisioLink(args);
      if (link) {
        const parsed = Utils.parseLink(link, baseUrl, shapeName);
        if (parsed) {
          if (parsed.external) {

            const fileUrl = new URL(parsed.url);
            if (props.forceOpeningOfficeFilesOnline && isOfficeFileExtension(parsed.url)) {
              fileUrl.searchParams.append('web', '1');
            }

            if (props.openHyperlinksInNewWindow)
              window.open(fileUrl, '_blank');
            else
              document.location = fileUrl.toString();
          } else {
            await deselectVisioShape(args);
            await reloadEmbed({...parsed, retry: 0 });
          }
        }
      }
    } catch (err) {
      console.warn(`[DiagramFrame] unable to navigate to shape "${shapeName}" ${err.message}`);
    }
  };

  const setPage = async (startPage: string) => {
    await Visio.run(refSession.current, async ctx => {
      console.log(`[DiagramFrame] set page "${startPage}"`);
      try {
        ctx.document.setActivePage(startPage);
        await ctx.sync();
      } catch (err) {
        if (err.code === 'ItemNotFound') {
          throw new Error(`Unable to set active page to "${startPage}" because it is not found. Please check you have specified an existing page in the web part settings.`);
        } else {
          throw err;
        }
      }
    })
  }

  const getPage = async () => {
    return await Visio.run(refSession.current, async ctx => {
      const page = ctx.document.getActivePage().load('name');
      await ctx.sync();
      console.log(`[DiagramFrame] get page returned "${page.name}"`);
      return page.name;
    })
  }

  const init = async (url: string, startPage: string, retry: number) => {
    try {
      await Visio.run(refSession.current, async (ctx) => {

        let loaded = false;
        const onVisioDocumentLoaded = async (args: Visio.DocumentLoadCompleteEventArgs) => {
          loaded = true;
          console.log(`[DiagramFrame] document loaded: ${args.success}`);
        }

        ctx.document.onDocumentLoadComplete.add(onVisioDocumentLoaded);

        await ctx.sync();

        // trying to call Visio online API before 'loaded' event results in all sort of odd errors on slow LAN
        // cna be tested by switching "Slow 3G" for example in chrone dev tools.
        for (let i = 0; !loaded && i < 4*10; ++i) {
          await sleep(250);
        }

        if (!loaded) {
          throw new Error('Timeout while waiting for the diagram to load');
        }

        ctx.document.onDocumentLoadComplete.remove(onVisioDocumentLoaded);

        ctx.document.application.showToolbars = !props.hideToolbars;
        ctx.document.application.showBorders = !props.hideBorders;

        ctx.document.view.hideDiagramBoundary = props.hideDiagramBoundary;
        ctx.document.view.disableHyperlinks = props.disableHyperlinks;
        ctx.document.view.disablePan = props.disablePan;
        ctx.document.view.disablePanZoomWindow = props.disablePanZoomWindow;
        ctx.document.view.disableZoom = props.disableZoom;

        const defaultPage = ctx.document.getActivePage().load('name');

        if (props.enableNavigation) {
          ctx.document.onSelectionChanged.add(onVisioSelectionChanged);
        }

        await ctx.sync();

        refDefaultPageName.current[url] = defaultPage.name;

        if (startPage) {
          // A Visio online issue, depends also on the network speed, see
          // https://github.com/OfficeDev/office-js/issues/1539
          await sleep(750 * (1 + retry*2));

          console.log(`[DiagramFrame] initialize page "${startPage}"`);
          try {
            ctx.document.setActivePage(startPage);
            await ctx.sync();
          } catch (err) {
            if (err.code === 'ItemNotFound') {
              throw new Error(`Unable to set active page to "${startPage}" because it is not found. Please check you have specified an existing page in the web part settings.`);
            } else {
              throw err;
            }
          }
        }

      });
    } catch (err) {
      console.error(`[DiagramFrame] error initializing diagram ${err.message}`);
      throw new Error(`Error initializing diagram parameters. The view may be not the expected one. Please try reloading the page. ${err.message}`);
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
    setIsLoading(true);
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

        console.log(`[DiagramFrame] loading "${newBaseUrl}#${newPageName}"`);
        refSession.current = new OfficeExtension.EmbeddedSession(resolved, {
          container: refContainer.current,
          height: '100%',
          width: '100%'
        });

        await refSession.current.init();
        await init(newBaseUrl, newPageName, opts.retry);
        reloaded = true;
      }

      const oldPageNameOrDefault = oldPageName || refDefaultPageName.current[oldBaseUrl];
      const newPageNameOrDefault = newPageName || refDefaultPageName.current[newBaseUrl];

      if (newPageNameOrDefault && (oldPageNameOrDefault !== newPageNameOrDefault || force)) {

        if (reloaded) {
          // Visio bug (hanging) on immediate page change with logo screen, timeout seems to help a bit
          // See https://github.com/OfficeDev/office-js/issues/1539
          // This is a heuristic workaround to eventually set the page

          let pageSet = false;
          for (let i = 0; i  < (opts.retry + 1) * 3; ++i) {

            const pageName = await getPage();
            if (pageName === newPageNameOrDefault) {
              pageSet = true;
              break;
            }

            console.warn(`[DiagramFrame] Page mismatch after ${1+i} seconds, resceduling check`);
            await sleep(1000);
          }

          if (!pageSet) {
            console.warn(`[DiagramFrame] Page mismatch, initiating reload`);
            if (opts.retry < 2) {
              reloadEmbed({...opts,  retry: opts.retry + 1});
              return;
            }

            throw new Error(`Error while loading diagram. The view may be not the expected one. Please try reloading the page.`);
          }

        } else {
          await setPage(newPageNameOrDefault);
        }
      }

      if (opts?.label && props.enableNavigation || force) {
        udpateBreadcrumb(opts);
      }

      refUrl.current = opts.url;

    } catch (err) {
      console.error(`[DiagramFrame] unable to initialize the diagram, ${err.message}`);
      setError(`${err}`);
    } finally {
      setIsLoading(false);
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
  }, [props.url, props.startPage, props.enableNavigation, props.openHyperlinksInNewWindow, props.forceOpeningOfficeFilesOnline]);

  const [error, setError] = React.useState('');
  const [isLoading, setIsLoading] = React.useState(false);

  const [breadcrumb, setBreadcrumb] = React.useState<IBreadcrumbItem[]>([]);

  return (
    <ThemeProvider className={styles.root} style={{ height: props.height, width: props.width }} >
      {props.enableNavigation && <Stack horizontal>
        <Stack.Item grow><Breadcrumb styles={{ root: { margin: 4 } }} items={breadcrumb} /></Stack.Item>
        <Stack.Item align='center'>{isLoading && <Spinner />}</Stack.Item>
      </Stack>
      }
      {error && <MessageBar onDismiss={() => setError('')} messageBarType={MessageBarType.severeWarning}>{error}</MessageBar>}
      {!props.url && <ErrorPlaceholder context={props.context} isReadOnly={props.isReadOnly} />}
      <div className={styles.diagram} ref={refContainer} />
    </ThemeProvider>
  );
}
