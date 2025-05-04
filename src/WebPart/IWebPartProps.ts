
export interface IWebPartProps {
  url: string;
  startPage: string;
  width: string;
  height: string;
  hideToolbars: boolean;
  hideBorders: boolean;

  hideDiagramBoundary: boolean;
  disableHyperlinks: boolean;
  disablePan: boolean;
  disablePanZoomWindow: boolean;
  disableZoom: boolean;

  enableNavigation: boolean;
  enableNavigationHeader: boolean;

  openHyperlinksInNewWindow: boolean;
  forceOpeningOfficeFilesOnline: boolean;

  zoom: number;
}
