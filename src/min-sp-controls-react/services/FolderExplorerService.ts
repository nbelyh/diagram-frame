// import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
// import { PageContext } from "@microsoft/sp-page-context";
import { IFolderExplorerService } from "./IFolderExplorerService";
import { IFolder } from "./IFolderExplorerService";
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { sp } from '@pnp/sp';
import '@pnp/sp/webs';

import { Logger, LogLevel } from '@pnp/logging';
import { HttpRequestError } from '@pnp/odata';
import { hOP } from '@pnp/common';

export async function handleError(e: Error | HttpRequestError): Promise<void> {

  if (hOP(e, 'isHttpRequestError')) {

    // we can read the json from the response
    const data = await (<HttpRequestError>e).response.json();

    // parse this however you want
    const message = typeof data['odata.error'] === 'object' ? data['odata.error'].message.value : e.message;

    // we use the status to determine a custom logging level
    const level: LogLevel = (<HttpRequestError>e).status === 404 ? LogLevel.Warning : LogLevel.Info;

    // create a custom log entry
    Logger.log({
      data,
      level,
      message,
    });

    throw message;

  } else {
    // not an HttpRequestError so we just log message
    console.error(e);
  }
}

export class FolderExplorerService implements IFolderExplorerService {

  protected context: BaseComponentContext;

  constructor(context: BaseComponentContext) {
    this.context = context;
  }

  private getClient(webAbsoluteUrl: string) {
    return sp.createIsolated({ baseUrl: webAbsoluteUrl, cloneGlobal: true });
  }

  /**
   * Get libraries within a given site
   * @param webAbsoluteUrl - the url of the target site
   */
  public GetDocumentLibraries = async (webAbsoluteUrl: string): Promise<IFolder[]> => {
    return this._getDocumentLibraries(webAbsoluteUrl);
  }

  /**
   * Get libraries within a given site
   * @param webAbsoluteUrl - the url of the target site
   */
  private _getDocumentLibraries = async (webAbsoluteUrl: string): Promise<IFolder[]> => {
    let results: IFolder[] = [];
    try {
      const sp = await this.getClient(webAbsoluteUrl);
      const libraries: any[] = await sp.web.lists.filter('BaseTemplate eq 101 and Hidden eq false').expand('RootFolder').select('Title', 'RootFolder/ServerRelativeUrl').orderBy('Title')();

      results = libraries.map((library): IFolder => {
        return { Name: library.Title, ServerRelativeUrl: library.RootFolder.ServerRelativeUrl };
      });
    } catch (error) {
      console.error('Error loading folders', error);
    }
    return results;

  }

  /**
 * Get folders within a given library or sub folder
 * @param webAbsoluteUrl - the url of the target site
 * @param folderRelativeUrl - the relative url of the folder
 */
  public GetFolders = async (webAbsoluteUrl: string, folderRelativeUrl: string, orderby: string, orderAscending: boolean): Promise<IFolder[]> => {
    return this._getFolders(webAbsoluteUrl, folderRelativeUrl, orderby, orderAscending);
  }

  /**
   * Get folders within a given library or sub folder
   * @param webAbsoluteUrl - the url of the target site
   * @param folderRelativeUrl - the relative url of the folder
   */
  private _getFolders = async (webAbsoluteUrl: string, folderRelativeUrl: string, orderby: string, orderAscending: boolean): Promise<IFolder[]> => {
    let results: IFolder[] = [];
    try {
      const sp = await this.getClient(webAbsoluteUrl);
      folderRelativeUrl = folderRelativeUrl.replace(/\'/ig, "''");
      let foldersResult: IFolder[] = await sp.web.getFolderByServerRelativePath(folderRelativeUrl).folders.select('Name', 'ServerRelativeUrl').orderBy(orderby, orderAscending)();
      results = foldersResult.filter(f => f.Name != "Forms");
    } catch (error) {
      console.error('Error loading folders', error);
    }
    return results;
  }

  /**
   * Create a new folder
   * @param webAbsoluteUrl - the url of the target site
   * @param folderRelativeUrl - the relative url of the base folder
   * @param name - the name of the folder to be created
   */
  public AddFolder = async (webAbsoluteUrl: string, folderRelativeUrl: string, name: string): Promise<IFolder> => {
    return this._addFolder(webAbsoluteUrl, folderRelativeUrl, name);
  }

  /**
 * Create a new folder
 * @param webAbsoluteUrl - the url of the target site
 * @param folderRelativeUrl - the relative url of the base folder
 * @param name - the name of the folder to be created
 */
  private _addFolder = async (webAbsoluteUrl: string, folderRelativeUrl: string, name: string): Promise<IFolder> => {
    let folder: IFolder = null;
    try {
      const sp = await this.getClient(webAbsoluteUrl);
      folderRelativeUrl = folderRelativeUrl.replace(/\'/ig, "''");
      let folderAddResult = await sp.web.getFolderByServerRelativePath(folderRelativeUrl).folders.addUsingPath(name);
      if (folderAddResult && folderAddResult.data) {
        folder = {
          Name: folderAddResult.data.Name,
          ServerRelativeUrl: folderAddResult.data.ServerRelativeUrl
        };
      }
    } catch (error) {
      await handleError(error);
    }
    return folder;
  }

}
