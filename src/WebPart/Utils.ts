import * as strings from 'WebPartStrings';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export class Utils {

  static resolvedUrls = {};

  public static joinPageUrl(baseUrl: string, newPageName: string) {
    return baseUrl + '#' + (newPageName || '');
  }

  public static splitPageUrl(url: string) {
    const [baseUrl, pageName] = url ? url.split('#') : ['', ''];
    return { baseUrl, pageName };
  }

  public static async resolveUrl(context: WebPartContext, fileUrl: string): Promise<string> {

    const resolved = Utils.resolvedUrls[fileUrl];
    if (resolved) {
      return resolved;
    }

    if (fileUrl) {
      const apiUrl = `${context.pageContext.web.absoluteUrl}/_api/SP.RemoteWeb(@a1)/Web/GetFileByUrl(@a1)/ListItemAllFields/GetWopiFrameUrl(0)?@a1='${encodeURIComponent(fileUrl)}'`;
      const oneDriveWopiFrameResult = await context.spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, {
        headers: {
          'accept': 'application/json;odata=nometadata',
          'content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      });

      if (!oneDriveWopiFrameResult || !oneDriveWopiFrameResult.ok) {
        if (oneDriveWopiFrameResult.status === 404) {
          throw new Error(`${strings.messageVisioFileNotFound} ${strings.messageWasTheFileDeleted} ${fileUrl}`);
        }
        if (oneDriveWopiFrameResult.status === 403) {
          throw new Error(`${strings.messageVisioFileCannotAccessed} ${strings.messageArePermissionsMissing} ${fileUrl}`);
        }
        throw new Error(`${strings.messageSomethingWentWrongResolveURL} ${oneDriveWopiFrameResult.statusText} ${fileUrl}`);
      }

      const oneDriveWopiFrameData = await oneDriveWopiFrameResult.json();
      if (!oneDriveWopiFrameData || !oneDriveWopiFrameData.value) {
        throw new Error(`${strings.messageCannotResolveFileURL} ${fileUrl}`);
      }

      const result = oneDriveWopiFrameData.value
        .replace('action=view', 'action=embedview')
        .replace('action=default', 'action=embedview');

      Utils.resolvedUrls[fileUrl] = result;
      return result;
    }
  }

  static isRelativeUrl(fileUrl: string) {
    return !fileUrl.startsWith('http://') && !fileUrl.startsWith('https://') && !fileUrl.startsWith('//');
  }

  static isVisioFileExtension(fileUrl: string) {
    return fileUrl.endsWith('.vsd') || fileUrl.endsWith('.vsdx') || fileUrl.endsWith('.vsdm');
  }

  public static parseLink(link: Visio.Hyperlink, baseUrl: string, defaultLabel: string): {
    external: boolean;
    url: string,
    label: string
  } {
    const { address, subAddress, description, extraInfo } = link;

    if (address) {
      if (Utils.isVisioFileExtension(address)) {
        if (Utils.isRelativeUrl(address)) {
          const absoluteUrl = baseUrl.substring(0, baseUrl.lastIndexOf('/') + 1) + address;
          return {
            external: false,
            url: Utils.joinPageUrl(absoluteUrl, subAddress),
            label: description || subAddress || address.replace(/\.vsdx$/, '').replace(/\.vsdm$/, '').replace(/\.vsd$/, '')
          }
        } else {
          return {
            external: false,
            url: Utils.joinPageUrl(address, subAddress),
            label: description || subAddress || defaultLabel
          }
        }
      } else {
        if (Utils.isRelativeUrl(address)) {
          const absoluteUrl = baseUrl.substring(0, baseUrl.lastIndexOf('/') + 1) + address;
          return {
            external: true,
            url: absoluteUrl,
            label: description || address || defaultLabel
          }
        } else {
          return {
            external: true,
            url: address + (extraInfo ? `?${extraInfo}` : ''),
            label: description || defaultLabel
          }
        }
      }
    } else {
      return {
        external: false,
        url: Utils.joinPageUrl(baseUrl, subAddress),
        label: description || subAddress || defaultLabel
      }
    }
  }
}
