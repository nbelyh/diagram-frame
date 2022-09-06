import * as React from 'react';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react';

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/files";
import { WebPartContext } from '@microsoft/sp-webpart-base';

export function PropertyPaneUrlFieldComponent(props: {
  url: string;
  setUrl: (url: string) => void;
  context: WebPartContext
}) {

  const onChangeImage = (results: IFilePickerResult[]) => {
    const result = results[0];
    props.setUrl(result.fileAbsoluteUrl);
  };

  const onUploadImage = async (results: IFilePickerResult[]) => {
    const result = results[0];
    const fileConent = await result.downloadFileContent();
    const siteAssetsList = await sp.web.lists.ensureSiteAssetsLibrary();
    const fileInfo = await siteAssetsList.rootFolder.files.add(result.fileName, fileConent, true);
    props.setUrl(fileInfo.data.ServerRelativeUrl);
  };

  const fileName = props.url.split('/').pop().split('?')[0].split('#')[0];

  return (
    <FilePicker
      label={fileName ?? 'Visio Document'}
      accepts={[".vsd", ".vsdx", ".vsdm"]}
      buttonLabel="Browse..."
      onSave={(filePickerResult: IFilePickerResult[]) => onUploadImage(filePickerResult)}
      onChange={(filePickerResult: IFilePickerResult[]) => onChangeImage(filePickerResult)}
      context={props.context}
      hideStockImages
      hideLocalMultipleUploadTab
    />
  );
}
