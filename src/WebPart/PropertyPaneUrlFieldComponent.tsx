import * as React from 'react';
import { FilePicker, IFilePickerResult } from './../min-sp-controls-react/controls/filePicker';

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/files";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { FolderExplorer, IFolder } from '../min-sp-controls-react/controls/folderExplorer';
import { IDefaultFolder } from './IDefaultFolder';

export function PropertyPaneUrlFieldComponent(props: {
  url: string;
  setUrl: (url: string) => void;
  context: WebPartContext;
  getDefaultFolder: () => Promise<IDefaultFolder>;
}) {

  const onChangeFile = (results: IFilePickerResult[]) => {
    const result = results[0];
    props.setUrl(result.fileAbsoluteUrl);
  };

  const ensureUploadFolder = async (uploadPath: string) => {
    try {
      const existingFolder = sp.web.getFolderByServerRelativePath(uploadPath);
      const folderInfo = await existingFolder.select('Exists')();
      if (folderInfo.Exists) {
        return existingFolder;
      } else {
        const { folder: newFolder } = await sp.web.folders.addUsingPath(uploadPath);
        return newFolder;
      }
    } catch (error) {
      const siteAssetsLib = await sp.web.lists.ensureSiteAssetsLibrary();
      return siteAssetsLib.rootFolder;
    }
  }

  const [selectedFolder, setSelectedFolder] = React.useState<string>();

  const onUploadFile = async (results: IFilePickerResult[]) => {
    const result = results[0];
    if (result.fileAbsoluteUrl) {
      props.setUrl(result.fileAbsoluteUrl);
    } else {
      const fileConent = await result.downloadFileContent();
      const targetFolder = await ensureUploadFolder(selectedFolder);
      const fileInfo = await targetFolder.files.add(result.fileName, fileConent, true);
      props.setUrl(fileInfo.data.ServerRelativeUrl);
    }
  };

  const rootFolder: IFolder = {
    Name: "Site",
    ServerRelativeUrl: props.context.pageContext.web.serverRelativeUrl
  };

  const [documentsFolder, setDocumentsFolder] = React.useState<IFolder>();

  React.useEffect(() => {
    props.getDefaultFolder().then(f => {
      setDocumentsFolder({ Name: f.name, ServerRelativeUrl: f.relativeUrl });
      setSelectedFolder(f.relativeUrl);
    })
  }, []);

  const renderCustomUploadTabContent = () => (
    <FolderExplorer
      context={props.context}
      rootFolder={rootFolder}
      defaultFolder={documentsFolder}
      onSelect={folder => setSelectedFolder(folder.ServerRelativeUrl)}
      canCreateFolders={true}
    />
  );

  const siteUrl = new URL(props.context.pageContext.web.absoluteUrl);
  const fileName = props.url?.split('/').pop().split('?')[0].split('#')[0];

  return (
    <FilePicker
      label={fileName ?? 'Visio Document'}
      accepts={[".vsd", ".vsdx", ".vsdm"]}
      buttonLabel="Browse..."
      onSave={(filePickerResult: IFilePickerResult[]) => onUploadFile(filePickerResult)}
      onChange={(filePickerResult: IFilePickerResult[]) => onChangeFile(filePickerResult)}
      defaultFolderAbsolutePath={`${siteUrl.origin}${documentsFolder?.ServerRelativeUrl}`}
      context={props.context}
      hideStockImages
      hideRecentTab
      hideLocalMultipleUploadTab
      renderCustomUploadTabContent={renderCustomUploadTabContent}
    />
  );
}
