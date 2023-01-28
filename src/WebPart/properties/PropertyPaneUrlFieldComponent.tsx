import * as React from 'react';
import { FilePicker, IFilePickerResult } from '../../min-sp-controls-react/controls/filePicker';

import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/folders';
import '@pnp/sp/lists';
import '@pnp/sp/files';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { FolderExplorer, IFolder } from '../../min-sp-controls-react/controls/folderExplorer';
import { IDefaultFolder } from './IDefaultFolder';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { mergeStyles } from '@fluentui/react';
import * as strings from 'WebPartStrings';

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

  const siteUrl = new URL(props.context.pageContext.web.absoluteUrl);

  const isSameOrigin = (url: string) => {
    const fileUrl = new URL(url);
    return (siteUrl.origin === fileUrl.origin);
  };

  const ensureSiteAssetsFolder = async () => {
    const lib = await sp.web.lists.ensureSiteAssetsLibrary();
    return lib.rootFolder;
  }

  const onUploadFile = async (results: IFilePickerResult[]) => {
    const result = results[0];
    if (result.fileAbsoluteUrl && isSameOrigin(result.fileAbsoluteUrl)) {
      props.setUrl(result.fileAbsoluteUrl);
    } else {
      const fileConent = await result.downloadFileContent();
      const targetFolder = result.fileAbsoluteUrl
        ? await ensureSiteAssetsFolder()
        : await ensureUploadFolder(selectedFolder);
      const fileInfo = await targetFolder.files.add(result.fileName, fileConent, true);
      props.setUrl(`${siteUrl.origin}${fileInfo.data.ServerRelativeUrl}`);
    }
  };

  const rootFolder: IFolder = {
    Name: strings.UploadTo,
    ServerRelativeUrl: props.context.pageContext.web.serverRelativeUrl
  };

  const [documentsFolder, setDocumentsFolder] = React.useState<IFolder>();

  React.useEffect(() => {
    props.getDefaultFolder().then(f => {
      setDocumentsFolder({ Name: f.name, ServerRelativeUrl: f.relativeUrl });
      setSelectedFolder(f.relativeUrl);
    })
  }, []);

  const styles = mergeStyles({ marginTop: 40 });

  const renderCustomUploadTabContent = () => (
    <FolderExplorer
      className={styles}
      context={props.context}
      rootFolder={rootFolder}
      defaultFolder={documentsFolder}
      onSelect={folder => setSelectedFolder(folder.ServerRelativeUrl)}
      canCreateFolders={true}
    />
  );

  return (
    <Stack tokens={{ childrenGap: 's2' }}>
      <FilePicker
        label={strings.VisioDocument}
        accepts={['.vsd', '.vsdx', '.vsdm']}
        buttonLabel={strings.UrlPickerBrowse}
        onSave={(filePickerResult: IFilePickerResult[]) => onUploadFile(filePickerResult)}
        onChange={(filePickerResult: IFilePickerResult[]) => onChangeFile(filePickerResult)}
        defaultFolderAbsolutePath={`${siteUrl.origin}${documentsFolder?.ServerRelativeUrl}`}
        context={props.context}
        hideStockImages
        hideRecentTab
        hideLocalMultipleUploadTab
        renderCustomUploadTabContent={renderCustomUploadTabContent}
      />
      <Text variant='small'>{props.url || strings.UrlNotSelected}</Text>
    </Stack>
  );
}
