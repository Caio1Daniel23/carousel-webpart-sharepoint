export interface IKnowledgeExplorerWebPartProps {
  rootFolderPath: string; // caminho relativo do servidor, ex: /sites/intranet/BaseConhecimento
  rootTitle: string;
  height: number;
  defaultCardImage: string;
  customFolderImages: ICustomFolderImage[];
}

export interface ICustomFolderImage {
  folderPath: string;
  imageUrl: string;
}

export interface IFolderItem {
  name: string;
  serverRelativeUrl: string;
  fileCount: number;
}

export interface IFileItem {
  name: string;
  serverRelativeUrl: string;
  modified: string;
  sizeBytes: number;
}

export interface IBreadcrumbItem {
  name: string;
  path: string;
}
