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
  fileCount: number | null; // null = ainda calculando em segundo plano
}

export interface IFileItem {
  name: string;
  serverRelativeUrl: string;
  modified: string;
  sizeBytes: number;
  fieldValues: { [internalName: string]: unknown };
}

// Uma coluna descoberta dinamicamente a partir da exibição padrão da biblioteca.
// Assim, se o usuário adicionar/remover uma coluna dessa exibição no SharePoint,
// o web part reflete automaticamente, sem precisar tocar no código.
export interface IDynamicColumn {
  internalName: string;
  displayName: string;
  typeAsString: string; // Text, Note, DateTime, Number, Currency, Boolean, Choice, MultiChoice, URL, User, UserMulti...
}

export interface IBreadcrumbItem {
  name: string;
  path: string;
}
