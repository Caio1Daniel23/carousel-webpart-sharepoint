export interface IKnowledgeExplorerWebPartProps {
  rootFolderPath: string; // caminho relativo do servidor, ex: /sites/intranet/BaseConhecimento
  rootTitle: string;
  height: number;
  defaultCardImage: string;
  customFolderImages: ICustomFolderImage[];
  sortMode: 'name' | 'countDesc';
  hiddenFolders: string[];
  hiddenFolderExceptions: IHiddenFolderException[];
  folderDescriptions: IFolderDescription[];
}

export interface ICustomFolderImage {
  folderPath: string;
  imageUrl: string;
}

// Um usuário específico que pode ver uma pasta mesmo com ela oculta para todo mundo.
export interface IHiddenFolderAllowedUser {
  loginName: string; // identificador único usado pra comparar com o usuário logado (claims/UPN)
  email: string;
  displayName: string;
}

// A lista de exceções de uma pasta oculta específica.
export interface IHiddenFolderException {
  folderPath: string;
  allowedUsers: IHiddenFolderAllowedUser[];
}

// Um texto de apresentação de uma pasta específica, mostrado acima da lista de
// subpastas/arquivos ao entrar nela. Sem entrada aqui, nada é exibido.
export interface IFolderDescription {
  folderPath: string;
  description: string;
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
