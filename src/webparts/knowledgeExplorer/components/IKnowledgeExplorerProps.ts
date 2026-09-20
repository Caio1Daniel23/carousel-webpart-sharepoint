import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ICustomFolderImage, IHiddenFolderException, IHiddenFolderAllowedUser } from '../IKnowledgeExplorerWebPartProps';
import { IFilePickerResult } from '@pnp/spfx-property-controls/lib/propertyFields/filePicker/filePickerControls/FilePicker.types';

export interface IKnowledgeExplorerProps {
  context: WebPartContext;
  rootFolderPath: string;
  rootTitle: string;
  height: number;
  defaultCardImage: string;
  customFolderImages: ICustomFolderImage[];
  sortMode: 'name' | 'countDesc';
  hiddenFolders: string[];
  hiddenFolderExceptions: IHiddenFolderException[];
  displayMode: DisplayMode;
  // Sobe o arquivo/imagem escolhido (upload ou link) e devolve a URL final utilizável —
  // reaproveita a mesma lógica de upload já usada para a imagem padrão.
  uploadPickedImage: (result: IFilePickerResult) => Promise<string>;
  // Grava (ou substitui) a imagem específica de uma pasta diretamente na propriedade da
  // web part, sem precisar abrir o painel de propriedades.
  onSetFolderImage: (folderPath: string, imageUrl: string) => void;
  // Remove a imagem específica de uma pasta, voltando a usar a imagem padrão.
  onRemoveFolderImage: (folderPath: string) => void;
  // Oculta ou reexibe uma pasta específica (alterna a presença dela em hiddenFolders).
  onToggleFolderHidden: (folderPath: string) => void;
  // Define a lista de usuários que podem ver uma pasta oculta mesmo assim (lista vazia remove a exceção).
  onSetFolderAllowedUsers: (folderPath: string, users: IHiddenFolderAllowedUser[]) => void;
}
