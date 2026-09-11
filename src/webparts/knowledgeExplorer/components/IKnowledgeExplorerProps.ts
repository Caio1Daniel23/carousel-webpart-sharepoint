import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ICustomFolderImage } from '../IKnowledgeExplorerWebPartProps';

export interface IKnowledgeExplorerProps {
  context: WebPartContext;
  rootFolderPath: string;
  rootTitle: string;
  height: number;
  defaultCardImage: string;
  customFolderImages: ICustomFolderImage[];
}
