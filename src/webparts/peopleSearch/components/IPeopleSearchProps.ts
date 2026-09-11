import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface IPeopleSearchProps {
  context: WebPartContext;
  placeholderText: string;
  height: number;
  showHeader: boolean;
  headerText: string;
  searchLabelText: string;
}
