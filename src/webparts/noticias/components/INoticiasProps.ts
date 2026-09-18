import { WebPartContext } from '@microsoft/sp-webpart-base';

export type LayoutMode = 'list' | 'carousel';

export interface INoticiasProps {
  context: WebPartContext;
  title: string;
  layoutMode: LayoutMode;
  pageSize: number;
  carouselCount: number;
  showAllNews: boolean;
  compactMode: boolean;
  showAuthor: boolean;
  showViews: boolean;
  showDate: boolean;
}
