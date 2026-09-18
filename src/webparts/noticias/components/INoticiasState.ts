import { INewsItem } from '../models/INewsItem';

export interface INoticiasState {
  loading: boolean;
  error: string;
  news: INewsItem[];
  currentPage: number;
  carouselIndex: number;
}
