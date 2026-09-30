import { INewsItem } from '../models/INewsItem';

export interface INoticiasState {
  loading: boolean;
  error: string;
  news: INewsItem[];
  currentPage: number;
  carouselIndex: number;
  // Largura real (em px) da própria web part, medida via ResizeObserver.
  // Não usamos @media (mede a janela do navegador) nem @container (o
  // compilador SASS deste projeto não processa essa sintaxe corretamente)
  // — medir em JavaScript funciona em qualquer navegador/toolchain.
  containerWidth: number;
}
