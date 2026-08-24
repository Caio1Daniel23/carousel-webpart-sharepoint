import { ISlide } from '../ICarouselWebPartProps';

export interface ICarouselProps {
  slides: ISlide[];
  autoplay: boolean;
  transitionTime: number;
  height: number;
  showArrows: boolean;
  showDots: boolean;
}
