export type CtaType = 'none' | 'button' | 'icon' | 'text' | 'card';

export interface ISlide {
  imageUrl: string;
  imageAlt?: string;
  preHeader?: string;
  title?: string;
  description?: string;
  ctaType: CtaType;
  ctaText?: string;
  ctaIcon?: string;
  ctaLink?: string;
}

export interface ICarouselWebPartProps {
  slides: ISlide[];
  autoplay: boolean;
  transitionTime: number; // segundos
  height: number; // px
  showArrows: boolean;
  showDots: boolean;
}
