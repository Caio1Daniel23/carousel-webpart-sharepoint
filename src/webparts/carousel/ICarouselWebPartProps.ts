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
  ctaOpenNewTab?: boolean;
}

export type DotsPosition = 'inside' | 'below';

export interface ICarouselWebPartProps {
  slides: ISlide[];
  autoplay: boolean;
  transitionTime: number; // segundos
  height: number; // px
  showArrows: boolean;
  showDots: boolean;
  dotsPosition: DotsPosition;
}
