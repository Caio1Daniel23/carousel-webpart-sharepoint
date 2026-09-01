import * as React from 'react';
import { useEffect, useRef, useState, useCallback } from 'react';
import styles from './Carousel.module.scss';
import { ICarouselProps } from './ICarouselProps';
import { ISlide } from '../ICarouselWebPartProps';
import { Icon } from '@fluentui/react/lib/Icon';

const Carousel: React.FC<ICarouselProps> = (props) => {
  const { slides, autoplay, transitionTime, height, showArrows, showDots } = props;
  const [current, setCurrent] = useState<number>(0);
  const timerRef = useRef<number | undefined>(undefined);

  const goTo = useCallback(
    (index: number) => {
      if (slides.length === 0) return;
      const next = (index + slides.length) % slides.length;
      setCurrent(next);
    },
    [slides.length]
  );

  const next = useCallback(() => goTo(current + 1), [current, goTo]);
  const prev = useCallback(() => goTo(current - 1), [current, goTo]);

  useEffect(() => {
    if (!autoplay || slides.length <= 1) {
      return undefined;
    }
    timerRef.current = window.setInterval(() => {
      setCurrent((c) => (c + 1) % slides.length);
    }, Math.max(1, transitionTime) * 1000);

    return () => {
      if (timerRef.current) {
        window.clearInterval(timerRef.current);
      }
    };
  }, [autoplay, transitionTime, slides.length]);

  if (!slides || slides.length === 0) {
    return (
      <div className={styles.carousel} style={{ height }}>
        <div className={styles.empty}>
          Nenhum slide configurado. Edite o web part e adicione imagens no painel de propriedades.
        </div>
      </div>
    );
  }

  const linkTargetProps = (slide: ISlide): { target?: string; rel?: string } => {
    return slide.ctaOpenNewTab
      ? { target: '_blank', rel: 'noopener noreferrer' }
      : {};
  };

  const renderCta = (slide: ISlide): JSX.Element | null => {
    if (!slide.ctaType || slide.ctaType === 'none' || !slide.ctaLink) {
      return null;
    }

    switch (slide.ctaType) {
      case 'button':
        return (
          <a className={styles.ctaButton} href={slide.ctaLink} {...linkTargetProps(slide)}>
            {slide.ctaText || 'Saiba mais'}
          </a>
        );
      case 'icon':
        return (
          <a
            className={styles.ctaIcon}
            href={slide.ctaLink}
            {...linkTargetProps(slide)}
            aria-label={slide.ctaText || 'Abrir link'}
            title={slide.ctaText || 'Abrir link'}
          >
            <Icon iconName={slide.ctaIcon || 'ChevronRight'} />
          </a>
        );
      case 'text':
        return (
          <a className={styles.ctaText} href={slide.ctaLink} {...linkTargetProps(slide)}>
            {slide.ctaText || 'Saiba mais'}
            <Icon iconName="ChevronRight" />
          </a>
        );
      case 'card':
        // Renderizado separadamente cobrindo o slide inteiro (ver abaixo).
        return null;
      default:
        return null;
    }
  };

  return (
    <div className={styles.carousel} style={{ height }}>
      <div
        className={styles.track}
        style={{ transform: `translateX(-${current * (100 / slides.length)}%)`, width: `${slides.length * 100}%` }}
      >
        {slides.map((slide, idx) => (
          <div className={styles.slide} style={{ width: `${100 / slides.length}%` }} key={idx}>
            {slide.ctaType === 'card' && slide.ctaLink && (
              <a
                className={styles.cardLink}
                href={slide.ctaLink}
                {...linkTargetProps(slide)}
                aria-label={slide.title || 'Abrir link'}
              />
            )}
            {slide.imageUrl && (
              <img className={styles.slideImage} src={slide.imageUrl} alt={slide.imageAlt || slide.title || ''} />
            )}
            {(slide.preHeader || slide.title || slide.description || slide.ctaType !== 'none') && (
              <div className={styles.overlay}>
                {slide.preHeader && <div className={styles.preHeader}>{slide.preHeader}</div>}
                {slide.title && <div className={styles.title}>{slide.title}</div>}
                {slide.description && <div className={styles.description}>{slide.description}</div>}
                {renderCta(slide)}
              </div>
            )}
          </div>
        ))}
      </div>

      {showArrows && slides.length > 1 && (
        <>
          <button className={`${styles.arrow} ${styles.arrowLeft}`} onClick={prev} aria-label="Slide anterior">
            <Icon iconName="ChevronLeft" />
          </button>
          <button className={`${styles.arrow} ${styles.arrowRight}`} onClick={next} aria-label="Próximo slide">
            <Icon iconName="ChevronRight" />
          </button>
        </>
      )}

      {showDots && slides.length > 1 && (
        <div className={styles.dots}>
          {slides.map((_, idx) => (
            <button
              key={idx}
              className={`${styles.dot} ${idx === current ? styles.active : ''}`}
              onClick={() => goTo(idx)}
              aria-label={`Ir para o slide ${idx + 1}`}
            />
          ))}
        </div>
      )}
    </div>
  );
};

export default Carousel;
