import * as React from 'react';
import styles from './Noticias.module.scss';
import { INoticiasProps } from './INoticiasProps';
import { INoticiasState } from './INoticiasState';
import { INewsItem } from '../models/INewsItem';
import { NewsService } from '../services/NewsService';

const MONTHS_PT = [
  'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
  'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'
];

function formatDatePt(date: Date): string {
  return `${date.getDate()} de ${MONTHS_PT[date.getMonth()]}`;
}

function initials(title: string): string {
  return (title || '')
    .split(' ')
    .filter(Boolean)
    .slice(0, 2)
    .map((w) => w.charAt(0).toUpperCase())
    .join('');
}

// O texto de "Descrição" vem do que o autor da notícia escreveu no
// SharePoint, e às vezes inclui emoji (✈️, 🤝, etc.) digitados por ele.
// Remove esses emoji do resumo exibido na lista, mantendo o texto normal.
const EMOJI_REGEX = /[\u{1F1E6}-\u{1FAFF}\u{2600}-\u{27BF}\u{2B00}-\u{2BFF}\u{2190}-\u{21FF}️‍]/gu;

function stripEmoji(text: string): string {
  return (text || '').replace(EMOJI_REGEX, '').replace(/\s{2,}/g, ' ').trim();
}

export default class Noticias extends React.Component<INoticiasProps, INoticiasState> {
  constructor(props: INoticiasProps) {
    super(props);
    this.state = {
      loading: true,
      error: '',
      news: [],
      currentPage: 1,
      carouselIndex: 0
    };
  }

  public componentDidMount(): void {
    this.loadNews();
  }

  public componentDidUpdate(prevProps: INoticiasProps): void {
    if (
      prevProps.layoutMode !== this.props.layoutMode ||
      prevProps.pageSize !== this.props.pageSize ||
      prevProps.carouselCount !== this.props.carouselCount ||
      prevProps.showAllNews !== this.props.showAllNews
    ) {
      // eslint-disable-next-line react/no-did-update-set-state
      this.setState({ currentPage: 1, carouselIndex: 0 });
    }
  }

  private loadNews = async (): Promise<void> => {
    try {
      const news = await NewsService.getAllNews(this.props.context);
      this.setState({ news, loading: false });
    } catch (err) {
      this.setState({ loading: false, error: (err as Error).message || 'Erro ao carregar notícias.' });
    }
  };

  private goToPage = (page: number): void => {
    this.setState({ currentPage: page });
  };

  private shiftCarousel = (dir: number): void => {
    const { carouselCount } = this.props;
    this.setState((prev) => ({
      carouselIndex: Math.max(0, Math.min(carouselCount - 1, prev.carouselIndex + dir))
    }));
  };

  private renderThumb(item: INewsItem): JSX.Element {
    return item.imageUrl ? (
      <img className={styles.thumbImg} src={item.imageUrl} alt="" />
    ) : (
      <div className={styles.thumbFallback}>{initials(item.title)}</div>
    );
  }

  private renderMeta(item: INewsItem): JSX.Element | null {
    const { showDate, showAuthor, showViews, showIcons } = this.props;
    if (!showDate && !showAuthor && !showViews) return null;
    return (
      <div className={styles.newsMeta}>
        {showDate && (
          <span>
            {showIcons && <span aria-hidden="true">🗓 </span>}
            {formatDatePt(item.publishedDate)}
          </span>
        )}
        {showAuthor && item.author && (
          <span>
            {showIcons ? <span aria-hidden="true">✎ </span> : 'Por '}
            {item.author}
          </span>
        )}
        {showViews && (
          <span>
            {showIcons && <span aria-hidden="true">👁 </span>}
            {item.views} visualizações
          </span>
        )}
      </div>
    );
  }

  private renderRow(item: INewsItem): JSX.Element {
    return (
      <a key={item.id} className={styles.newsRow} href={item.url}>
        <div className={styles.thumb}>{this.renderThumb(item)}</div>
        <div>
          <h3>{item.title}</h3>
          {item.description && <div className={styles.desc}>{stripEmoji(item.description)}</div>}
          {this.renderMeta(item)}
        </div>
      </a>
    );
  }

  private renderPagination(totalItems: number, onGo: (p: number) => void): JSX.Element | null {
    const { showAllNews, pageSize } = this.props;
    if (showAllNews) return null;
    const totalPages = Math.max(1, Math.ceil(totalItems / pageSize));
    if (totalPages <= 1) return null;

    const currentPage = Math.min(this.state.currentPage, totalPages);
    const nums: number[] = [1];
    if (currentPage > 3) nums.push(-1); // ellipsis marker
    for (let p = Math.max(2, currentPage - 1); p <= Math.min(totalPages - 1, currentPage + 1); p++) {
      nums.push(p);
    }
    if (currentPage < totalPages - 2) nums.push(-1);
    if (totalPages > 1) nums.push(totalPages);

    return (
      <>
        <div className={styles.pagination}>
          <button
            className={styles.pageArrow}
            disabled={currentPage <= 1}
            onClick={() => onGo(currentPage - 1)}
            aria-label="Página anterior"
          >
            ‹
          </button>
          {nums.map((p, i) =>
            p === -1 ? (
              <span key={`e${i}`} className={styles.pageEllipsis}>…</span>
            ) : (
              <button
                key={p}
                className={`${styles.pageNum} ${p === currentPage ? styles.active : ''}`}
                aria-current={p === currentPage ? 'page' : undefined}
                onClick={() => onGo(p)}
              >
                {p}
              </button>
            )
          )}
          <button
            className={styles.pageArrow}
            disabled={currentPage >= totalPages}
            onClick={() => onGo(currentPage + 1)}
            aria-label="Próxima página"
          >
            ›
          </button>
        </div>
        <div className={styles.progressNote}>
          Página {currentPage} de {totalPages} · {totalItems} notícias no total
        </div>
      </>
    );
  }

  private renderList(): JSX.Element {
    const { showAllNews, pageSize } = this.props;
    const { news, currentPage } = this.state;
    const offset = showAllNews ? 0 : (currentPage - 1) * pageSize;
    const visible = showAllNews ? news : news.slice(offset, offset + pageSize);

    return (
      <div>
        {visible.map((item) => this.renderRow(item))}
        {this.renderPagination(news.length, this.goToPage)}
      </div>
    );
  }

  private renderCarousel(): JSX.Element {
    const { carouselCount, showAllNews, pageSize } = this.props;
    const { news, currentPage, carouselIndex } = this.state;

    const featured = news.slice(0, carouselCount);
    const rest = news.slice(carouselCount);
    const restOffset = showAllNews ? 0 : (currentPage - 1) * pageSize;
    const visibleRest = showAllNews ? rest : rest.slice(restOffset, restOffset + pageSize);

    return (
      <div>
        <div className={styles.sectionLabel}>Em destaque · {carouselCount} mais recentes</div>
        <div className={styles.carouselWrap}>
          <button
            className={`${styles.carNav} ${styles.prev}`}
            onClick={() => this.shiftCarousel(-1)}
            aria-label="Anterior"
          >
            ‹
          </button>
          <div className={styles.carouselTrack}>
            {featured.map((item) => (
              <a key={item.id} className={styles.carCard} href={item.url}>
                <div className={styles.carThumb}>{this.renderThumb(item)}</div>
                <div className={styles.carBody}>
                  <h3>{item.title}</h3>
                  {this.props.showDate && (
                    <div className={styles.carDate}>{formatDatePt(item.publishedDate)}</div>
                  )}
                </div>
              </a>
            ))}
          </div>
          <button
            className={`${styles.carNav} ${styles.next}`}
            onClick={() => this.shiftCarousel(1)}
            aria-label="Próxima"
          >
            ›
          </button>
          <div className={styles.carDots}>
            {featured.map((_, i) => (
              <span key={i} className={i === carouselIndex ? styles.active : ''} />
            ))}
          </div>
        </div>

        <div className={styles.sectionLabel}>Demais notícias</div>
        {visibleRest.map((item) => this.renderRow(item))}
        {this.renderPagination(rest.length, this.goToPage)}
      </div>
    );
  }

  public render(): React.ReactElement<INoticiasProps> {
    const { title, layoutMode, compactMode } = this.props;
    const { loading, error, news } = this.state;

    return (
      <div className={`${styles.noticias} ${compactMode ? styles.compact : ''}`}>
        <div className={styles.header}>
          <h2>{title}</h2>
          {!loading && !error && <span className={styles.count}>{news.length} publicadas</span>}
        </div>

        <div className={styles.body}>
          {loading && <div className={styles.stateMsg}>Carregando notícias…</div>}
          {!loading && error && <div className={styles.stateMsg}>{error}</div>}
          {!loading && !error && news.length === 0 && (
            <div className={styles.stateMsg}>Nenhuma notícia publicada ainda.</div>
          )}
          {!loading && !error && news.length > 0 && (
            layoutMode === 'carousel' ? this.renderCarousel() : this.renderList()
          )}
        </div>
      </div>
    );
  }
}
