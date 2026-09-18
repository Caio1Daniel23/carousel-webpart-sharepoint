import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { INewsItem } from '../models/INewsItem';

/**
 * Reads every promoted (modern) news page in the site's "Site Pages" library,
 * newest first. Pulled once per load and paginated/sliced client-side by the
 * web part, so "load more" / page numbers never trigger a new network call.
 */
export class NewsService {
  public static async getAllNews(context: WebPartContext): Promise<INewsItem[]> {
    const siteUrl = context.pageContext.web.absoluteUrl;
    // O caminho da biblioteca de páginas modernas é sempre "/SitePages" no
    // servidor, mesmo em tenants com UI em português (onde ela aparece como
    // "Páginas do Site" e o Title da lista também vem traduzido). Referenciar
    // por getbytitle('Site Pages') falha (404) em tenants não-EN — usar
    // GetList pelo caminho do servidor funciona em qualquer idioma.
    const webServerRelativeUrl = context.pageContext.web.serverRelativeUrl.replace(/\/$/, '');
    const sitePagesPath = `${webServerRelativeUrl}/SitePages`;

    const baseFields = [
      'Id',
      'Title',
      'Description',
      'BannerImageUrl',
      'FirstPublishedDate',
      'Created',
      'Author/Title',
      'FileRef'
    ];

    const buildEndpoint = (fields: string[]): string =>
      `${siteUrl}/_api/web/GetList('${sitePagesPath}')/items` +
      `?$select=${fields.join(',')}` +
      `&$expand=Author` +
      `&$filter=PromotedState eq 2` +
      `&$orderby=FirstPublishedDate desc` +
      `&$top=2000`;

    // O campo de visualizações (analytics) nem sempre está disponível via
    // REST — depende da configuração de analytics do tenant/site. Tenta
    // primeiro com ele; se o servidor recusar o pedido (400), tenta de novo
    // sem esse campo, pra nunca deixar a lista de notícias inteira quebrar
    // por causa de uma feature opcional (número de visualizações fica 0).
    let response: SPHttpClientResponse = await context.spHttpClient.get(
      buildEndpoint([...baseFields, 'OData__ViewsLifeTime']),
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      response = await context.spHttpClient.get(
        buildEndpoint(baseFields),
        SPHttpClient.configurations.v1
      );
    }

    if (!response.ok) {
      throw new Error(`Falha ao carregar notícias (${response.status})`);
    }

    const json = await response.json();
    const items: any[] = json.value || [];

    return items.map((item) => NewsService.mapItem(item, siteUrl));
  }

  private static mapItem(item: any, siteUrl: string): INewsItem {
    let imageUrl = '';
    if (item.BannerImageUrl) {
      // BannerImageUrl comes back either as an object ({Url, ...}) or as a
      // JSON string, depending on tenant/CDN config — handle both.
      if (typeof item.BannerImageUrl === 'string') {
        try {
          const parsed = JSON.parse(item.BannerImageUrl);
          imageUrl = parsed.Url || parsed.serverRelativeUrl || '';
        } catch {
          imageUrl = item.BannerImageUrl;
        }
      } else {
        imageUrl = item.BannerImageUrl.Url || item.BannerImageUrl.serverRelativeUrl || '';
      }
    }

    const publishedRaw = item.FirstPublishedDate || item.Created;

    return {
      id: item.Id,
      title: item.Title || '',
      description: item.Description || '',
      imageUrl,
      url: item.FileRef ? `${new URL(siteUrl).origin}${item.FileRef}` : '#',
      author: item.Author && item.Author.Title ? item.Author.Title : '',
      publishedDate: publishedRaw ? new Date(publishedRaw) : new Date(),
      views: item.OData__ViewsLifeTime ? Number(item.OData__ViewsLifeTime) : 0
    };
  }
}
