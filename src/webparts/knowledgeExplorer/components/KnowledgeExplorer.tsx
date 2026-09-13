import * as React from 'react';
import { useEffect, useState, useCallback, useRef } from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { DisplayMode } from '@microsoft/sp-core-library';
import { Icon } from '@fluentui/react/lib/Icon';
import { Callout, DirectionalHint } from '@fluentui/react/lib/Callout';
import { FilePicker } from '@pnp/spfx-property-controls/lib/propertyFields/filePicker/filePickerControls/FilePicker';
import { IFilePickerResult } from '@pnp/spfx-property-controls/lib/propertyFields/filePicker/filePickerControls/FilePicker.types';
import styles from './KnowledgeExplorer.module.scss';
import { IKnowledgeExplorerProps } from './IKnowledgeExplorerProps';
import { IFolderItem, IFileItem, IBreadcrumbItem, IDynamicColumn } from '../IKnowledgeExplorerWebPartProps';

type LoadState = 'loading' | 'loaded' | 'error';

// Pastas do sistema que o SharePoint cria sozinho e que não são categorias de verdade.
const IGNORED_FOLDER_NAMES = ['Forms'];

// Colunas "pseudo" que aparecem na exibição padrão mas que não são dados de verdade
// do arquivo (ícone, menu de contexto, nome — que já mostramos por conta própria).
const SKIP_VIEW_FIELDS = [
  'DocIcon',
  'LinkFilename',
  'LinkFilenameNoMenu',
  'LinkFilename2',
  'FileLeafRef',
  'Edit',
  'SelectTitle',
  'ItemChildCount',
  'FolderChildCount',
  'Attachments',
  '_CommentCount',
  '_LikeCount'
];

// Fallback usado somente se não for possível ler a exibição padrão da biblioteca
// (ex: permissão insuficiente) — assim o web part continua funcional.
const FALLBACK_COLUMNS: IDynamicColumn[] = [{ internalName: 'Modified', displayName: 'Data', typeAsString: 'DateTime' }];

const fileIconInfo = (name: string): { label: string; className: string } => {
  const ext = name.split('.').pop()?.toLowerCase() || '';
  if (ext === 'pdf') return { label: 'PDF', className: styles.iconPdf };
  if (['doc', 'docx'].indexOf(ext) >= 0) return { label: 'DOC', className: styles.iconDoc };
  if (['xls', 'xlsx'].indexOf(ext) >= 0) return { label: 'XLS', className: styles.iconXls };
  if (['ppt', 'pptx'].indexOf(ext) >= 0) return { label: 'PPT', className: styles.iconPpt };
  return { label: ext ? ext.toUpperCase().slice(0, 4) : 'ARQ', className: styles.iconGeneric };
};

const formatSize = (bytes: number): string => {
  if (!bytes) return '';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
};

const formatDate = (iso: string): string => {
  if (!iso) return '';
  const d = new Date(iso);
  return d.toLocaleDateString('pt-BR');
};

// Formata o valor de um campo de acordo com o tipo real da coluna no SharePoint,
// para que datas, pessoas, escolhas, links etc. apareçam de forma legível
// em vez do objeto/valor cru retornado pela API REST.
const formatFieldValue = (value: unknown, typeAsString: string): string => {
  if (value === null || value === undefined || value === '') return '—';

  switch (typeAsString) {
    case 'DateTime':
      return formatDate(value as string);

    case 'Boolean':
      return value ? 'Sim' : 'Não';

    case 'Number':
    case 'Currency':
      return Number(value).toLocaleString('pt-BR');

    case 'URL': {
      const urlVal = value as { Url?: string; Description?: string };
      return urlVal?.Description || urlVal?.Url || '—';
    }

    case 'User': {
      const userVal = value as { Title?: string };
      return userVal?.Title || '—';
    }

    case 'UserMulti': {
      const arr = (value as { results?: { Title: string }[] })?.results || [];
      return arr.length ? arr.map((u) => u.Title).join(', ') : '—';
    }

    case 'MultiChoice': {
      const arr = (value as { results?: string[] })?.results || [];
      return arr.length ? arr.join(', ') : '—';
    }

    default:
      return String(value);
  }
};

// Monta $select/$expand da chamada REST de arquivos com base nas colunas
// descobertas dinamicamente, incluindo o /Title extra necessário para
// campos de Pessoa conseguirem mostrar o nome de exibição.
const buildFilesQuery = (columns: IDynamicColumn[]): { select: string; expand: string } => {
  const selectSet = new Set<string>(['Name', 'ServerRelativeUrl', 'TimeLastModified', 'Length']);
  const expandSet = new Set<string>(['ListItemAllFields']);

  columns.forEach((col) => {
    if (col.typeAsString === 'User' || col.typeAsString === 'UserMulti') {
      selectSet.add(`ListItemAllFields/${col.internalName}/Title`);
      expandSet.add(`ListItemAllFields/${col.internalName}`);
    } else {
      selectSet.add(`ListItemAllFields/${col.internalName}`);
    }
  });

  const selectList: string[] = [];
  selectSet.forEach((value) => selectList.push(value));
  const expandList: string[] = [];
  expandSet.forEach((value) => expandList.push(value));

  return { select: selectList.join(','), expand: expandList.join(',') };
};

// --- Controle de concorrência e cache para a contagem recursiva de arquivos ---
//
// Sem isso, cada pasta (e cada subpasta dela, em cada nível) disparava chamadas REST
// em paralelo sem limite nenhum. Em bibliotecas com muitas pastas, isso faz o SharePoint
// responder com 429 (limitação de chamadas) — e como o código antigo tratava qualquer
// erro como "0 arquivos", pastas com conteúdo apareciam zeradas na tela.

// Nº máximo de chamadas de contagem em andamento ao mesmo tempo, não importa quantas
// pastas/subpastas existam — evita disparar centenas de chamadas simultâneas.
const MAX_CONCURRENT_COUNT_REQUESTS = 4;

// Se o SharePoint responder com limitação (429/503), espera um pouco e tenta de novo
// em vez de desistir e cravar 0. Os tempos crescem a cada nova tentativa.
const COUNT_RETRY_DELAYS_MS = [500, 1500, 3000];

// Guarda o resultado de cada pasta já contada com sucesso durante a sessão da página.
// Reseta sozinho ao recarregar a página — não guardamos isso em localStorage porque
// a contagem é sempre um retrato de "agora", e o conteúdo pode mudar entre visitas.
const folderCountCache = new Map<string, number>();

let activeCountRequests = 0;
const countRequestQueue: (() => void)[] = [];

const acquireCountSlot = (): Promise<void> =>
  new Promise((resolve) => {
    const tryAcquire = (): void => {
      if (activeCountRequests < MAX_CONCURRENT_COUNT_REQUESTS) {
        activeCountRequests++;
        resolve();
      } else {
        countRequestQueue.push(tryAcquire);
      }
    };
    tryAcquire();
  });

const releaseCountSlot = (): void => {
  activeCountRequests = Math.max(0, activeCountRequests - 1);
  const next = countRequestQueue.shift();
  if (next) next();
};

const delay = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

// Busca um endpoint REST com novas tentativas em caso de limitação (429/503).
// Só desiste de verdade depois de esgotar as tentativas — assim um "0" na tela
// significa que a pasta está mesmo vazia, não que a chamada falhou.
const fetchJsonWithRetry = async (
  context: IKnowledgeExplorerProps['context'],
  url: string
): Promise<{ ok: boolean; json: { value?: unknown[] } | undefined }> => {
  for (let attempt = 0; attempt <= COUNT_RETRY_DELAYS_MS.length; attempt++) {
    const response: SPHttpClientResponse = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (response.ok) {
      return { ok: true, json: await response.json() };
    }
    const isThrottled = response.status === 429 || response.status === 503;
    if (isThrottled && attempt < COUNT_RETRY_DELAYS_MS.length) {
      await delay(COUNT_RETRY_DELAYS_MS[attempt]);
      continue;
    }
    return { ok: false, json: undefined };
  }
  return { ok: false, json: undefined };
};

// Conta recursivamente (subpastas inclusas) quantos arquivos existem dentro de um caminho,
// percorrendo a árvore de pastas de verdade via REST. Mais lento que usar a Pesquisa, porém
// sempre correto na hora — não depende do índice de busca estar atualizado.
const countFilesRecursive = async (
  context: IKnowledgeExplorerProps['context'],
  serverRelativeUrl: string
): Promise<number> => {
  const cached = folderCountCache.get(serverRelativeUrl);
  if (cached !== undefined) return cached;

  // O "lugar na fila" (slot) só é ocupado durante a chamada REST em si — nunca
  // enquanto esperamos a recursão das subpastas terminar. Se segurássemos o slot
  // durante a espera, uma árvore de pastas mais funda que o limite de concorrência
  // travaria: pais esperando filhos, e filhos sem slot livre pra rodar.
  let filesResult: { ok: boolean; json: { value?: unknown[] } | undefined };
  let foldersResult: { ok: boolean; json: { value?: unknown[] } | undefined };

  await acquireCountSlot();
  try {
    const filesUrl =
      `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')` +
      `/Files?$select=Name`;
    const foldersUrl =
      `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')` +
      `/Folders?$select=Name,ServerRelativeUrl`;

    [filesResult, foldersResult] = await Promise.all([
      fetchJsonWithRetry(context, filesUrl),
      fetchJsonWithRetry(context, foldersUrl)
    ]);
  } catch (err) {
    return 0;
  } finally {
    releaseCountSlot();
  }

  // Não guarda no cache um resultado que pode estar incompleto por causa de uma
  // falha real da API — só cacheia contagens que sabemos que são de verdade.
  if (!filesResult.ok || !foldersResult.ok) return 0;

  const filesJson = filesResult.json as { value?: { Name: string }[] };
  const foldersJson = foldersResult.json as { value?: { Name: string; ServerRelativeUrl: string }[] };

  const directCount = (filesJson.value || []).length;
  const subfolders: { Name: string; ServerRelativeUrl: string }[] = (foldersJson.value || []).filter(
    (f: { Name: string }) => IGNORED_FOLDER_NAMES.indexOf(f.Name) === -1
  );

  let total = directCount;
  if (subfolders.length > 0) {
    const subCounts = await Promise.all(subfolders.map((f) => countFilesRecursive(context, f.ServerRelativeUrl)));
    total += subCounts.reduce((sum, c) => sum + c, 0);
  }

  folderCountCache.set(serverRelativeUrl, total);
  return total;
};

const KnowledgeExplorer: React.FC<IKnowledgeExplorerProps> = (props) => {
  const {
    context,
    rootFolderPath,
    rootTitle,
    height,
    defaultCardImage,
    customFolderImages,
    displayMode,
    uploadPickedImage,
    onSetFolderImage,
    onRemoveFolderImage
  } = props;

  const [breadcrumb, setBreadcrumb] = useState<IBreadcrumbItem[]>([]);
  const [folders, setFolders] = useState<IFolderItem[]>([]);
  const [files, setFiles] = useState<IFileItem[]>([]);
  const [columns, setColumns] = useState<IDynamicColumn[]>([]);
  const [state, setState] = useState<LoadState>('loading');
  const [errorDetail, setErrorDetail] = useState<string>('');

  // Controla qual card tem o seletor de imagem aberto no momento (no máximo um por vez).
  const [imagePickerOpenFor, setImagePickerOpenFor] = useState<string | null>(null);
  // Referências aos botões de "3 pontinhos" de cada card, usadas como âncora do Callout.
  const menuButtonRefs = useRef<{ [serverRelativeUrl: string]: HTMLButtonElement | null }>({});

  const getRecursiveFileCount = useCallback(
    (serverRelativeUrl: string): Promise<number> => countFilesRecursive(context, serverRelativeUrl),
    [context]
  );

  // Descobre, a partir de qualquer pasta da biblioteca, quais colunas estão configuradas
  // na exibição padrão ("DefaultView") dela — a mesma que aparece no SharePoint. Assim,
  // adicionar ou remover uma coluna lá se reflete aqui automaticamente, e funciona para
  // qualquer biblioteca que o web part aponte, não só esta.
  const resolveSchema = useCallback(
    async (anyFolderPath: string): Promise<IDynamicColumn[]> => {
      try {
        const webUrl = context.pageContext.web.absoluteUrl;

        const parentListUrl =
          `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(anyFolderPath)}')` +
          `/ListItemAllFields/ParentList?$select=Id`;
        const parentListResponse = await context.spHttpClient.get(parentListUrl, SPHttpClient.configurations.v1);
        if (!parentListResponse.ok) return FALLBACK_COLUMNS;
        const parentListJson = await parentListResponse.json();
        const listId = parentListJson.Id as string;
        if (!listId) return FALLBACK_COLUMNS;

        const viewFieldsUrl = `${webUrl}/_api/web/lists(guid'${listId}')/DefaultView/ViewFields/Items`;
        const fieldsUrl =
          `${webUrl}/_api/web/lists(guid'${listId}')/fields?$select=InternalName,Title,TypeAsString&$filter=Hidden eq false`;

        const [viewFieldsResponse, fieldsResponse] = await Promise.all([
          context.spHttpClient.get(viewFieldsUrl, SPHttpClient.configurations.v1),
          context.spHttpClient.get(fieldsUrl, SPHttpClient.configurations.v1)
        ]);
        if (!viewFieldsResponse.ok || !fieldsResponse.ok) return FALLBACK_COLUMNS;

        const viewFieldsJson = await viewFieldsResponse.json();
        const fieldsJson = await fieldsResponse.json();

        const orderedInternalNames: string[] = viewFieldsJson.value || [];
        const fieldsMeta: { InternalName: string; Title: string; TypeAsString: string }[] = fieldsJson.value || [];
        const fieldsByName: { [key: string]: { Title: string; TypeAsString: string } } = {};
        fieldsMeta.forEach((f) => {
          fieldsByName[f.InternalName] = { Title: f.Title, TypeAsString: f.TypeAsString };
        });

        const dynamicColumns: IDynamicColumn[] = orderedInternalNames
          .filter((name) => SKIP_VIEW_FIELDS.indexOf(name) === -1 && !!fieldsByName[name])
          .map((name) => ({
            internalName: name,
            displayName: fieldsByName[name].Title,
            typeAsString: fieldsByName[name].TypeAsString
          }));

        return dynamicColumns.length > 0 ? dynamicColumns : FALLBACK_COLUMNS;
      } catch (err) {
        return FALLBACK_COLUMNS;
      }
    },
    [context]
  );

  const loadFolder = useCallback(
    async (path: string, columnsOverride?: IDynamicColumn[]): Promise<void> => {
      setState('loading');
      setErrorDetail('');
      try {
        const activeColumns = columnsOverride || columns;
        const { select, expand } = buildFilesQuery(activeColumns);

        const foldersUrl =
          `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(path)}')` +
          `/Folders?$select=Name,ServerRelativeUrl`;
        const filesUrl =
          `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(path)}')` +
          `/Files?$select=${select}&$expand=${expand}`;

        const [foldersResponse, filesResponse] = await Promise.all([
          context.spHttpClient.get(foldersUrl, SPHttpClient.configurations.v1),
          context.spHttpClient.get(filesUrl, SPHttpClient.configurations.v1)
        ]);

        if (!foldersResponse.ok || !filesResponse.ok) {
          throw new Error(`HTTP ${foldersResponse.status} / ${filesResponse.status}`);
        }

        const foldersJson = await foldersResponse.json();
        const filesJson = await filesResponse.json();

        const rawFolders: { Name: string; ServerRelativeUrl: string }[] = (foldersJson.value || []).filter(
          (f: { Name: string }) => IGNORED_FOLDER_NAMES.indexOf(f.Name) === -1
        );

        // Mostra as pastas JÁ, com a contagem como "carregando" (null) — a tela não
        // fica esperando todas as contagens recursivas para aparecer.
        const folderItems: IFolderItem[] = rawFolders.map((f) => ({
          name: f.Name,
          serverRelativeUrl: f.ServerRelativeUrl,
          fileCount: null
        }));

        const fileItems: IFileItem[] = (filesJson.value || []).map(
          (f: {
            Name: string;
            ServerRelativeUrl: string;
            TimeLastModified: string;
            Length: string;
            ListItemAllFields?: { [key: string]: unknown };
          }) => {
            const fieldValues: { [internalName: string]: unknown } = {};
            activeColumns.forEach((col) => {
              fieldValues[col.internalName] = f.ListItemAllFields ? f.ListItemAllFields[col.internalName] : undefined;
            });
            return {
              name: f.Name,
              serverRelativeUrl: f.ServerRelativeUrl,
              modified: f.TimeLastModified,
              sizeBytes: Number(f.Length),
              fieldValues
            };
          }
        );

        setFolders(folderItems);
        setFiles(fileItems);
        setState('loaded');

        // Calcula a contagem de cada pasta em segundo plano, atualizando o card
        // assim que o resultado daquela pasta específica ficar pronto.
        rawFolders.forEach((f) => {
          getRecursiveFileCount(f.ServerRelativeUrl)
            .then((count) => {
              setFolders((prev) =>
                prev.map((item) => (item.serverRelativeUrl === f.ServerRelativeUrl ? { ...item, fileCount: count } : item))
              );
            })
            .catch(() => undefined);
        });
      } catch (err) {
        setState('error');
        setErrorDetail(path);
      }
    },
    [context, getRecursiveFileCount, columns]
  );

  useEffect(() => {
    if (!rootFolderPath) {
      setState('error');
      setErrorDetail('');
      return;
    }
    setBreadcrumb([{ name: rootTitle || 'Início', path: rootFolderPath }]);
    (async (): Promise<void> => {
      const cols = await resolveSchema(rootFolderPath);
      setColumns(cols);
      await loadFolder(rootFolderPath, cols);
    })().catch(() => undefined);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rootFolderPath]);

  const openFolder = (folder: IFolderItem): void => {
    setBreadcrumb((prev) => [...prev, { name: folder.name, path: folder.serverRelativeUrl }]);
    loadFolder(folder.serverRelativeUrl).catch(() => undefined);
  };

  const goToBreadcrumb = (index: number): void => {
    const target = breadcrumb[index];
    setBreadcrumb(breadcrumb.slice(0, index + 1));
    loadFolder(target.path).catch(() => undefined);
  };

  const goBack = (): void => {
    if (breadcrumb.length <= 1) return;
    goToBreadcrumb(breadcrumb.length - 2);
  };

  // Abre o arquivo numa janela pop-up. Se for um atalho (.url), primeiro lê o
  // conteúdo do arquivo pra descobrir o link de verdade pra onde ele aponta,
  // em vez de abrir/baixar o próprio arquivo de atalho.
  const openFile = async (file: IFileItem): Promise<void> => {
    let target = file.serverRelativeUrl;
    const ext = file.name.split('.').pop()?.toLowerCase();

    if (ext === 'url') {
      try {
        const contentUrl =
          `${context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(
            file.serverRelativeUrl
          )}')/$value`;
        const response = await context.spHttpClient.get(contentUrl, SPHttpClient.configurations.v1);
        const text = await response.text();
        const match = /URL=(.+)/i.exec(text);
        if (match && match[1]) {
          target = match[1].trim();
        }
      } catch (err) {
        // Se não conseguir ler o atalho, cai no comportamento padrão (abre o próprio arquivo).
      }
    } else if (['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx'].indexOf(ext || '') >= 0) {
      // Força abrir no visualizador do navegador em vez de baixar.
      target = `${target}?web=1`;
    }

    window.open(
      target,
      'baseConhecimentoDoc',
      'width=1100,height=800,resizable=yes,scrollbars=yes,toolbar=no,menubar=no,location=no,status=no'
    );
  };

  // Imagem customizada só se aplica a cards de 1º e 2º nível: quando estamos vendo a
  // raiz (breadcrumb.length === 1) os cards são de 1º nível; quando estamos dentro de
  // uma pasta de 1º nível (breadcrumb.length === 2) os cards são de 2º nível.
  const allowCustomImage = breadcrumb.length <= 2;
  const isEditMode = displayMode === DisplayMode.Edit;

  const getFolderImage = (folder: IFolderItem): string | undefined => {
    if (!allowCustomImage) return undefined;
    const override = (customFolderImages || []).filter((c) => c.folderPath === folder.serverRelativeUrl)[0];
    return override?.imageUrl || defaultCardImage || undefined;
  };

  const getCustomImageUrl = (folder: IFolderItem): string | undefined =>
    (customFolderImages || []).filter((c) => c.folderPath === folder.serverRelativeUrl)[0]?.imageUrl;

  // Chamado quando o usuário escolhe/envia uma imagem pelo seletor flutuante do card:
  // sobe o arquivo (reaproveitando a mesma lógica da imagem padrão) e grava direto na
  // propriedade da web part, sem passar pelo painel lateral.
  const handleImagePicked = async (folder: IFolderItem, result: IFilePickerResult): Promise<void> => {
    const url = await uploadPickedImage(result);
    if (url) {
      onSetFolderImage(folder.serverRelativeUrl, url);
    }
    setImagePickerOpenFor(null);
  };

  const handleRemoveFolderImage = (folder: IFolderItem): void => {
    onRemoveFolderImage(folder.serverRelativeUrl);
    setImagePickerOpenFor(null);
  };

  if (!rootFolderPath) {
    return (
      <div className={styles.explorer} style={{ minHeight: height }}>
        <div className={styles.emptyState}>
          Configure o caminho da pasta raiz no painel de propriedades da web part (ex:
          /sites/intranet/BaseConhecimento).
        </div>
      </div>
    );
  }

  return (
    <div className={styles.explorer} style={{ minHeight: height }}>
      <div className={styles.headerRow}>
        <div className={styles.breadcrumb}>
          {breadcrumb.map((item, idx) => (
            <span key={idx} className={styles.breadcrumbSegment}>
              {idx > 0 && <Icon iconName="ChevronRight" className={styles.breadcrumbSep} />}
              <button
                className={styles.breadcrumbButton}
                onClick={() => goToBreadcrumb(idx)}
                disabled={idx === breadcrumb.length - 1}
              >
                {item.name}
              </button>
            </span>
          ))}
        </div>
        {breadcrumb.length > 1 && (
          <button className={styles.backButton} onClick={goBack}>
            <Icon iconName="Back" /> Voltar
          </button>
        )}
      </div>

      {state === 'loading' && <div className={styles.loading}>Carregando...</div>}

      {state === 'error' && (
        <div className={styles.emptyState}>
          Não foi possível carregar a pasta{errorDetail ? ` "${errorDetail}"` : ''}. Confira se o caminho está
          correto e se você tem permissão de leitura nessa biblioteca.
        </div>
      )}

      {state === 'loaded' && (
        <>
          {folders.length > 0 && (
            <div className={styles.grid}>
              {folders.map((folder) => {
                const folderImage = getFolderImage(folder);
                const customImageUrl = getCustomImageUrl(folder);
                const showMenuButton = isEditMode && allowCustomImage;
                const isPickerOpen = imagePickerOpenFor === folder.serverRelativeUrl;

                return (
                  <div key={folder.serverRelativeUrl} className={styles.folderCard}>
                    <button className={styles.folderCardMain} onClick={() => openFolder(folder)}>
                      <span className={styles.folderMedia}>
                        <span className={styles.folderBadge}>{folder.fileCount === null ? '…' : folder.fileCount}</span>
                        {folderImage ? (
                          <img className={styles.folderImage} src={folderImage} alt="" />
                        ) : (
                          <Icon iconName="FabricFolder" className={styles.folderIcon} />
                        )}
                      </span>
                      <span className={styles.folderName}>{folder.name}</span>
                    </button>

                    {showMenuButton && (
                      <button
                        ref={(el): void => {
                          menuButtonRefs.current[folder.serverRelativeUrl] = el;
                        }}
                        className={styles.folderMenuButton}
                        onClick={(ev): void => {
                          ev.stopPropagation();
                          setImagePickerOpenFor((prev) => (prev === folder.serverRelativeUrl ? null : folder.serverRelativeUrl));
                        }}
                        aria-label={`Escolher imagem para ${folder.name}`}
                        title="Escolher imagem desta pasta"
                      >
                        <Icon iconName="MoreVertical" />
                      </button>
                    )}

                    {showMenuButton && isPickerOpen && menuButtonRefs.current[folder.serverRelativeUrl] && (
                      <Callout
                        target={menuButtonRefs.current[folder.serverRelativeUrl]}
                        onDismiss={(): void => setImagePickerOpenFor(null)}
                        directionalHint={DirectionalHint.bottomLeftEdge}
                        // O seletor de arquivo/imagem abre seu próprio painel modal por cima de tudo,
                        // fora da árvore do Callout — sem isso, o Callout se fecharia sozinho assim
                        // que o usuário clicasse dentro daquele painel para escolher a imagem.
                        preventDismissOnEvent={(ev: Event): boolean => {
                          const targetEl = ev.target as HTMLElement;
                          return !!(targetEl && targetEl.closest && targetEl.closest('.ms-Panel, .ms-Dialog, .ms-Modal'));
                        }}
                      >
                        <div className={styles.imagePickerCallout}>
                          <div className={styles.imagePickerTitle}>Imagem desta pasta</div>
                          {customImageUrl ? (
                            <img className={styles.imagePickerPreview} src={customImageUrl} alt="" />
                          ) : (
                            <div className={styles.imagePickerHint}>Usando a imagem padrão.</div>
                          )}
                          <FilePicker
                            {...({
                              context,
                              buttonIcon: 'Photo2',
                              buttonLabel: customImageUrl ? 'Alterar imagem' : 'Escolher imagem',
                              accepts: ['.gif', '.jpg', '.jpeg', '.png', '.webp', '.svg'],
                              onSave: (result: IFilePickerResult): void => {
                                handleImagePicked(folder, result).catch(() => undefined);
                              }
                            } as any)}
                          />
                          {customImageUrl && (
                            <button className={styles.imagePickerRemove} onClick={() => handleRemoveFolderImage(folder)}>
                              Usar imagem padrão
                            </button>
                          )}
                        </div>
                      </Callout>
                    )}
                  </div>
                );
              })}
            </div>
          )}

          {files.length > 0 && (
            <div
              className={styles.fileList}
              style={{ '--file-columns': columns.length } as React.CSSProperties}
            >
              <div className={styles.fileListHeader}>
                <span className={styles.fileHeaderName}>Nome</span>
                {columns.map((col) => (
                  <span key={col.internalName} className={styles.fileHeaderCell}>
                    {col.displayName}
                  </span>
                ))}
              </div>
              {files.map((file) => {
                const iconInfo = fileIconInfo(file.name);
                return (
                  <button
                    key={file.serverRelativeUrl}
                    className={styles.fileRow}
                    onClick={() => {
                      openFile(file).catch(() => undefined);
                    }}
                  >
                    <span className={styles.fileNameCell}>
                      <span className={`${styles.fileIcon} ${iconInfo.className}`}>{iconInfo.label}</span>
                      <span className={styles.fileInfo}>
                        <span className={styles.fileName}>{file.name}</span>
                        <span className={styles.fileMeta}>{file.sizeBytes ? formatSize(file.sizeBytes) : ''}</span>
                      </span>
                    </span>
                    {columns.map((col) => (
                      <span key={col.internalName} className={styles.fileCell}>
                        {formatFieldValue(file.fieldValues[col.internalName], col.typeAsString)}
                      </span>
                    ))}
                  </button>
                );
              })}
            </div>
          )}

          {folders.length === 0 && files.length === 0 && (
            <div className={styles.emptyState}>Esta pasta está vazia.</div>
          )}
        </>
      )}
    </div>
  );
};

export default KnowledgeExplorer;
