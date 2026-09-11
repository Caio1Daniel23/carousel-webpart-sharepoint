import * as React from 'react';
import { useEffect, useState, useCallback } from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './KnowledgeExplorer.module.scss';
import { IKnowledgeExplorerProps } from './IKnowledgeExplorerProps';
import { IFolderItem, IFileItem, IBreadcrumbItem } from '../IKnowledgeExplorerWebPartProps';

type LoadState = 'loading' | 'loaded' | 'error';

// Pastas do sistema que o SharePoint cria sozinho e que não são categorias de verdade.
const IGNORED_FOLDER_NAMES = ['Forms'];

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

// Conta recursivamente (subpastas inclusas) quantos arquivos existem dentro de um caminho,
// percorrendo a árvore de pastas de verdade via REST. Mais lento que usar a Pesquisa, porém
// sempre correto na hora — não depende do índice de busca estar atualizado.
const countFilesRecursive = async (
  context: IKnowledgeExplorerProps['context'],
  serverRelativeUrl: string
): Promise<number> => {
  try {
    const filesUrl =
      `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')` +
      `/Files?$select=Name`;
    const foldersUrl =
      `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')` +
      `/Folders?$select=Name,ServerRelativeUrl`;

    const [filesResponse, foldersResponse] = await Promise.all([
      context.spHttpClient.get(filesUrl, SPHttpClient.configurations.v1),
      context.spHttpClient.get(foldersUrl, SPHttpClient.configurations.v1)
    ]);

    if (!filesResponse.ok || !foldersResponse.ok) return 0;

    const filesJson = await filesResponse.json();
    const foldersJson = await foldersResponse.json();

    const directCount = (filesJson.value || []).length;
    const subfolders: { Name: string; ServerRelativeUrl: string }[] = (foldersJson.value || []).filter(
      (f: { Name: string }) => IGNORED_FOLDER_NAMES.indexOf(f.Name) === -1
    );

    if (subfolders.length === 0) {
      return directCount;
    }

    const subCounts = await Promise.all(
      subfolders.map((f) => countFilesRecursive(context, f.ServerRelativeUrl))
    );

    return directCount + subCounts.reduce((sum, c) => sum + c, 0);
  } catch (err) {
    return 0;
  }
};

const KnowledgeExplorer: React.FC<IKnowledgeExplorerProps> = (props) => {
  const { context, rootFolderPath, rootTitle, height, defaultCardImage, customFolderImages } = props;

  const [breadcrumb, setBreadcrumb] = useState<IBreadcrumbItem[]>([]);
  const [folders, setFolders] = useState<IFolderItem[]>([]);
  const [files, setFiles] = useState<IFileItem[]>([]);
  const [state, setState] = useState<LoadState>('loading');
  const [errorDetail, setErrorDetail] = useState<string>('');

  const getRecursiveFileCount = useCallback(
    (serverRelativeUrl: string): Promise<number> => countFilesRecursive(context, serverRelativeUrl),
    [context]
  );

  const loadFolder = useCallback(
    async (path: string): Promise<void> => {
      setState('loading');
      setErrorDetail('');
      try {
        const foldersUrl =
          `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(path)}')` +
          `/Folders?$select=Name,ServerRelativeUrl`;
        const filesUrl =
          `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(path)}')` +
          `/Files?$select=Name,ServerRelativeUrl,TimeLastModified,Length`;

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

        // Busca a contagem recursiva de cada subpasta em paralelo.
        const counts = await Promise.all(rawFolders.map((f) => getRecursiveFileCount(f.ServerRelativeUrl)));

        const folderItems: IFolderItem[] = rawFolders.map((f, idx) => ({
          name: f.Name,
          serverRelativeUrl: f.ServerRelativeUrl,
          fileCount: counts[idx]
        }));

        const fileItems: IFileItem[] = (filesJson.value || []).map(
          (f: { Name: string; ServerRelativeUrl: string; TimeLastModified: string; Length: string }) => ({
            name: f.Name,
            serverRelativeUrl: f.ServerRelativeUrl,
            modified: f.TimeLastModified,
            sizeBytes: Number(f.Length)
          })
        );

        setFolders(folderItems);
        setFiles(fileItems);
        setState('loaded');
      } catch (err) {
        setState('error');
        setErrorDetail(path);
      }
    },
    [context, getRecursiveFileCount]
  );

  useEffect(() => {
    if (!rootFolderPath) {
      setState('error');
      setErrorDetail('');
      return;
    }
    setBreadcrumb([{ name: rootTitle || 'Início', path: rootFolderPath }]);
    loadFolder(rootFolderPath).catch(() => undefined);
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

  // Imagem customizada só se aplica a cards de 1º e 2º nível: quando estamos vendo a
  // raiz (breadcrumb.length === 1) os cards são de 1º nível; quando estamos dentro de
  // uma pasta de 1º nível (breadcrumb.length === 2) os cards são de 2º nível.
  const allowCustomImage = breadcrumb.length <= 2;

  const getFolderImage = (folder: IFolderItem): string | undefined => {
    if (!allowCustomImage) return undefined;
    const override = (customFolderImages || []).filter((c) => c.folderPath === folder.serverRelativeUrl)[0];
    return override?.imageUrl || defaultCardImage || undefined;
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
                return (
                  <button key={folder.serverRelativeUrl} className={styles.folderCard} onClick={() => openFolder(folder)}>
                    <span className={styles.folderMedia}>
                      <span className={styles.folderBadge}>{folder.fileCount}</span>
                      {folderImage ? (
                        <img className={styles.folderImage} src={folderImage} alt="" />
                      ) : (
                        <Icon iconName="FabricFolder" className={styles.folderIcon} />
                      )}
                    </span>
                    <span className={styles.folderName}>{folder.name}</span>
                  </button>
                );
              })}
            </div>
          )}

          {files.length > 0 && (
            <div className={styles.fileList}>
              {files.map((file) => {
                const iconInfo = fileIconInfo(file.name);
                return (
                  <a
                    key={file.serverRelativeUrl}
                    className={styles.fileRow}
                    href={file.serverRelativeUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                  >
                    <span className={`${styles.fileIcon} ${iconInfo.className}`}>{iconInfo.label}</span>
                    <span className={styles.fileInfo}>
                      <span className={styles.fileName}>{file.name}</span>
                      <span className={styles.fileMeta}>
                        {formatDate(file.modified)} {file.sizeBytes ? `· ${formatSize(file.sizeBytes)}` : ''}
                      </span>
                    </span>
                  </a>
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
