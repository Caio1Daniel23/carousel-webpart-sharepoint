import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './PeopleSearch.module.scss';
import { IPeopleSearchProps } from './IPeopleSearchProps';
import { IPersonResult } from '../IPeopleSearchWebPartProps';

// GUID fixo do SharePoint para a fonte de resultados "Local People Results"
// (a mesma usada internamente pelo web part nativo "Pessoas").
const PEOPLE_SOURCE_ID = 'b09a7990-05ea-4af9-81ef-edfab16c4e31';

const PeopleSearch: React.FC<IPeopleSearchProps> = (props) => {
  const { context, placeholderText, height } = props;
  const [query, setQuery] = useState<string>('');
  const [results, setResults] = useState<IPersonResult[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [selected, setSelected] = useState<IPersonResult | undefined>(undefined);
  const [showDropdown, setShowDropdown] = useState<boolean>(false);
  const debounceRef = useRef<number | undefined>(undefined);

  const photoUrl = (mail?: string): string => {
    if (!mail) return '';
    return `${context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=M&username=${encodeURIComponent(mail)}`;
  };

  // Transforma o formato de tabela do REST de pesquisa do SharePoint em objetos simples.
  const parseSearchResults = (json: any): IPersonResult[] => {
    const rows = json?.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];
    return rows.map((row: any) => {
      const cells: { Key: string; Value: string }[] = row.Cells || [];
      const getValue = (key: string): string | undefined => {
        const cell = cells.filter((c) => c.Key === key)[0];
        return cell ? cell.Value : undefined;
      };

      const workPhone = getValue('WorkPhone');

      return {
        id: getValue('AccountName') || getValue('PreferredName') || '',
        displayName: getValue('PreferredName') || getValue('AccountName') || '',
        mail: getValue('WorkEmail'),
        jobTitle: getValue('JobTitle'),
        businessPhones: workPhone ? [workPhone] : [],
        officeLocation: getValue('SPS-Location') || getValue('Office'),
        department: getValue('Department')
      } as IPersonResult;
    });
  };

  const search = async (text: string): Promise<void> => {
    if (!text || text.trim().length < 2) {
      setResults([]);
      setShowDropdown(false);
      return;
    }
    setLoading(true);
    try {
      const safeText = text.replace(/'/g, "''");
      const selectProps = 'PreferredName,WorkEmail,JobTitle,Department,WorkPhone,SPS-Location,Office,AccountName';
      const url =
        `${context.pageContext.web.absoluteUrl}/_api/search/query` +
        `?querytext='${encodeURIComponent(safeText)}*'` +
        `&sourceid='${PEOPLE_SOURCE_ID}'` +
        `&selectproperties='${encodeURIComponent(selectProps)}'` +
        `&rowlimit=6`;

      const response: SPHttpClientResponse = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }

      const json = await response.json();
      setResults(parseSearchResults(json));
      setShowDropdown(true);
    } catch (err) {
      setResults([]);
      setShowDropdown(false);
    } finally {
      setLoading(false);
    }
  };

  const onChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const value = e.target.value;
    setQuery(value);
    setSelected(undefined);

    if (debounceRef.current) {
      window.clearTimeout(debounceRef.current);
    }
    debounceRef.current = window.setTimeout(() => {
      search(value).catch(() => undefined);
    }, 350);
  };

  const onSelectPerson = (person: IPersonResult): void => {
    setSelected(person);
    setShowDropdown(false);
    setQuery(person.displayName);
  };

  const clearSearch = (): void => {
    setQuery('');
    setResults([]);
    setSelected(undefined);
    setShowDropdown(false);
  };

  useEffect(() => {
    return () => {
      if (debounceRef.current) window.clearTimeout(debounceRef.current);
    };
  }, []);

  return (
    <div className={styles.peopleSearch} style={{ minHeight: height }}>
      <div className={styles.header}>
        <Icon iconName="People" className={styles.headerIcon} />
        <span className={styles.headerTitle}>Pessoas</span>
      </div>

      <div className={styles.searchLabel}>Localizar pessoas</div>

      <div className={styles.searchBox}>
        <input
          type="text"
          className={styles.searchInput}
          placeholder={placeholderText || 'Insira um nome ou sobrenome'}
          value={query}
          onChange={onChange}
          onFocus={() => results.length > 0 && setShowDropdown(true)}
        />
        {query && (
          <button className={styles.clearButton} onClick={clearSearch} aria-label="Limpar">
            <Icon iconName="Cancel" />
          </button>
        )}
        <div className={styles.searchIconWrap}>
          <Icon iconName="Search" />
        </div>
      </div>

      {loading && <div className={styles.loading}>Buscando...</div>}

      {showDropdown && !selected && results.length > 0 && (
        <div className={styles.resultsList}>
          {results.map((person, idx) => (
            <button key={person.id || idx} className={styles.resultRow} onClick={() => onSelectPerson(person)}>
              <img className={styles.resultPhoto} src={photoUrl(person.mail)} alt={person.displayName} />
              <div>
                <div className={styles.resultName}>{person.displayName}</div>
                {person.jobTitle && <div className={styles.resultSubtitle}>{person.jobTitle}</div>}
              </div>
            </button>
          ))}
        </div>
      )}

      {showDropdown && !selected && !loading && query.trim().length >= 2 && results.length === 0 && (
        <div className={styles.noResults}>Nenhuma pessoa encontrada.</div>
      )}

      {selected && (
        <div className={styles.card}>
          <div className={styles.cardHeader}>
            <img className={styles.cardPhoto} src={photoUrl(selected.mail)} alt={selected.displayName} />
            <div className={styles.cardHeaderInfo}>
              <div className={styles.cardName}>{selected.displayName}</div>
              {selected.jobTitle && <div className={styles.cardJobTitle}>{selected.jobTitle}</div>}
            </div>
            <div className={styles.quickActions}>
              {selected.mail && (
                <a
                  className={styles.quickActionButton}
                  href={`https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(selected.mail)}`}
                  target="_blank"
                  rel="noopener noreferrer"
                  aria-label="Conversar no Teams"
                  title="Conversar no Teams"
                >
                  <Icon iconName="OfficeChat" />
                </a>
              )}
              {selected.mail && (
                <a
                  className={styles.quickActionButton}
                  href={`mailto:${selected.mail}`}
                  aria-label="Enviar e-mail"
                  title="Enviar e-mail"
                >
                  <Icon iconName="Mail" />
                </a>
              )}
              {selected.businessPhones && selected.businessPhones.length > 0 && (
                <a
                  className={styles.quickActionButton}
                  href={`tel:${selected.businessPhones[0]}`}
                  aria-label="Ligar"
                  title="Ligar"
                >
                  <Icon iconName="Phone" />
                </a>
              )}
            </div>
          </div>

          <div className={styles.cardSectionTitle}>Contato</div>
          <div className={styles.cardInfoList}>
            {selected.mail && (
              <div className={styles.cardInfoRow}>
                <Icon iconName="Mail" />
                <a href={`mailto:${selected.mail}`}>{selected.mail}</a>
              </div>
            )}
            {selected.businessPhones && selected.businessPhones.length > 0 && (
              <div className={styles.cardInfoRow}>
                <Icon iconName="Phone" />
                <span>{selected.businessPhones[0]}</span>
              </div>
            )}
            {selected.officeLocation && (
              <div className={styles.cardInfoRow}>
                <Icon iconName="MapPin" />
                <span>{selected.officeLocation}</span>
              </div>
            )}
            {selected.department && (
              <div className={styles.cardInfoRow}>
                <Icon iconName="Group" />
                <span>{selected.department}</span>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
};

export default PeopleSearch;
