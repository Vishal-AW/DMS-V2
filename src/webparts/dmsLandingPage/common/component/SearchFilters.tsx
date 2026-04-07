/* eslint-disable */
import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  SearchBox,
  ComboBox,
  IComboBoxOption,
  Checkbox,
  DefaultButton,
} from '@fluentui/react';
import {
  ChevronDown16Regular,
  ChevronUp16Regular,
  Dismiss12Regular,
  DocumentText20Regular,
  CheckmarkCircle20Regular,
  CalendarLtr20Regular,
  Filter20Regular,
  Person20Regular,
  TextT20Regular,
  Search20Regular
} from '@fluentui/react-icons';
import { Button } from "@fluentui/react-components";
import '../../components/styles/global.css';
import { getListData } from '../../../../Services/GeneralDocument';
import { getDataByLibraryName } from '../../../../Services/MasTileService';
import { getConfigActive } from '../../../../Services/ConfigService';
import Select from 'react-select';


export interface DynamicFilterConfig {
  key: string;
  label: string;
  columnType: string;
  options?: { key: string; text: string; }[];
}

export interface ActiveFilter {
  key: string;
  label: string;
  value: any;
  displayValue: string;
}

interface SearchFiltersProps {
  context: any;
  siteUrl: string;
  libraryName: string;

  searchQuery: string;
  onSearchChange: (query: string) => void;
  onSearch: () => void;
  onConfigLoaded?: (dynamicControl: any[], filters: DynamicFilterConfig[]) => void;
  activeFilters: ActiveFilter[];
  onFilterChange: (key: string, value: any, displayValue: string) => void;
  onRemoveFilter: (key: string) => void;
  onClearFilters: () => void;
}


const COLUMN_TYPE_ICONS: Record<string, any> = {
  'Dropdown': DocumentText20Regular,
  'Multiple Select': DocumentText20Regular,
  'Date and Time': CalendarLtr20Regular,
  'Single line of Text': TextT20Regular,
  'Multiple lines of Text': TextT20Regular,
  'Person or Group': Person20Regular,
  'Radio': CheckmarkCircle20Regular,
};

export default function SearchFilters({
  context,
  siteUrl,
  libraryName,
  searchQuery,
  onSearchChange,
  onSearch,
  onConfigLoaded,
  activeFilters,
  onFilterChange,
  onRemoveFilter,
  onClearFilters,
}: SearchFiltersProps) {

  const [dynamicControl, setDynamicControl] = useState<any[]>([]);
  const [dynamicFilters, setDynamicFilters] = useState<DynamicFilterConfig[]>([]);
  const [options, setOptions] = useState<{ [key: string]: { value: string; label: string; }[]; }>({});
  const [configLoading, setConfigLoading] = useState<boolean>(true);
  const [expandedSections, setExpandedSections] = useState<Set<string>>(new Set());

  useEffect(() => {
    if (!context || !libraryName) {
      setConfigLoading(false);
      return;
    }
    loadConfig();
  }, [libraryName]);

  const escapeODataUrl = (url: string): string => {
    return url.replace(/'/g, "''");
  };

  const getAllFolders = async (siteUrl: string, libraryName: string): Promise<any[]> => {
    const safeLibrary = escapeODataUrl(libraryName);

    const response = await fetch(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${safeLibrary}')/Folders?$filter=Name ne 'Forms'`,
      {
        headers: { Accept: "application/json;odata=nometadata" }
      }
    );

    const data = await response.json();
    return data.value || [];
  };

  const checkFolderHasDocuments = async (
    siteUrl: string,
    folderRelativeUrl: string
  ): Promise<boolean> => {
    const safeUrl = escapeODataUrl(folderRelativeUrl);

    const response = await fetch(
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${safeUrl}')/Files?$top=1`,
      {
        headers: { Accept: "application/json;odata=nometadata" }
      }
    );

    const data = await response.json();
    return data.value && data.value.length > 0;
  };

  const loadConfig = async () => {
    setConfigLoading(true);
    try {
      const libraryData = await getDataByLibraryName(siteUrl, context.spHttpClient, libraryName);
      console.log('libraryData ', libraryData);
      if (!libraryData?.value?.length) {
        setConfigLoading(false);
        return;
      }

      const rawDynamicControl: any[] = JSON.parse(
        libraryData.value[0].DynamicControl || '[]'
      );
      console.log('rawDynamicControl ', rawDynamicControl);
      setDynamicControl(rawDynamicControl);

      const configData = await getConfigActive(siteUrl, context.spHttpClient);
      console.log('configData', configData);
      const configItems: any[] = configData?.value || [];
      const filterConfigs: DynamicFilterConfig[] = [];
      console.log('filterConfigs', filterConfigs);
      for (const item of rawDynamicControl) {
        if (!item.IsShowAsFilter) continue;
        const configRow = configItems.find((c: any) => c.Id === item.Id);
        if (!configRow) continue;

        filterConfigs.push({
          key: item.InternalTitleName,
          label: item.Title,
          columnType: configRow.ColumnType,
        });
      }

      setDynamicFilters(filterConfigs);
      setExpandedSections(new Set(filterConfigs.map(f => f.key)));
      await bindDropdown(rawDynamicControl, configItems, filterConfigs);
      onConfigLoaded?.(rawDynamicControl, filterConfigs);

    } catch (err) {
      console.error('SearchFilters: error loading config', err);
    } finally {
      setConfigLoading(false);
    }
  };

  const bindDropdown = async (
    controlArr: any[],
    configItems: any[],
    filterConfigs: DynamicFilterConfig[]
  ) => {
    for (const item of controlArr) {
      if (item.ColumnType !== 'Dropdown' && item.ColumnType !== 'Multiple Select') continue;

      const configRow = configItems.find((c: any) => c.Id === item.Id);

      if (!configRow) continue;

      let dropdownOptions: { value: string; label: string; }[] = [];

      if (configRow.IsStaticValue && configRow.StaticDataObject) {
        dropdownOptions = configRow.StaticDataObject
          .split(';')
          .filter(Boolean)
          .map((ele: string) => ({ value: ele, label: ele }));
      } else if (configRow.InternalListName) {
        console.log('Fetching list ', configRow.InternalListName);
        console.log('Display field ', configRow.DisplayValue);
        try {
          const data = await getListData(
            `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${configRow.InternalListName}')/items?$top=5000&$filter=Active eq 1&$orderby=${configRow.DisplayValue} asc`,
            context
          );
          console.log('Dropdown API data ', data);
          dropdownOptions = (data?.value || []).map((ele: any) => ({
            value: ele[configRow.DisplayValue],
            label: ele[configRow.DisplayValue],
          }));
          console.log('dropdownOptions ', dropdownOptions);
        } catch (e) {
          console.warn(`Could not load options for ${item.InternalTitleName}`, e);
        }
      }

      if (item.ColumnType === 'Radio' && configRow.IsStaticValue && configRow.StaticDataObject) {
        dropdownOptions = configRow.StaticDataObject
          .split(';')
          .filter(Boolean)
          .map((ele: string) => ({ value: ele, label: ele }));
      }

      setOptions(prev => {
        const updated = { ...prev, [item.InternalTitleName]: dropdownOptions };
        console.log('Updated options state 👉', updated);
        return updated;
      });

      const fc = filterConfigs.find(f => f.key === item.InternalTitleName);
      if (fc) {
        fc.options = dropdownOptions.map(o => ({ key: o.value, text: o.label }));
      }
    }
    setDynamicFilters(prev =>
      prev.map(f => {
        const updated = filterConfigs.find(fc => fc.key === f.key);
        return updated ? { ...f, options: updated.options } : f;
      })
    );
  };

  const toggleSection = (key: string) => {
    setExpandedSections(prev => {
      const newSet = new Set(prev);
      if (newSet.has(key)) newSet.delete(key);
      else newSet.add(key);
      return newSet;
    });
  };

  const handleContentSearch = async () => {
    if (!searchQuery?.trim()) return;

    const query = searchQuery.trim() || "*";

    const allFolders = await getAllFolders(siteUrl, libraryName);
    const foldersWithDocs: any[] = [];

    for (const folder of allFolders) {
      const hasDocs = await checkFolderHasDocuments(
        siteUrl,
        folder.ServerRelativeUrl
      );

      if (hasDocs) {
        foldersWithDocs.push(folder);
      }
    }

    if (foldersWithDocs.length > 0) {
      const routePath = `${context.pageContext.web.absoluteUrl}/SitePages/Search.aspx?env=WebViewList&q=${encodeURIComponent(
        query
      )}&Library=${encodeURIComponent(libraryName)}`;

      window.open(routePath, "_blank");
    }
  };


  const renderFilter = (filter: DynamicFilterConfig) => {
    const colType = filter.columnType;

    if (colType === 'Dropdown' || colType === 'Multiple Select') {
      const selectOptions =
        (options[filter.key] || []).map(opt => ({
          value: opt.value,
          label: opt.label,
        }));

      const selectedValue = activeFilters.find(f => f.key === filter.key);

      return (
        <div className="filter-section-body">
          <Select
            options={selectOptions}
            value={
              selectedValue
                ? { value: selectedValue.value, label: selectedValue.displayValue }
                : null
            }
            onChange={(option: any) => {
              if (option) {
                onFilterChange(filter.key, option.value, option.label);
              } else {
                onRemoveFilter(filter.key);
              }
            }}
            isSearchable
            isClearable
            placeholder={`Select ${filter.label.toLowerCase()}`}
            classNamePrefix="react-select"
          />
        </div>
      );
    }
    if (colType === 'Radio') {
      const radioOpts = options[filter.key] || [];
      return (
        <div className="filter-section-body filter-checkbox-group">
          {radioOpts.map(opt => (
            <Checkbox
              key={opt.value}
              label={opt.label}
              checked={activeFilters.find(f => f.key === filter.key)?.value === opt.value}
              onChange={(_, checked) => {
                if (checked) onFilterChange(filter.key, opt.value, opt.label);
                else onRemoveFilter(filter.key);
              }}
              className="filter-checkbox-item"
              data-testid={`filter-${filter.key}-${opt.value}`}
            />
          ))}
        </div>
      );
    }

    if (colType === 'Date and Time') {

      const existing = activeFilters.find(f => f.key === filter.key) || { value: {} };
      return (
        <div className="filter-section-body">
          <label style={{ fontSize: '11px', color: '#605e5c', display: 'block', marginBottom: '4px' }}>
            From
          </label>
          <input
            type="date"
            value={
              existing?.value?.from
                ? new Date(existing.value.from).toISOString().split('T')[0]
                : ''
            }
            style={{
              width: '100%',
              padding: '6px 8px',
              border: '1px solid #c8c6c4',
              borderRadius: '4px',
              fontSize: '13px',
              color: '#323130',
              backgroundColor: '#fff',
              boxSizing: 'border-box',
              cursor: 'pointer',
            }}
            onChange={(e) => {
              const date = e.target.value ? new Date(e.target.value) : null;
              if (date) {
                // const toDate = existing?.value?.to || null;
                const toDate = existing?.value?.to ?? null;
                onFilterChange(
                  filter.key,
                  { from: date, to: toDate },
                  `${date.toLocaleDateString('en-GB')}${toDate ? ' → ' + new Date(toDate).toLocaleDateString('en-GB') : ''}`
                );
              }
            }}
            data-testid={`filter-${filter.key}-from`}
          />

          <label style={{ fontSize: '11px', color: '#605e5c', display: 'block', margin: '8px 0 4px' }}>
            To
          </label>
          <input
            type="date"
            value={
              existing?.value?.to
                ? new Date(existing.value.to).toISOString().split('T')[0]
                : ''
            }
            style={{
              width: '100%',
              padding: '6px 8px',
              border: '1px solid #c8c6c4',
              borderRadius: '4px',
              fontSize: '13px',
              color: '#323130',
              backgroundColor: '#fff',
              boxSizing: 'border-box',
              cursor: 'pointer',
            }}
            onChange={(e) => {
              const date = e.target.value ? new Date(e.target.value) : null;
              if (date) {
                const fromDate = existing?.value?.from ?? null;
                onFilterChange(
                  filter.key,
                  { from: fromDate, to: date },
                  `${fromDate ? new Date(fromDate).toLocaleDateString('en-GB') + ' → ' : ''}${date.toLocaleDateString('en-GB')}`
                );
              }
            }}
            data-testid={`filter-${filter.key}-to`}
          />
        </div>
      );
    }


    if (colType === 'Single line of Text' || colType === 'Multiple lines of Text') {
      return (
        <div className="filter-section-body">
          <input
            type="text"
            placeholder={`Filter by ${filter.label.toLowerCase()}...`}
            value={activeFilters.find(f => f.key === filter.key)?.value || ''}
            style={{
              width: '100%',
              padding: '6px 8px',
              border: '1px solid #c8c6c4',
              borderRadius: '4px',
              fontSize: '13px',
              color: '#323130',
              backgroundColor: '#fff',
              boxSizing: 'border-box',
            }}
            onChange={(e) => {
              const val = e.target.value;

              const clean = val.replace(/[^a-zA-Z0-9\s]/g, '');
              if (clean) onFilterChange(filter.key, clean, clean);
              else onRemoveFilter(filter.key);
            }}
            data-testid={`filter-${filter.key}`}
          />
        </div>
      );
    }

    if (colType === 'Person or Group') {
      return (
        <div className="filter-section-body">
          <input
            type="text"
            placeholder={`Search by ${filter.label.toLowerCase()}...`}
            value={activeFilters.find(f => f.key === filter.key)?.value || ''}
            style={{
              width: '100%',
              padding: '6px 8px',
              border: '1px solid #c8c6c4',
              borderRadius: '4px',
              fontSize: '13px',
              color: '#323130',
              backgroundColor: '#fff',
              boxSizing: 'border-box',
            }}
            onChange={(e) => {
              const val = e.target.value;
              if (val.trim()) onFilterChange(filter.key, val, val);
              else onRemoveFilter(filter.key);
            }}
            data-testid={`filter-${filter.key}`}
          />
        </div>
      );
    }

    return null;
  };


  return (
    <div className="search-filters" data-testid="container-search-filters">

      <div className="search-filters-header">
        <Filter20Regular className="search-filters-header-icon" />
        <span>Filters</span>
      </div>

      <div className="search-bar" style={{ display: "flex", gap: "8px", alignItems: "center" }}>
        <SearchBox
          placeholder="Search documents..."
          value={searchQuery}
          onChange={(_, value) => onSearchChange(value || '')}
          onSearch={handleContentSearch}
          className="search-bar-input"
          data-testid="input-search"
          styles={{
            root: { flex: 1 },
          }}
        />
        <Button
          onClick={handleContentSearch}
          icon={<Search20Regular />}
          appearance='primary'
        />
      </div>

      {activeFilters.length > 0 && (
        <div className="filter-chips-section">
          <div className="filter-chips" data-testid="container-active-filters">
            {activeFilters.map(filter => (
              <div key={filter.key} className="filter-chip">
                <span className="filter-chip-text">
                  {filter.label}: {filter.displayValue}
                </span>
                <span
                  className="filter-chip-remove"
                  onClick={() => onRemoveFilter(filter.key)}
                  data-testid={`button-remove-filter-${filter.key}`}
                >
                  <Dismiss12Regular />
                </span>
              </div>
            ))}
          </div>
          <DefaultButton
            className="filter-clear-btn"
            onClick={onClearFilters}
            data-testid="button-clear-filters"
          >
            Clear all
          </DefaultButton>
        </div>
      )}

      {configLoading ? (
        <div style={{ padding: '16px', fontSize: '13px', color: '#605e5c' }}>
          Loading filters...
        </div>
      ) : (
        <div className="filter-sections">
          {dynamicFilters.map(filter => {
            const SectionIcon = COLUMN_TYPE_ICONS[filter.columnType] || Filter20Regular;
            const isExpanded = expandedSections.has(filter.key);
            return (
              <div
                key={filter.key}
                className={`filter-section ${isExpanded ? 'filter-section-expanded' : ''}`}
              >
                <div
                  className="filter-section-title"
                  onClick={() => toggleSection(filter.key)}
                  data-testid={`toggle-filter-section-${filter.key}`}
                >
                  <div className="filter-section-title-left">
                    <SectionIcon className="filter-section-icon" />
                    <span>{filter.label}</span>
                  </div>
                  {isExpanded ? (
                    <ChevronUp16Regular className="filter-section-chevron" />
                  ) : (
                    <ChevronDown16Regular className="filter-section-chevron" />
                  )}
                </div>
                {isExpanded && renderFilter(filter)}
              </div>
            );
          })}
        </div>
      )}

      <span
        onClick={onSearch}
        data-testid="button-apply-filters"
        style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          width: '100%',
          height: '40px',
          borderRadius: '6px',
          backgroundColor: '#0078d4',
          color: '#ffffff',
          fontWeight: 600,
          fontSize: '14px',
          cursor: 'pointer',
          userSelect: 'none',
          marginTop: '12px',
        }}
      >
        Apply Filters
      </span>
    </div>
  );
}