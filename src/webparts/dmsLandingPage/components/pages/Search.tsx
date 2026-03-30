/* eslint-disable */
import * as React from "react";
import { useState, useMemo } from 'react';
import { useNavigate, useLocation } from 'react-router-dom';
import {
  ArrowLeft20Regular,
  Search20Regular,
  DocumentSearch20Regular,
} from '@fluentui/react-icons';
import { DefaultButton, MessageBar, MessageBarType } from '@fluentui/react';
import SearchFilters, { DynamicFilterConfig, ActiveFilter } from '../../common/component/SearchFilters';
import ReactTableComponent from "../ResuableComponents/ReusableDataTable";
import type { ColDef } from "ag-grid-community";
import { getDocument } from "../../../../Services/GeneralDocument";

interface SearchProps {
  context: any;
}

const formatCellValue = (value: any, columnType: string): string => {
  if (value === null || value === undefined || value === '') return '—';

  if (columnType === 'Date and Time') {
    if (!value) return '—';

    const date = new Date(value);
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();

    return `${day}/${month}/${year}`;
  }

  if (columnType === 'Person or Group') {
    if (Array.isArray(value)) {
      return value.map((v: any) => v?.Title || v).join(', ');
    }
    if (typeof value === 'object') return value?.Title || '—';
  }

  return String(value);
};

export default function Search({ context }: SearchProps) {
  const navigate = useNavigate();
  const location = useLocation();
  const returnPath = (location.state as any)?.from || '/';

  const libName: string = (location.state as any)?.libName
    || sessionStorage.getItem('LibName')
    || '';


  const [searchQuery, setSearchQuery] = useState('');
  const [activeFilters, setActiveFilters] = useState<ActiveFilter[]>([]);
  const [hasSearched, setHasSearched] = useState(false);
  const [searchResults, setSearchResults] = useState<any[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [errorMessage, setErrorMessage] = useState('');


  const [dynamicControl, setDynamicControl] = useState<any[]>([]);
  const [dynamicFilters, setDynamicFilters] = useState<DynamicFilterConfig[]>([]);

  const siteUrl: string = context?.pageContext?.web?.absoluteUrl || '';


  const handleConfigLoaded = (control: any[], filters: DynamicFilterConfig[]) => {
    setDynamicControl(control);
    setDynamicFilters(filters);
  };


  const dynamicColumnDefs = useMemo((): ColDef[] => {
    const baseCols: ColDef[] = [
      {
        field: 'srNo',
        headerName: 'SR.NO',
        maxWidth: 80,
        valueGetter: (params: any) => (params.node?.rowIndex ?? 0) + 1,
      },
      {
        field: 'ActualName',
        headerName: 'File Name',
        minWidth: 220,
        cellRenderer: (params: any) => {
          const name: string = params.value || 'N/A';
          const fileUrl: string = params.data?.fileUrl || '';

          if (fileUrl && name !== 'N/A') {
            return (
              <span
                style={{
                  color: '#0078d4',
                  cursor: 'pointer'
                }}
                onClick={(e) => {
                  e.preventDefault();
                  e.stopPropagation();
                  window.open(fileUrl, '_blank', 'noopener,noreferrer');
                }}
              >
                {name}
              </span>
            );
          }

          return name;
        },
      },
      {
        field: 'FolderDocumentPath',
        headerName: 'Folder Path',
        minWidth: 200,
        valueFormatter: (params: any) => params.value || '—',
      },
    ];


    const dynCols: ColDef[] = dynamicControl
      .filter((field: any) => field.IsShowAsFilter)
      .map((field: any) => ({
        field: field.InternalTitleName,
        headerName: field.Title,
        minWidth: 140,
        valueFormatter: (params: any) =>
          formatCellValue(params.value, field.ColumnType),
      }));

    return [...baseCols, ...dynCols];
  }, [dynamicControl]);


  const fetchSearchResults = async () => {
    if (!context) {
      setErrorMessage('SharePoint context is not available.');
      return;
    }

    setIsLoading(true);
    setErrorMessage('');

    try {

      let filter = "InternalStatus eq 'Published' and Active eq 1";
      if (searchQuery.trim()) {
        filter += ` and substringof('${encodeURIComponent(searchQuery.trim())}', ActualName)`;
      }

      for (const af of activeFilters) {
        const controlItem = dynamicControl.find(
          (c: any) => c.InternalTitleName === af.key
        );
        const colType = controlItem?.ColumnType || 'Single line of Text';

        if (colType === 'Date and Time') {

          if (af.value?.from instanceof Date) {
            const from = new Date(af.value.from);
            from.setHours(0, 0, 0, 0);
            filter += ` and ${af.key} ge datetime'${from.toISOString()}'`;
          }
          if (af.value?.to instanceof Date) {
            const to = new Date(af.value.to);
            to.setHours(23, 59, 59, 999);
            filter += ` and ${af.key} le datetime'${to.toISOString()}'`;
          }
        } else if (colType === 'Person or Group') {

          filter += ` and ${af.key}/Title eq '${encodeURIComponent(af.value)}'`;
        } else {

          filter += ` and ${af.key} eq '${encodeURIComponent(af.value)}'`;
        }
      }

      const response = await getDocument(siteUrl, context.spHttpClient, filter, libName);
      const dataArr: any[] = response?.value || [];

      const mapped = dataArr
        .filter((el: any) => {
          return (
            el.File?.ServerRelativeUrl !== undefined &&
            el.DisplayStatus !== 'Pending With Approver' &&
            el.DisplayStatus !== 'Rejected'
          );
        })
        .map((el: any, index: number) => {

          const row: Record<string, any> = {
            srNo: index + 1,
            ActualName: el.ActualName || el.FileLeafRef || 'N/A',
            fileUrl: el.File?.ServerRelativeUrl || '',
            FolderDocumentPath: el.FolderDocumentPath || '—',
            _displayStatus: el.DisplayStatus,
            _internalStatus: el.InternalStatus,
          };

          dynamicControl.forEach((field: any) => {
            row[field.InternalTitleName] = el[field.InternalTitleName];
          });

          return row;
        });

      setSearchResults(mapped);
    } catch (error) {
      console.error('Search: error fetching results', error);
      setErrorMessage('An error occurred while searching. Please try again.');
    } finally {
      setIsLoading(false);
    }
  };

  const handleSearch = () => {
    setHasSearched(true);
    fetchSearchResults();
  };

  const handleFilterChange = (key: string, value: any, displayValue: string) => {
    const config = dynamicFilters.find(f => f.key === key);
    if (!config) return;
    setActiveFilters(prev => [
      ...prev.filter(f => f.key !== key),
      { key, label: config.label, value, displayValue },
    ]);
  };

  const handleRemoveFilter = (key: string) => {
    setActiveFilters(prev => prev.filter(f => f.key !== key));
  };

  const handleClearFilters = () => {
    setActiveFilters([]);
  };

  return (
    <div className="search-page" data-testid="page-search">


      <div className="search-topbar">
        <DefaultButton
          className="search-back-btn"
          onClick={() => navigate(returnPath)}
          data-testid="link-back-dashboard"
        >
          <ArrowLeft20Regular className="search-back-icon" />
          <span>Go Back</span>
        </DefaultButton>
        <div className="search-topbar-title">
          <DocumentSearch20Regular className="search-topbar-title-icon" />
          <span>Advanced Search</span>
        </div>
      </div>

      <div className="search-body">

        <div className="search-sidebar">
          <SearchFilters
            context={context}
            siteUrl={siteUrl}
            libraryName={libName}
            searchQuery={searchQuery}
            onSearchChange={setSearchQuery}
            onSearch={handleSearch}
            onConfigLoaded={handleConfigLoaded}
            activeFilters={activeFilters}
            onFilterChange={handleFilterChange}
            onRemoveFilter={handleRemoveFilter}
            onClearFilters={handleClearFilters}
          />
        </div>

        <div className="search-results">

          {isLoading && (
            <div className="empty-state">
              <div className="empty-state-icon">
                <Search20Regular className="empty-state-search-icon" />
              </div>
              <h2 className="empty-state-title">Searching...</h2>
              <p className="empty-state-description">
                Fetching documents. Please wait.
              </p>
            </div>
          )}

          {!isLoading && errorMessage && (
            <MessageBar
              messageBarType={MessageBarType.error}
              className="search-results-bar"
            >
              {errorMessage}
            </MessageBar>
          )}


          {!isLoading && !errorMessage && !hasSearched && (
            <div className="empty-state">
              <div className="empty-state-icon">
                <Search20Regular className="empty-state-search-icon" />
              </div>
              <h2 className="empty-state-title" data-testid="text-empty-search-title">
                Start Your Search
              </h2>
              <p className="empty-state-description" data-testid="text-empty-search-desc">
                Enter keywords in the search box or use the filters on the left
                to find documents across the library.
              </p>
            </div>
          )}

          {!isLoading && !errorMessage && hasSearched && searchResults.length === 0 && (
            <div className="empty-state">
              <div className="empty-state-icon">
                <Search20Regular className="empty-state-search-icon" />
              </div>
              <h2 className="empty-state-title" data-testid="text-no-results-title">
                No Results Found
              </h2>
              <p className="empty-state-description" data-testid="text-no-results-desc">
                Try adjusting your search terms or filters to find what you're looking for.
              </p>
            </div>
          )}


          {!isLoading && !errorMessage && hasSearched && searchResults.length > 0 && (
            <>
              <MessageBar
                messageBarType={MessageBarType.info}
                className="search-results-bar"
                data-testid="info-results-count"
              >
                Found {searchResults.length} document(s) matching your search criteria.
              </MessageBar>

              <div style={{ padding: '20px' }}>
                <ReactTableComponent
                  rowData={searchResults}
                  columnDefs={dynamicColumnDefs}
                  pagination={true}
                />
              </div>
            </>
          )}

        </div>
      </div>
    </div>
  );
}