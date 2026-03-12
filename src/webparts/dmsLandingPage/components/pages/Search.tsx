import * as React from "react";
import { useState, useMemo } from 'react';
import { useNavigate, useLocation } from 'react-router-dom';
import {
  ArrowLeft20Regular,
  Search20Regular,
  DocumentSearch20Regular,
} from '@fluentui/react-icons';
import { DefaultButton, MessageBar, MessageBarType } from '@fluentui/react';
import SearchFilters, { FilterConfig, ActiveFilter } from '../../common/component/SearchFilters';


// todo: remove mock functionality - replace with SharePoint search API


const filterConfigs: FilterConfig[] = [
  {
    key: 'documentType',
    label: 'Document Type',
    type: 'dropdown',
    options: [
      { key: 'pdf', text: 'PDF' },
      { key: 'docx', text: 'Word Document' },
      { key: 'xlsx', text: 'Excel Spreadsheet' },
      { key: 'pptx', text: 'PowerPoint' },
      { key: 'png', text: 'PNG Image' },
      { key: 'jpg', text: 'JPEG Image' },
      { key: 'dwg', text: 'AutoCAD Drawing' },
    ],
  },
  {
    key: 'status',
    label: 'Status',
    type: 'checkbox',
    options: [
      { key: 'draft', text: 'Draft' },
      { key: 'pending', text: 'Pending Review' },
      { key: 'approved', text: 'Approved' },
      { key: 'rejected', text: 'Rejected' },
    ],
  },
  {
    key: 'ocrStatus',
    label: 'OCR Status',
    type: 'checkbox',
    options: [
      { key: 'pending', text: 'Pending' },
      { key: 'processing', text: 'Processing' },
      { key: 'completed', text: 'Completed' },
      { key: 'failed', text: 'Failed' },
    ],
  },
  {
    key: 'modifiedDate',
    label: 'Modified Date',
    type: 'date',
  },
  {
    key: 'workspace',
    label: 'Workspace',
    type: 'dropdown',
    options: [
      { key: '1', text: 'Project Documents' },
      { key: '2', text: 'HR Documents' },
      { key: '3', text: 'Financial Records' },
      { key: '4', text: 'Marketing Assets' },
      { key: '5', text: 'Legal Documents' },
      { key: '6', text: 'Engineering Specs' },
    ],
  },
];

export default function Search() {
  const navigate = useNavigate();
  const location = useLocation();
  const returnPath = (location.state as any)?.from || '/';
  const libName = (location.state as any)?.libName || '/';
  const [searchQuery, setSearchQuery] = useState('');
  const [activeFilters, setActiveFilters] = useState<ActiveFilter[]>([]);
  const [hasSearched, setHasSearched] = useState(false);
  const [mockSearchResults, setMockSearchResults] = useState<any[]>([]);
  const filteredResults = useMemo(() => {
    if (!hasSearched) return [];

    let results = [...mockSearchResults];

    if (searchQuery) {
      const query = searchQuery.toLowerCase();
      results = results.filter(doc =>
        doc.name.toLowerCase().includes(query)
      );
    }

    activeFilters.forEach(filter => {
      if (filter.key === 'documentType') {
        results = results.filter(doc => doc.fileType === filter.value);
      } else if (filter.key === 'status') {
        results = results.filter(doc => doc.status === filter.value);
      } else if (filter.key === 'ocrStatus') {
        results = results.filter(doc => doc.ocrStatus === filter.value);
      }
    });

    return results;
  }, [searchQuery, activeFilters, hasSearched]);

  const handleSearch = () => {
    setHasSearched(true);
    console.log('Searching for:', searchQuery, 'with filters:', activeFilters);
  };

  const handleFilterChange = (key: string, value: any, displayValue: string) => {
    const config = filterConfigs.find(f => f.key === key);
    if (config) {
      setActiveFilters(prev => [
        ...prev.filter(f => f.key !== key),
        { key, label: config.label, value, displayValue },
      ]);
    }
  };

  const handleRemoveFilter = (key: string) => {
    setActiveFilters(prev => prev.filter(f => f.key !== key));
  };

  const handleClearFilters = () => {
    setActiveFilters([]);
  };

  const handleDocumentClick = (doc: any) => {
    console.log('Navigate to document:', doc.name);
    // todo: implement navigation to document location
    navigate('/workspace/1');
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
            searchQuery={searchQuery}
            onSearchChange={setSearchQuery}
            onSearch={handleSearch}
            filters={filterConfigs}
            activeFilters={activeFilters}
            onFilterChange={handleFilterChange}
            onRemoveFilter={handleRemoveFilter}
            onClearFilters={handleClearFilters}
          />
        </div>

        <div className="search-results">
          {!hasSearched ? (
            <div className="empty-state">
              <div className="empty-state-icon">
                <Search20Regular className="empty-state-search-icon" />
              </div>
              <h2 className="empty-state-title" data-testid="text-empty-search-title">
                Start Your Search
              </h2>
              <p className="empty-state-description" data-testid="text-empty-search-desc">
                Enter keywords in the search box or use the filters on the left to find documents across all workspaces.
              </p>
            </div>
          ) : filteredResults.length === 0 ? (
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
          ) : (
            <>
              <MessageBar
                messageBarType={MessageBarType.info}
                className="search-results-bar"
                data-testid="info-results-count"
              >
                Found {filteredResults.length} document(s) matching your search criteria.
              </MessageBar>


            </>
          )}
        </div>
      </div>
    </div>
  );
}
