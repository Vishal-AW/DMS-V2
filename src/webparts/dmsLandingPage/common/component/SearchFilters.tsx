import * as React from 'react';
import { useState } from 'react';
import {
  SearchBox,
  ComboBox,
  IComboBoxOption,
  DatePicker,
  Checkbox,
  DefaultButton,
  PrimaryButton,
} from '@fluentui/react';
import {
  ChevronDown16Regular,
  ChevronUp16Regular,
  Dismiss12Regular,
  DocumentText20Regular,
  CheckmarkCircle20Regular,
  Scan20Regular,
  CalendarLtr20Regular,
  Building20Regular,
  Filter20Regular,
} from '@fluentui/react-icons';

export interface FilterConfig {
  key: string;
  label: string;
  type: 'dropdown' | 'date' | 'checkbox' | 'dateRange';
  options?: { key: string; text: string; }[];
}

export interface ActiveFilter {
  key: string;
  label: string;
  value: any;
  displayValue: string;
}

interface SearchFiltersProps {
  searchQuery: string;
  onSearchChange: (query: string) => void;
  onSearch: () => void;
  filters: FilterConfig[];
  activeFilters: ActiveFilter[];
  onFilterChange: (key: string, value: any, displayValue: string) => void;
  onRemoveFilter: (key: string) => void;
  onClearFilters: () => void;
}

const sectionIcons: Record<string, typeof Filter20Regular> = {
  documentType: DocumentText20Regular,
  status: CheckmarkCircle20Regular,
  ocrStatus: Scan20Regular,
  modifiedDate: CalendarLtr20Regular,
  workspace: Building20Regular,
};

export default function SearchFilters({
  searchQuery,
  onSearchChange,
  onSearch,
  filters,
  activeFilters,
  onFilterChange,
  onRemoveFilter,
  onClearFilters,
}: SearchFiltersProps) {
  const [expandedSections, setExpandedSections] = useState<Set<string>>(new Set(filters.map(f => f.key)));

  const toggleSection = (key: string) => {
    setExpandedSections(prev => {
      const newSet = new Set(prev);
      if (newSet.has(key)) {
        newSet.delete(key);
      } else {
        newSet.add(key);
      }
      return newSet;
    });
  };

  const renderFilter = (filter: FilterConfig) => {
    switch (filter.type) {
      case 'dropdown':
        const options: IComboBoxOption[] = filter.options?.map(opt => ({
          key: opt.key,
          text: opt.text,
        })) || [];

        return (
          <div className="filter-section-body">
            <ComboBox
              options={options}
              onChange={(_, option) => {
                if (option) {
                  onFilterChange(filter.key, option.key, option.text);
                }
              }}
              allowFreeform
              autoComplete="on"
              placeholder={`Select ${filter.label.toLowerCase()}`}
              data-testid={`filter-${filter.key}`}
            />
          </div>
        );

      case 'date':
        return (
          <div className="filter-section-body">
            <DatePicker
              onSelectDate={(date) => {
                if (date) {
                  onFilterChange(filter.key, date, date.toLocaleDateString());
                }
              }}
              placeholder="Select date"
              data-testid={`filter-${filter.key}`}
            />
          </div>
        );

      case 'checkbox':
        return (
          <div className="filter-section-body filter-checkbox-group">
            {filter.options?.map(opt => (
              <Checkbox
                key={opt.key}
                label={opt.text}
                onChange={(_, checked) => {
                  if (checked) {
                    onFilterChange(filter.key, opt.key, opt.text);
                  } else {
                    onRemoveFilter(`${filter.key}:${opt.key}`);
                  }
                }}
                className="filter-checkbox-item"
                data-testid={`filter-${filter.key}-${opt.key}`}
              />
            ))}
          </div>
        );

      default:
        return null;
    }
  };

  return (
    <div className="search-filters" data-testid="container-search-filters">
      <div className="search-filters-header">
        <Filter20Regular className="search-filters-header-icon" />
        <span>Filters</span>
      </div>

      <div className="search-bar">
        <SearchBox
          placeholder="Search documents..."
          value={searchQuery}
          onChange={(_, value) => onSearchChange(value || '')}
          onSearch={onSearch}
          className="search-bar-input"
          data-testid="input-search"
        />
      </div>

      {activeFilters.length > 0 && (
        <div className="filter-chips-section">
          <div className="filter-chips" data-testid="container-active-filters">
            {activeFilters.map(filter => (
              <div key={filter.key} className="filter-chip">
                <span className="filter-chip-text">{filter.label}: {filter.displayValue}</span>
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

      <div className="filter-sections">
        {filters.map(filter => {
          const SectionIcon = sectionIcons[filter.key] || Filter20Regular;
          const isExpanded = expandedSections.has(filter.key);
          return (
            <div key={filter.key} className={`filter-section ${isExpanded ? 'filter-section-expanded' : ''}`}>
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

      <div className="filter-apply-section">
        <PrimaryButton
          className="filter-apply-btn"
          onClick={onSearch}
          data-testid="button-apply-filters"
        >
          <span>Apply Filters</span>
        </PrimaryButton>
      </div>
    </div>
  );
}
