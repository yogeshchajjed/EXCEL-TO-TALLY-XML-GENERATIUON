import React, { useState, useEffect, useRef } from 'react';
import { Search, Plus, Check } from 'lucide-react';

interface SearchableComboboxProps {
  value: string;
  onChange: (val: string) => void;
  options: string[];
  disabled?: boolean;
  placeholder?: string;
  className?: string;
}

// Calculates Levenshtein distance between two strings
function editDistance(s1: string, s2: string): number {
  const costs = [];
  for (let i = 0; i <= s1.length; i++) {
    let lastValue = i;
    for (let j = 0; j <= s2.length; j++) {
      if (i === 0) {
        costs[j] = j;
      } else {
        if (j > 0) {
          let newValue = costs[j - 1];
          if (s1.charAt(i - 1) !== s2.charAt(j - 1)) {
            newValue = Math.min(Math.min(newValue, lastValue), costs[j]) + 1;
          }
          costs[j - 1] = lastValue;
          lastValue = newValue;
        }
      }
    }
    if (i > 0) costs[s2.length] = lastValue;
  }
  return costs[s2.length];
}

// Calculates a similarity score between 0.0 and 1.0 based on edit distance
function getSimilarity(s1: string, s2: string): number {
  const norm1 = s1.toLowerCase().replace(/[^a-z0-9]/g, '');
  const norm2 = s2.toLowerCase().replace(/[^a-z0-9]/g, '');
  if (norm1 === norm2) return 1.0;
  
  const longer = norm1.length > norm2.length ? norm1 : norm2;
  const shorter = norm1.length > norm2.length ? norm2 : norm1;
  
  if (longer.length === 0) return 1.0;
  
  const distance = editDistance(longer, shorter);
  return (longer.length - distance) / longer.length;
}

export const SearchableCombobox: React.FC<SearchableComboboxProps> = ({
  value,
  onChange,
  options = [],
  disabled = false,
  placeholder = "Type or select...",
  className = ""
}) => {
  const [isOpen, setIsOpen] = useState(false);
  const [search, setSearch] = useState(value);
  const [highlightedIndex, setHighlightedIndex] = useState(-1);
  const containerRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  // Keep internal search input in sync with external value
  useEffect(() => {
    setSearch(value);
  }, [value]);

  // Filter & rank options based on search query with a 5-tier algorithm
  const filteredOptions = React.useMemo(() => {
    const query = search.toLowerCase().trim();
    if (!query) return options;

    const queryNoSpace = query.replace(/\s+/g, '');

    const scored = options.map(opt => {
      const optLower = opt.toLowerCase();
      const optNoSpace = optLower.replace(/\s+/g, '');
      
      let score = 0;

      if (optLower === query) {
        score = 100; // Tier 1: exact match first
      } else if (optLower.startsWith(query)) {
        score = 80;  // Tier 2: startsWith second
      } else if (optLower.includes(query)) {
        score = 60;  // Tier 3: contains third
      } else if (optNoSpace === queryNoSpace) {
        score = 50;  // Tier 4: normalized-space exact match fourth
      } else if (optNoSpace.startsWith(queryNoSpace) || optNoSpace.includes(queryNoSpace)) {
        score = 40;  // Tier 4 part b: normalized-space contains/startsWith
      } else {
        // Tier 5: Levenshtein fuzzy match
        const sim = getSimilarity(optLower, query);
        if (sim >= 0.3) {
          score = 10 + Math.round(sim * 25); // Score between 17 and 35
        }
      }

      return { option: opt, score };
    });

    return scored
      .filter(item => item.score > 0)
      .sort((a, b) => b.score - a.score)
      .map(item => item.option);
  }, [options, search]);

  const showCustomOption = search.trim() && !options.some(opt => opt.toLowerCase() === search.toLowerCase().trim());

  // Close dropdown on click outside
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setIsOpen(false);
        // If the user typed a custom value and closed, propagate it
        if (search !== value) {
          onChange(search);
        }
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [search, value, onChange]);

  const handleSelect = (val: string) => {
    onChange(val);
    setSearch(val);
    setIsOpen(false);
    setHighlightedIndex(-1);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (disabled) return;

    if (!isOpen) {
      if (e.key === 'ArrowDown' || e.key === 'Enter') {
        setIsOpen(true);
        e.preventDefault();
      }
      return;
    }

    const totalItems = filteredOptions.length + (showCustomOption ? 1 : 0);

    switch (e.key) {
      case 'ArrowDown':
        setHighlightedIndex(prev => (prev + 1) % totalItems);
        e.preventDefault();
        break;
      case 'ArrowUp':
        setHighlightedIndex(prev => (prev - 1 + totalItems) % totalItems);
        e.preventDefault();
        break;
      case 'Enter':
        if (highlightedIndex >= 0 && highlightedIndex < filteredOptions.length) {
          handleSelect(filteredOptions[highlightedIndex]);
        } else if (showCustomOption && highlightedIndex === filteredOptions.length) {
          handleSelect(search);
        } else {
          handleSelect(search);
        }
        e.preventDefault();
        break;
      case 'Escape':
        setIsOpen(false);
        setSearch(value); // revert
        setHighlightedIndex(-1);
        e.preventDefault();
        break;
      case 'Tab':
        // commit current search value on blur
        onChange(search);
        setIsOpen(false);
        break;
    }
  };

  return (
    <div ref={containerRef} className={`relative w-full ${className}`}>
      <div className="relative flex items-center">
        <input
          ref={inputRef}
          type="text"
          value={search}
          onChange={(e) => {
            const val = e.target.value;
            setSearch(val);
            onChange(val); // Propagate value as they type
            setIsOpen(true);
            setHighlightedIndex(-1);
          }}
          onFocus={() => setIsOpen(true)}
          onKeyDown={handleKeyDown}
          disabled={disabled}
          placeholder={placeholder}
          className="w-full pl-2 pr-8 py-1.5 bg-white border border-zinc-200 rounded-lg text-xs font-semibold text-zinc-800 outline-none focus:ring-1 focus:ring-zinc-950 disabled:opacity-50 transition shadow-sm"
        />
        <div className="absolute right-2.5 top-1/2 -translate-y-1/2 text-zinc-400 pointer-events-none">
          <Search className="w-3.5 h-3.5" />
        </div>
      </div>

      {isOpen && !disabled && (
        <div className="absolute left-0 right-0 mt-1 max-h-52 overflow-y-auto bg-white border border-zinc-200 rounded-lg shadow-lg z-50 divide-y divide-zinc-50 scrollbar-thin">
          {filteredOptions.length > 0 ? (
            filteredOptions.map((opt, idx) => {
              const isSelected = opt === value;
              const isHighlighted = idx === highlightedIndex;
              return (
                <div
                  key={opt}
                  onClick={() => handleSelect(opt)}
                  className={`px-3 py-1.5 text-xs font-medium cursor-pointer flex items-center justify-between ${
                    isHighlighted ? 'bg-zinc-50 text-zinc-900' : 'text-zinc-700 hover:bg-zinc-50/50'
                  }`}
                >
                  <span className="truncate">{opt}</span>
                  {isSelected && <Check className="w-3.5 h-3.5 text-emerald-600 shrink-0 ml-2" />}
                </div>
              );
            })
          ) : (
            !showCustomOption && (
              <div className="px-3 py-2 text-xs text-zinc-400 italic">No matches found</div>
            )
          )}

          {showCustomOption && (
            <div
              onClick={() => handleSelect(search)}
              className={`px-3 py-1.5 text-xs font-bold cursor-pointer flex items-center text-zinc-900 gap-1.5 ${
                highlightedIndex === filteredOptions.length ? 'bg-zinc-50' : 'hover:bg-zinc-50/50'
              }`}
            >
              <Plus className="w-3.5 h-3.5 text-zinc-500" />
              <span className="truncate">Use custom: "{search}"</span>
            </div>
          )}
        </div>
      )}
    </div>
  );
};
