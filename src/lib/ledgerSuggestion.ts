/**
 * Advanced Transaction-Based Ledger Suggestion Engine
 */

export interface HistoricalMapping {
  narration: string;
  ledger: string;
  source?: 'Daybook History' | 'Direct Tally Daybook' | 'Uploaded Transaction XML' | 'User Correction' | 'Keyword/Fuzzy';
}

export interface SuggestionResult {
  ledgerName: string;
  confidence: number;
  reasoning: string;
  isFromHistory: boolean;
  source: 'Daybook History' | 'Direct Tally Daybook' | 'Uploaded Transaction XML' | 'User Correction' | 'Keyword/Fuzzy';
}

/**
 * Calculates Levenshtein distance between two strings
 */
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

/**
 * Calculates a similarity score between 0.0 and 1.0 based on edit distance
 */
export function getSimilarity(s1: string, s2: string): number {
  const norm1 = s1.toLowerCase().replace(/[^a-z0-9]/g, '');
  const norm2 = s2.toLowerCase().replace(/[^a-z0-9]/g, '');
  if (norm1 === norm2) return 1.0;
  
  const longer = norm1.length > norm2.length ? norm1 : norm2;
  const shorter = norm1.length > norm2.length ? norm2 : norm1;
  
  if (longer.length === 0) return 1.0;
  
  const distance = editDistance(longer, shorter);
  return (longer.length - distance) / longer.length;
}

/**
 * Suggests a Ledger name based on transaction details and history
 */
export function suggestLedgerForTransaction(
  narration: string,
  reference: string,
  amount: number,
  ledgers: string[],
  historicalMappings: HistoricalMapping[] = []
): SuggestionResult {
  if (!narration) {
    return { ledgerName: '', confidence: 0, reasoning: 'No narration provided', isFromHistory: false, source: 'Keyword/Fuzzy' };
  }

  const cleanNarr = narration.toLowerCase().trim();
  const normNarr = cleanNarr.replace(/[^a-z0-9]/g, '');
  const cleanRef = (reference || '').toLowerCase().trim();

  const isValidLedger = (name: string) => ledgers.includes(name);

  // 1. Exact historical mapping match
  if (historicalMappings && historicalMappings.length > 0) {
    const exactHist = historicalMappings.find(h => h.narration.toLowerCase().trim() === cleanNarr);
    if (exactHist && isValidLedger(exactHist.ledger)) {
      return {
        ledgerName: exactHist.ledger,
        confidence: 1.0,
        reasoning: 'Matched exact historical narration from Tally history',
        isFromHistory: true,
        source: exactHist.source || 'Daybook History'
      };
    }
  }

  // 2. Normalized historical match (ignores spaces/symbols)
  if (historicalMappings && historicalMappings.length > 0) {
    const normHist = historicalMappings.find(h => h.narration.toLowerCase().replace(/[^a-z0-9]/g, '') === normNarr);
    if (normHist && isValidLedger(normHist.ledger)) {
      return {
        ledgerName: normHist.ledger,
        confidence: 0.98,
        reasoning: 'Matched normalized historical transaction narration',
        isFromHistory: true,
        source: normHist.source || 'Daybook History'
      };
    }
  }

  // 3. Exact ledger name found in narration
  const sortedLedgers = [...ledgers].sort((a, b) => b.length - a.length);
  for (const ledger of sortedLedgers) {
    const ledgerLower = ledger.toLowerCase().trim();
    if (ledgerLower.length > 3 && cleanNarr.includes(ledgerLower)) {
      return {
        ledgerName: ledger,
        confidence: 0.95,
        reasoning: `Found exact ledger name "${ledger}" in transaction narration`,
        isFromHistory: false,
        source: 'Keyword/Fuzzy'
      };
    }
  }

  // 4. Fuzzy historical match (Levenshtein similarity >= 0.8)
  if (historicalMappings && historicalMappings.length > 0) {
    let bestMatch: { ledger: string; score: number; source?: any } | null = null;
    for (const h of historicalMappings) {
      const score = getSimilarity(h.narration, narration);
      if (score >= 0.8) {
        if (!bestMatch || score > bestMatch.score) {
          bestMatch = { ledger: h.ledger, score, source: h.source };
        }
      }
    }
    if (bestMatch && isValidLedger(bestMatch.ledger)) {
      return {
        ledgerName: bestMatch.ledger,
        confidence: Math.round(bestMatch.score * 100) / 100,
        reasoning: `Fuzzy matched historical narration with ${(bestMatch.score * 100).toFixed(0)}% similarity`,
        isFromHistory: true,
        source: bestMatch.source || 'Daybook History'
      };
    }
  }

  // 5. Intelligent Transaction Pattern & Keyword Rules
  const keywords: { words: string[]; ledgerKeywords: string[]; fallbackLedger: string; confidence: number; reason: string }[] = [
    {
      words: ['salary', 'salaries', 'wages', 'wage'],
      ledgerKeywords: ['salary', 'salaries', 'wages', 'staff welfare'],
      fallbackLedger: 'Salary Account',
      confidence: 0.85,
      reason: 'Narration indicates employee salary or wages payment'
    },
    {
      words: ['rent', 'rental', 'lease'],
      ledgerKeywords: ['rent', 'rental', 'office rent', 'building rent'],
      fallbackLedger: 'Rent Account',
      confidence: 0.85,
      reason: 'Narration indicates rent or lease payment'
    },
    {
      words: ['interest', 'int received', 'int. received', 'fd interest', 'fd int'],
      ledgerKeywords: ['interest received', 'interest earned', 'fd interest', 'other income'],
      fallbackLedger: 'Interest Received',
      confidence: 0.80,
      reason: 'Narration indicates interest earned / FD interest'
    },
    {
      words: ['int paid', 'interest charges', 'interest paid'],
      ledgerKeywords: ['interest paid', 'interest on loan', 'finance costs'],
      fallbackLedger: 'Interest Paid',
      confidence: 0.80,
      reason: 'Narration indicates interest paid / finance charges'
    },
    {
      words: ['charges', 'bank charges', 'chg', 'commission', 'processing fee', 'annual fee', 'sms charges'],
      ledgerKeywords: ['bank charges', 'charges', 'bank expense'],
      fallbackLedger: 'Bank Charges',
      confidence: 0.85,
      reason: 'Narration indicates bank fee, commission, or service charges'
    },
    {
      words: ['gst', 'cgst', 'sgst', 'igst', 'tax', 'tds', 'income tax', 'advance tax'],
      ledgerKeywords: ['gst', 'cgst', 'sgst', 'igst', 'tds', 'taxes & duties', 'income tax'],
      fallbackLedger: 'Taxes & Duties',
      confidence: 0.80,
      reason: 'Narration indicates tax, GST, or TDS entry'
    },
    {
      words: ['cash withdrawal', 'cash w/d', 'withdrawn cash', 'self cash'],
      ledgerKeywords: ['cash', 'cash in hand', 'petty cash'],
      fallbackLedger: 'Cash',
      confidence: 0.90,
      reason: 'Narration indicates cash withdrawal from bank ATM or self'
    },
    {
      words: ['cash deposit', 'deposited cash', 'cash dep'],
      ledgerKeywords: ['cash', 'cash in hand', 'petty cash'],
      fallbackLedger: 'Cash',
      confidence: 0.90,
      reason: 'Narration indicates cash deposit into bank account'
    },
    {
      words: ['electricity', 'power', 'eb bill', 'electric'],
      ledgerKeywords: ['electricity', 'power & fuel', 'eb charges', 'office expenses'],
      fallbackLedger: 'Electricity Expenses',
      confidence: 0.85,
      reason: 'Narration indicates utility / electricity bill payment'
    },
    {
      words: ['telephone', 'mobile', 'internet', 'broadband', 'wifi', 'telecom'],
      ledgerKeywords: ['telephone', 'mobile expenses', 'internet expenses', 'office expenses'],
      fallbackLedger: 'Telephone & Internet Expenses',
      confidence: 0.85,
      reason: 'Narration indicates telecom or internet bill payment'
    }
  ];

  for (const rule of keywords) {
    if (rule.words.some(word => cleanNarr.includes(word))) {
      // Find closest existing ledger matching rule's ledgerKeywords
      for (const lKw of rule.ledgerKeywords) {
        const matchingLedger = ledgers.find(l => l.toLowerCase().includes(lKw));
        if (matchingLedger) {
          return {
            ledgerName: matchingLedger,
            confidence: rule.confidence,
            reasoning: rule.reason + ` (matched existing ledger "${matchingLedger}")`,
            isFromHistory: false,
            source: 'Keyword/Fuzzy'
          };
        }
      }
      // If no close ledger exists, try finding a ledger that exactly/partially matches fallbackLedger
      const matchingFallback = ledgers.find(l => l.toLowerCase().includes(rule.fallbackLedger.toLowerCase()));
      if (matchingFallback) {
        return {
          ledgerName: matchingFallback,
          confidence: rule.confidence - 0.05,
          reasoning: rule.reason + ` (suggested close ledger "${matchingFallback}")`,
          isFromHistory: false,
          source: 'Keyword/Fuzzy'
        };
      }
    }
  }

  // 6. Split words fallback lookup
  for (const ledger of sortedLedgers) {
    const ledgerWords = ledger.toLowerCase().split(/[^a-zA-Z0-9]/).filter(w => w.length > 3);
    if (ledgerWords.length > 0 && ledgerWords.every(w => cleanNarr.includes(w))) {
      return {
        ledgerName: ledger,
        confidence: 0.75,
        reasoning: `All descriptive words of ledger "${ledger}" found in transaction narration`,
        isFromHistory: false,
        source: 'Keyword/Fuzzy'
      };
    }
  }

  // 7. Suspense fallback
  const suspenseLedger = ledgers.find(l => /^suspense$/i.test(l) || /^suspense account$/i.test(l));
  if (suspenseLedger) {
    return {
      ledgerName: suspenseLedger,
      confidence: 0.30,
      reasoning: 'No clear mapping pattern found; suggested Suspense Account for review.',
      isFromHistory: false,
      source: 'Keyword/Fuzzy'
    };
  }

  return {
    ledgerName: ledgers[0] || '',
    confidence: 0.10,
    reasoning: 'Default fallback (first available ledger)',
    isFromHistory: false,
    source: 'Keyword/Fuzzy'
  };
}
