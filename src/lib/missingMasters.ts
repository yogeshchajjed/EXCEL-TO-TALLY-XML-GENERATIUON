import { XMLBuilder } from 'fast-xml-parser';

export interface MissingLedgerItem {
  id: string; // unique id for list rendering
  name: string; // the typed ledger name
  proposedGroup: string; // e.g. "Suspense"
  sourceRowOrVoucher: string; // row number or voucher ID/No
  type: string; // e.g., "Party", "Sales", "GST", etc.
  action: 'Create' | 'Ignore' | 'Replace';
  replacementName: string; // selected existing ledger
  possibleMatches: string[];
}

export interface MissingStockItem {
  id: string;
  name: string;
  proposedStockGroup: string; // e.g. "Suspense Stock Group"
  unit: string; // unit name
  hsn: string; // HSN/SAC code
  gstRate: number; // GST Rate %
  sourceRowOrVoucher: string;
  action: 'Create' | 'Ignore' | 'Replace';
  replacementName: string; // selected existing stock item
  possibleMatches: string[];
  unitMatches: string[]; // unit dropdown
}

export function normalizeName(name: string): string {
  return (name || '').trim().replace(/\s+/g, ' ');
}

export function getLevenshteinDistance(a: string, b: string): number {
  const matrix: number[][] = [];
  for (let i = 0; i <= b.length; i++) {
    matrix[i] = [i];
  }
  for (let j = 0; j <= a.length; j++) {
    matrix[0][j] = j;
  }
  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // substitution
          matrix[i][j - 1] + 1,     // insertion
          matrix[i - 1][j] + 1      // deletion
        );
      }
    }
  }
  return matrix[b.length][a.length];
}

export function suggestExistingMasterMatch(name: string, masters: string[]): { exactMatch: string | null; possibleMatches: string[] } {
  const normalizedTyped = normalizeName(name);
  const typedLower = normalizedTyped.toLowerCase();

  if (!typedLower) {
    return { exactMatch: null, possibleMatches: [] };
  }

  // 1. Check exact case-insensitive match
  const exact = masters.find(m => normalizeName(m).toLowerCase() === typedLower);
  if (exact) {
    return { exactMatch: exact, possibleMatches: [] };
  }

  // 2. Fuzzy checks (Levenshtein distance <= 3 or substring matching)
  const possible: string[] = [];
  for (const master of masters) {
    const normalizedMaster = normalizeName(master);
    const masterLower = normalizedMaster.toLowerCase();

    // Skip trivial empty masters
    if (!masterLower) continue;

    const distance = getLevenshteinDistance(typedLower, masterLower);
    
    // Similarity criteria: Levenshtein distance <= 3 OR one contains the other
    if (distance <= 3 || typedLower.includes(masterLower) || masterLower.includes(typedLower)) {
      possible.push(master);
    }
  }

  return { exactMatch: null, possibleMatches: possible.slice(0, 5) }; // Limit to top 5 matches
}

// Detection for Ledgers
export function detectMissingLedgers(
  usedLedgers: { name: string; source: string; type: string }[],
  masters: string[]
): MissingLedgerItem[] {
  const missing: MissingLedgerItem[] = [];
  const processed = new Set<string>();

  for (const item of usedLedgers) {
    const rawName = item.name;
    if (!rawName || !rawName.trim()) continue;

    const normalized = normalizeName(rawName);
    const lower = normalized.toLowerCase();

    // Already processed this ledger in this detection run?
    if (processed.has(lower)) {
      // Append source info to existing record if found
      const existing = missing.find(m => m.name.toLowerCase() === lower);
      if (existing && !existing.sourceRowOrVoucher.includes(item.source)) {
        existing.sourceRowOrVoucher += `, ${item.source}`;
      }
      continue;
    }

    processed.add(lower);

    // Check match
    const matchResult = suggestExistingMasterMatch(rawName, masters);
    if (!matchResult.exactMatch) {
      missing.push({
        id: `ledger_${Math.random().toString(36).substr(2, 9)}`,
        name: normalized,
        proposedGroup: 'Suspense',
        sourceRowOrVoucher: item.source,
        type: item.type,
        action: 'Create',
        replacementName: '',
        possibleMatches: matchResult.possibleMatches
      });
    }
  }

  return missing;
}

// Detection for Stock Items
export function detectMissingStockItems(
  usedStockItems: { name: string; source: string; unit?: string; hsn?: string; gstRate?: number }[],
  stockMasters: string[],
  unitMasters: string[] = []
): MissingStockItem[] {
  const missing: MissingStockItem[] = [];
  const processed = new Set<string>();

  for (const item of usedStockItems) {
    const rawName = item.name;
    if (!rawName || !rawName.trim()) continue;

    const normalized = normalizeName(rawName);
    const lower = normalized.toLowerCase();

    if (processed.has(lower)) {
      const existing = missing.find(m => m.name.toLowerCase() === lower);
      if (existing && !existing.sourceRowOrVoucher.includes(item.source)) {
        existing.sourceRowOrVoucher += `, ${item.source}`;
      }
      continue;
    }

    processed.add(lower);

    const matchResult = suggestExistingMasterMatch(rawName, stockMasters);
    if (!matchResult.exactMatch) {
      missing.push({
        id: `stock_${Math.random().toString(36).substr(2, 9)}`,
        name: normalized,
        proposedStockGroup: 'Suspense Stock Group',
        unit: item.unit || '',
        hsn: item.hsn || '',
        gstRate: item.gstRate !== undefined ? item.gstRate : 18,
        sourceRowOrVoucher: item.source,
        action: 'Create',
        replacementName: '',
        possibleMatches: matchResult.possibleMatches,
        unitMatches: unitMasters
      });
    }
  }

  return missing;
}

// Generate Group Master XML
export function generateGroupMasterXML(groupName: string, parentGroup: string = 'Suspense A/c'): string {
  const builder = new XMLBuilder({
    ignoreAttributes: false,
    format: true,
    suppressEmptyNode: true,
  });

  const tallyData = {
    TALLYMESSAGE: {
      '@_xmlns:UDF': 'TallyUDF',
      GROUP: {
        '@_NAME': groupName,
        '@_ACTION': 'Create',
        NAME: groupName,
        PARENT: parentGroup,
        ISSUBLEDGER: 'No',
        ISBILLWISEON: 'No',
        ISCOSTCENTRESON: 'No'
      }
    }
  };

  return builder.build(tallyData);
}

// Generate Unit Master XML
export function generateUnitMasterXML(symbol: string): string {
  const builder = new XMLBuilder({
    ignoreAttributes: false,
    format: true,
    suppressEmptyNode: true,
  });

  const tallyData = {
    TALLYMESSAGE: {
      '@_xmlns:UDF': 'TallyUDF',
      UNIT: {
        '@_NAME': symbol,
        '@_ACTION': 'Create',
        NAME: symbol,
        ISSIMPLEUNIT: 'Yes',
        DECIMALPLACES: 0
      }
    }
  };

  return builder.build(tallyData);
}

// Generate Stock Group XML
export function generateStockGroupMasterXML(groupName: string): string {
  const builder = new XMLBuilder({
    ignoreAttributes: false,
    format: true,
    suppressEmptyNode: true,
  });

  const tallyData = {
    TALLYMESSAGE: {
      '@_xmlns:UDF': 'TallyUDF',
      STOCKGROUP: {
        '@_NAME': groupName,
        '@_ACTION': 'Create',
        'NAME.LIST': {
          NAME: groupName,
        },
        PARENT: 'Primary'
      }
    }
  };

  return builder.build(tallyData);
}

// Generate Missing Ledgers XML
export function generateMissingLedgerMastersXML(missingLedgers: MissingLedgerItem[]): string {
  const builder = new XMLBuilder({
    ignoreAttributes: false,
    format: true,
    suppressEmptyNode: true,
  });

  const itemsToCreate = missingLedgers.filter(l => l.action === 'Create');
  if (itemsToCreate.length === 0) return '';

  const messages = itemsToCreate.map(l => ({
    LEDGER: {
      '@_NAME': l.name,
      '@_ACTION': 'Create',
      NAME: l.name,
      PARENT: l.proposedGroup,
      ISBILLWISEON: 'No',
      ISCOSTCENTRESON: 'No',
      ISINTERESTON: 'No',
      AFFECTSSTOCK: 'No',
      ISGSTAPPLICABLE: 'Not Applicable'
    }
  }));

  return messages.map(msg => builder.build({ TALLYMESSAGE: { '@_xmlns:UDF': 'TallyUDF', ...msg } })).join('\n');
}

// Generate Missing Stock Items XML
export function generateMissingStockItemMastersXML(missingStockItems: MissingStockItem[]): string {
  const builder = new XMLBuilder({
    ignoreAttributes: false,
    format: true,
    suppressEmptyNode: true,
  });

  const itemsToCreate = missingStockItems.filter(i => i.action === 'Create');
  if (itemsToCreate.length === 0) return '';

  const messages = itemsToCreate.map(item => {
    const cgstRate = (item.gstRate / 2).toFixed(2);
    const sgstRate = (item.gstRate / 2).toFixed(2);
    const igstRate = item.gstRate.toFixed(2);

    return {
      STOCKITEM: {
        '@_NAME': item.name,
        '@_ACTION': 'Create',
        NAME: item.name,
        PARENT: item.proposedStockGroup,
        BASEUNITS: item.unit,
        GSTAPPLICABLE: 'Applicable',
        GSTTYPEOFSUPPLY: 'Goods',
        GSTHSNNAME: item.hsn || undefined,
        ...(item.gstRate ? {
          'GSTDETAILS.LIST': {
            GSTTAXABILITY: 'Taxable',
            'STATEGSTDETAILS.LIST': {
              'RATEDETAILS.LIST': [
                {
                  GSTRATEDUTYHEAD: 'Integrated Tax',
                  GSTRATE: igstRate
                },
                {
                  GSTRATEDUTYHEAD: 'Central Tax',
                  GSTRATE: cgstRate
                },
                {
                  GSTRATEDUTYHEAD: 'State Tax',
                  GSTRATE: sgstRate
                }
              ]
            }
          }
        } : {})
      }
    };
  });

  return messages.map(msg => builder.build({ TALLYMESSAGE: { '@_xmlns:UDF': 'TallyUDF', ...msg } })).join('\n');
}

// Helper to bundle all XML files in proper Tally sequence
export function generateCombinedImportXML(options: {
  groupsXml?: string;
  unitsXml?: string;
  ledgersXml?: string;
  stockGroupsXml?: string;
  stockItemsXml?: string;
  voucherXml: string;
}): string {
  // Extract tally messages from inner strings, or wrap them if they are standalone
  const parts: string[] = [];

  const cleanXml = (xmlStr?: string) => {
    if (!xmlStr || !xmlStr.trim()) return [];
    // If it contains a full ENVELOPE structure, strip envelope to get raw TALLYMESSAGE nodes
    // Otherwise, split by <TALLYMESSAGE or extract XML content
    const messages: string[] = [];
    const regex = /<TALLYMESSAGE[\s\S]*?<\/TALLYMESSAGE>/gi;
    let match;
    while ((match = regex.exec(xmlStr)) !== null) {
      messages.push(match[0]);
    }
    if (messages.length === 0) {
      // Standalone single element that doesn't have TALLYMESSAGE or might have it
      if (xmlStr.includes('<TALLYMESSAGE')) {
        messages.push(xmlStr);
      } else {
        // Just wrap it if needed or split by lines if it's multiple TALLYMESSAGEs already
        messages.push(xmlStr);
      }
    }
    return messages;
  };

  // Import sequence:
  // 1. Missing Groups
  // 2. Missing Units
  // 3. Missing Ledgers
  // 4. Missing Stock Groups
  // 5. Missing Stock Items
  // 6. Vouchers
  const groups = cleanXml(options.groupsXml);
  const units = cleanXml(options.unitsXml);
  const ledgers = cleanXml(options.ledgersXml);
  const stockGroups = cleanXml(options.stockGroupsXml);
  const stockItems = cleanXml(options.stockItemsXml);
  const vouchers = cleanXml(options.voucherXml);

  parts.push(...groups);
  parts.push(...units);
  parts.push(...ledgers);
  parts.push(...stockGroups);
  parts.push(...stockItems);
  parts.push(...vouchers);

  // Re-build a clean consolidated ENVELOPE
  return `<?xml version="1.0"?>
<ENVELOPE>
  <HEADER>
    <TALLYREQUEST>Import Data</TALLYREQUEST>
  </HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>All Masters</REPORTNAME>
      </REQUESTDESC>
      <REQUESTDATA>
        ${parts.join('\n        ')}
      </REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>`;
}
