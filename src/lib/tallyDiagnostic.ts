import ExcelJS from 'exceljs';

interface TallyContextShape {
  ledgers: string[];
  groups: string[];
  stockGroups?: string[];
  stockItems?: string[];
  units?: string[];
  ledgerGroupMap?: Record<string, string>;
  stockItemStockGroupMap?: Record<string, string>;
  groupParentMap?: Record<string, string>;
  historicalMappings?: { narration: string; ledger: string }[];
  transactions?: any[];
  parseErrors?: string[];
}

export function extractTransactionsFromXmlObj(obj: any): { transactions: any[], parseErrors: string[] } {
  const transactions: any[] = [];
  const parseErrors: string[] = [];

  const traverse = (current: any) => {
    if (!current) return;
    
    if (current.VOUCHER) {
      const list = Array.isArray(current.VOUCHER) ? current.VOUCHER : [current.VOUCHER];
      list.forEach((v: any, index: number) => {
        try {
          const rawDate = v.DATE || v.Date || '';
          
          const voucherType = v.VOUCHERTYPENAME || v['@_VCHTYPE'] || '';
          const narration = v.NARRATION || '';
          const reference = v.REFERENCE || '';
          const voucherNo = v.VOUCHERNUMBER || '';

          const entriesList = v.ALLLEDGERENTRIES_LIST || v['ALLLEDGERENTRIES.LIST'] || v.LEDGERENTRIES_LIST || v['LEDGERENTRIES.LIST'] || v.ALLLEDGERENTRIES || v.LEDGERENTRIES;
          if (entriesList) {
            const entries = Array.isArray(entriesList) ? entriesList : [entriesList];
            entries.forEach((ent: any) => {
              const ledger = ent.LEDGERNAME || ent.LedgerName || '';
              const amountVal = ent.AMOUNT !== undefined ? parseFloat(String(ent.AMOUNT)) : 0;
              if (ledger) {
                transactions.push({
                  date: String(rawDate),
                  voucherType: String(voucherType),
                  narration: String(narration),
                  ledger: String(ledger),
                  amount: amountVal,
                  reference: String(reference),
                  voucherNo: String(voucherNo)
                });
              }
            });
          }
        } catch (err: any) {
          parseErrors.push(`Error parsing voucher at index ${index}: ${err.message}`);
        }
      });
    }

    if (Array.isArray(current)) {
      current.forEach(item => traverse(item));
    } else if (typeof current === 'object') {
      Object.values(current).forEach(val => traverse(val));
    }
  };

  traverse(obj);
  return { transactions, parseErrors };
}

function getTransactionKey(tx: any): string {
  const normDate = tx.date ? String(tx.date).replace(/[-/]/g, '') : 'N/A';
  const vType = tx.voucherType ? String(tx.voucherType).trim().toUpperCase() : 'N/A';
  const vNo = tx.voucherNo ? String(tx.voucherNo).trim().toUpperCase() : '';
  const ledger = tx.ledger ? String(tx.ledger).trim().toUpperCase() : 'N/A';
  const amount = tx.amount !== undefined ? Number(tx.amount).toFixed(2) : '0.00';
  const ref = tx.reference ? String(tx.reference).trim().toUpperCase() : '';
  
  return `${normDate}|${vType}|${vNo}|${ledger}|${amount}|${ref}`;
}

export async function generateDiagnosticReportExcel(
  manual: TallyContextShape | null,
  direct: TallyContextShape | null
): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'TallyGen Pro';
  wb.created = new Date();

  const getCount = (arr: any[] | undefined) => arr ? arr.length : 0;

  const mCountLedgers = getCount(manual?.ledgers);
  const dCountLedgers = getCount(direct?.ledgers);

  const mCountGroups = getCount(manual?.groups);
  const dCountGroups = getCount(direct?.groups);

  const mCountStockItems = getCount(manual?.stockItems);
  const dCountStockItems = getCount(direct?.stockItems);

  const mCountUnits = getCount(manual?.units);
  const dCountUnits = getCount(direct?.units);

  const mCountMappings = getCount(manual?.historicalMappings);
  const dCountMappings = getCount(direct?.historicalMappings);

  const manualTxs = manual?.transactions || [];
  const directTxs = direct?.transactions || [];

  const mCountTxs = getCount(manualTxs);
  const dCountTxs = getCount(directTxs);

  // 1. OVERVIEW SUMMARY SHEET
  const summarySheet = wb.addWorksheet('Overview Summary');
  summarySheet.views = [{ showGridLines: true }];

  // Column headers
  summarySheet.getRow(1).values = ['Diagnostic Parameter', 'Manual Upload Count', 'Direct Tally Fetch Count', 'Variance', 'Status / Recommendation'];
  summarySheet.columns = [
    { key: 'param', width: 30 },
    { key: 'manual', width: 25 },
    { key: 'direct', width: 25 },
    { key: 'variance', width: 15 },
    { key: 'status', width: 45 }
  ];

  const rows = [
    {
      param: 'Ledger Masters',
      manual: mCountLedgers,
      direct: dCountLedgers,
      variance: dCountLedgers - mCountLedgers,
      status: dCountLedgers === mCountLedgers ? 'PASSED: Perfect Parity' : 'MISMATCH: Check skipped masters or active filters'
    },
    {
      param: 'Account Groups',
      manual: mCountGroups,
      direct: dCountGroups,
      variance: dCountGroups - mCountGroups,
      status: dCountGroups === mCountGroups ? 'PASSED: Perfect Parity' : 'MISMATCH: Verify missing groups'
    },
    {
      param: 'Stock Items',
      manual: mCountStockItems,
      direct: dCountStockItems,
      variance: dCountStockItems - mCountStockItems,
      status: dCountStockItems === mCountStockItems ? 'PASSED: Perfect Parity' : 'MISMATCH: Verify stock mapping configurations'
    },
    {
      param: 'Measurement Units',
      manual: mCountUnits,
      direct: dCountUnits,
      variance: dCountUnits - mCountUnits,
      status: dCountUnits === mCountUnits ? 'PASSED: Perfect Parity' : 'MISMATCH: Check unit codes'
    },
    {
      param: 'Historical Mappings',
      manual: mCountMappings,
      direct: dCountMappings,
      variance: dCountMappings - mCountMappings,
      status: dCountMappings === mCountMappings ? 'PASSED: Perfect Parity' : 'INFO: Direct fetch learns from active Daybook range'
    },
    {
      param: 'Daybook Transactions',
      manual: mCountTxs,
      direct: dCountTxs,
      variance: dCountTxs - mCountTxs,
      status: dCountTxs === mCountTxs ? 'PASSED: Perfect Parity' : 'RECONCILE: Verify Daybook range and sync filters'
    }
  ];

  rows.forEach(r => summarySheet.addRow(r));

  // Style overview headers
  const headerRow = summarySheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFF' }, size: 11 };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '18181B' } // Zinc 900
  };
  headerRow.alignment = { horizontal: 'center', vertical: 'middle' };

  // Format variance column and cells
  for (let i = 2; i <= rows.length + 1; i++) {
    const r = summarySheet.getRow(i);
    const varianceVal = r.getCell('variance').value as number;
    if (varianceVal === 0) {
      r.getCell('status').font = { color: { argb: '16A34A' }, bold: true }; // Emerald 600
    } else {
      r.getCell('status').font = { color: { argb: 'EA580C' }, bold: true }; // Orange 600
    }
    r.getCell('variance').font = { bold: true };
    r.getCell('param').font = { bold: true };
  }

  // 2. LEDGERS RECONCILIATION SHEET
  const ledgersSheet = wb.addWorksheet('Ledgers Reconciliation');
  ledgersSheet.views = [{ showGridLines: true }];
  ledgersSheet.getRow(1).values = ['Ledger Name', 'In Manual Upload', 'In Direct Fetch', 'Manual Parent Group', 'Direct Parent Group', 'Group Matches?'];
  ledgersSheet.columns = [
    { key: 'name', width: 35 },
    { key: 'manual', width: 20 },
    { key: 'direct', width: 20 },
    { key: 'mGroup', width: 25 },
    { key: 'dGroup', width: 25 },
    { key: 'groupMatch', width: 18 }
  ];

  const ledgerHeader = ledgersSheet.getRow(1);
  ledgerHeader.font = { bold: true, color: { argb: 'FFFFFF' } };
  ledgerHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F2937' } };

  const allLedgers = Array.from(new Set([
    ...(manual?.ledgers || []),
    ...(direct?.ledgers || [])
  ])).sort();

  allLedgers.forEach(ledger => {
    const inManual = manual?.ledgers?.includes(ledger) ? 'YES' : 'NO';
    const inDirect = direct?.ledgers?.includes(ledger) ? 'YES' : 'NO';
    const manualGroup = manual?.ledgerGroupMap?.[ledger] || 'N/A';
    const directGroup = direct?.ledgerGroupMap?.[ledger] || 'N/A';
    const groupMatches = manualGroup === directGroup ? 'YES' : (manualGroup !== 'N/A' && directGroup !== 'N/A' ? 'NO (Mismatch)' : 'N/A');

    ledgersSheet.addRow({
      name: ledger,
      manual: inManual,
      direct: inDirect,
      mGroup: manualGroup,
      dGroup: directGroup,
      groupMatch: groupMatches
    });
  });

  for (let r = 2; r <= allLedgers.length + 1; r++) {
    const row = ledgersSheet.getRow(r);
    const manualVal = row.getCell('manual').value;
    const directVal = row.getCell('direct').value;
    const gMatch = row.getCell('groupMatch').value;

    if (manualVal === 'NO' || directVal === 'NO') {
      row.getCell('name').font = { color: { argb: 'DC2626' }, bold: true };
    }
    if (gMatch === 'NO (Mismatch)') {
      row.getCell('groupMatch').font = { color: { argb: 'DC2626' }, bold: true };
    }
  }

  // 3. STOCK ITEMS RECONCILIATION SHEET
  const stockSheet = wb.addWorksheet('Stock Items Reconciliation');
  stockSheet.views = [{ showGridLines: true }];
  stockSheet.getRow(1).values = ['Stock Item Name', 'In Manual Upload', 'In Direct Fetch', 'Manual Group', 'Direct Group'];
  stockSheet.columns = [
    { key: 'name', width: 35 },
    { key: 'manual', width: 20 },
    { key: 'direct', width: 20 },
    { key: 'mGroup', width: 25 },
    { key: 'dGroup', width: 25 }
  ];

  const stockHeader = stockSheet.getRow(1);
  stockHeader.font = { bold: true, color: { argb: 'FFFFFF' } };
  stockHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F2937' } };

  const allStock = Array.from(new Set([
    ...(manual?.stockItems || []),
    ...(direct?.stockItems || [])
  ])).sort();

  allStock.forEach(item => {
    const inManual = manual?.stockItems?.includes(item) ? 'YES' : 'NO';
    const inDirect = direct?.stockItems?.includes(item) ? 'YES' : 'NO';
    const manualGroup = manual?.stockItemStockGroupMap?.[item] || 'N/A';
    const directGroup = direct?.stockItemStockGroupMap?.[item] || 'N/A';

    stockSheet.addRow({
      name: item,
      manual: inManual,
      direct: inDirect,
      mGroup: manualGroup,
      dGroup: directGroup
    });
  });

  // PREPARE TRANSACTION LISTS WITH KEYS
  const manualMap = new Map<string, any>();
  const directMap = new Map<string, any>();

  manualTxs.forEach(tx => {
    manualMap.set(getTransactionKey(tx), tx);
  });

  directTxs.forEach(tx => {
    directMap.set(getTransactionKey(tx), tx);
  });

  const allTxKeys = Array.from(new Set([
    ...Array.from(manualMap.keys()),
    ...Array.from(directMap.keys())
  ]));

  // 4. TRANSACTIONS RECONCILIATION SHEET
  const txSheet = wb.addWorksheet('Transactions Reconciliation');
  txSheet.views = [{ showGridLines: true }];
  txSheet.getRow(1).values = ['Date', 'Voucher Type', 'Voucher No', 'Ledger', 'Amount', 'Reference', 'In Manual?', 'In Direct?', 'Reconciliation Status'];
  txSheet.columns = [
    { key: 'date', width: 12 },
    { key: 'vtype', width: 15 },
    { key: 'vno', width: 15 },
    { key: 'ledger', width: 30 },
    { key: 'amount', width: 15 },
    { key: 'ref', width: 15 },
    { key: 'manual', width: 15 },
    { key: 'direct', width: 15 },
    { key: 'status', width: 25 }
  ];

  const txHeader = txSheet.getRow(1);
  txHeader.font = { bold: true, color: { argb: 'FFFFFF' } };
  txHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F2937' } };

  allTxKeys.forEach(key => {
    const mTx = manualMap.get(key);
    const dTx = directMap.get(key);
    const tx = mTx || dTx;

    const inManual = mTx ? 'YES' : 'NO';
    const inDirect = dTx ? 'YES' : 'NO';
    
    let status = 'MATCHED';
    if (!mTx) status = 'EXTRA IN DIRECT FETCH';
    if (!dTx) status = 'MISSING IN DIRECT FETCH';

    txSheet.addRow({
      date: tx.date,
      vtype: tx.voucherType,
      vno: tx.voucherNo,
      ledger: tx.ledger,
      amount: tx.amount,
      ref: tx.reference,
      manual: inManual,
      direct: inDirect,
      status: status
    });
  });

  for (let r = 2; r <= allTxKeys.length + 1; r++) {
    const row = txSheet.getRow(r);
    const statusVal = row.getCell('status').value as string;
    if (statusVal === 'MISSING IN DIRECT FETCH') {
      row.getCell('status').font = { color: { argb: 'DC2626' }, bold: true };
    } else if (statusVal === 'EXTRA IN DIRECT FETCH') {
      row.getCell('status').font = { color: { argb: '2563EB' }, bold: true };
    } else {
      row.getCell('status').font = { color: { argb: '16A34A' }, bold: true };
    }
  }

  // 5. MISSING TRANSACTIONS IN DIRECT FETCH
  const missingSheet = wb.addWorksheet('Missing in Direct Fetch');
  missingSheet.views = [{ showGridLines: true }];
  missingSheet.getRow(1).values = ['Date', 'Voucher Type', 'Voucher No', 'Ledger', 'Amount', 'Reference', 'Primary Reason / Troubleshooting'];
  missingSheet.columns = [
    { key: 'date', width: 12 },
    { key: 'vtype', width: 15 },
    { key: 'vno', width: 15 },
    { key: 'ledger', width: 30 },
    { key: 'amount', width: 15 },
    { key: 'ref', width: 15 },
    { key: 'reason', width: 50 }
  ];

  const missingHeader = missingSheet.getRow(1);
  missingHeader.font = { bold: true, color: { argb: 'FFFFFF' } };
  missingHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'B91C1C' } };

  allTxKeys.forEach(key => {
    const mTx = manualMap.get(key);
    const dTx = directMap.get(key);
    if (mTx && !dTx) {
      let reason = 'Voucher range filter, active Company mismatch, or specific Voucher Type configuration exclusion';
      if (mTx.date && direct && direct.transactions && direct.transactions.length > 0) {
        const dDates = direct.transactions.map(t => String(t.date).replace(/[-/]/g, ''));
        const mDateNorm = String(mTx.date).replace(/[-/]/g, '');
        const validDates = dDates.map(d => parseInt(d) || 0).filter(d => d > 0);
        if (validDates.length > 0) {
          const maxDDate = Math.max(...validDates);
          const minDDate = Math.min(...validDates);
          const mDateInt = parseInt(mDateNorm) || 0;
          if (mDateInt && (mDateInt < minDDate || mDateInt > maxDDate)) {
            reason = `Date Out Of Scope: Transaction date ${mTx.date} lies outside Direct Tally Fetch range (${minDDate} to ${maxDDate})`;
          }
        }
      }
      missingSheet.addRow({
        date: mTx.date,
        vtype: mTx.voucherType,
        vno: mTx.voucherNo,
        ledger: mTx.ledger,
        amount: mTx.amount,
        ref: mTx.reference,
        reason: reason
      });
    }
  });

  // 6. EXTRA TRANSACTIONS IN DIRECT FETCH
  const extraSheet = wb.addWorksheet('Extra in Direct Fetch');
  extraSheet.views = [{ showGridLines: true }];
  extraSheet.getRow(1).values = ['Date', 'Voucher Type', 'Voucher No', 'Ledger', 'Amount', 'Reference', 'Primary Reason / Troubleshooting'];
  extraSheet.columns = [
    { key: 'date', width: 12 },
    { key: 'vtype', width: 15 },
    { key: 'vno', width: 15 },
    { key: 'ledger', width: 30 },
    { key: 'amount', width: 15 },
    { key: 'ref', width: 15 },
    { key: 'reason', width: 50 }
  ];

  const extraHeader = extraSheet.getRow(1);
  extraHeader.font = { bold: true, color: { argb: 'FFFFFF' } };
  extraHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1D4ED8' } };

  allTxKeys.forEach(key => {
    const mTx = manualMap.get(key);
    const dTx = directMap.get(key);
    if (!mTx && dTx) {
      extraSheet.addRow({
        date: dTx.date,
        vtype: dTx.voucherType,
        vno: dTx.voucherNo,
        ledger: dTx.ledger,
        amount: dTx.amount,
        ref: dTx.reference,
        reason: 'Export parameters discrepancy: Direct Fetch has loaded a broader date range or includes secondary Daybook voucher logs'
      });
    }
  });

  // 7. TRANSACTION PARSE ERRORS SHEET
  const errorSheet = wb.addWorksheet('Transaction Parse Errors');
  errorSheet.views = [{ showGridLines: true }];
  errorSheet.getRow(1).values = ['Discrepancy Source', 'Error Description / XML Node Context'];
  errorSheet.columns = [
    { key: 'source', width: 25 },
    { key: 'desc', width: 85 }
  ];

  const errHeader = errorSheet.getRow(1);
  errHeader.font = { bold: true, color: { argb: 'FFFFFF' } };
  errHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '374151' } };

  let hasErrors = false;
  if (manual?.parseErrors && manual.parseErrors.length > 0) {
    hasErrors = true;
    manual.parseErrors.forEach(err => {
      errorSheet.addRow({ source: 'Manual XML Upload', desc: err });
    });
  }
  if (direct?.parseErrors && direct.parseErrors.length > 0) {
    hasErrors = true;
    direct.parseErrors.forEach(err => {
      errorSheet.addRow({ source: 'Direct Tally Fetch', desc: err });
    });
  }
  if (!hasErrors) {
    errorSheet.addRow({ source: 'N/A', desc: 'No transaction parse errors detected in either Manual or Direct feeds.' });
  }

  // Export and Download file
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `Tally_Direct_vs_Manual_Diagnostic_${new Date().toISOString().split('T')[0]}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}
