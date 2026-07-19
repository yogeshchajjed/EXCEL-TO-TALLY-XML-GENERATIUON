import React, { useState, useEffect, useRef } from 'react';
import { 
  GoogleAuthProvider, 
  signInWithPopup, 
  signOut, 
  onAuthStateChanged, 
  User 
} from 'firebase/auth';
import { 
  collection, 
  query, 
  where, 
  onSnapshot, 
  addDoc, 
  updateDoc,
  setDoc,
  serverTimestamp,
  doc,
  getDocFromServer,
  deleteDoc
} from 'firebase/firestore';
import { auth, db, handleFirestoreError, OperationType } from './lib/firebase';
import { 
  getAppMode, 
  isElectron, 
  OFFLINE_USER, 
  subscribeToConversions, 
  subscribeToTallyContext, 
  saveTallyContext, 
  deleteTallyContext, 
  saveConversion, 
  updateConversion, 
  clearOfflineWorkspace,
  saveTallyPushLog,
  getTallyPushLogs,
  TallyPushLog
} from './lib/storageAdapter';
import { 
  isDirectTallyAvailable, 
  testTallyConnection, 
  fetchTallyCompany, 
  fetchTallyMasters, 
  fetchTallyDaybook, 
  pushXmlToTally, 
  parseTallyResponse 
} from './lib/tallyConnector';
import { 
  FileSpreadsheet, 
  LogOut, 
  LogIn, 
  History, 
  Download, 
  AlertCircle,
  CheckCircle2,
  Loader2,
  Sparkles,
  ArrowRight,
  ArrowLeft,
  FileCode,
  Database,
  FileUp,
  FileText,
  Search,
  RotateCcw,
  Upload,
  HelpCircle,
  Building,
  Trash2,
  RefreshCw,
  X,
  Check
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { XMLParser } from 'fast-xml-parser';
import * as pdfjsLib from 'pdfjs-dist';
import { getAIColumnMapping, ColumnMapping, mapBankTransactions, BankTransaction, MappedTransaction, parseBankStatementText, isGeminiAvailable } from './services/geminiService';
import { 
  generateTallyXML, 
  generateLedgersXML, 
  generateStockItemsXML, 
  generateStockGroupsXML, 
  generateUnitsXML, 
  normalizeTallyDate,
  getFinalXMLNarration,
  TallyVoucher, 
  TallyLedgerMaster, 
  TallyStockItemMaster, 
  TallyStockGroupMaster, 
  TallyUnitMaster,
  SalesPurchaseInvoice,
  generateSalesPurchaseXML
} from './lib/tallyXml';
import { parseSalesPurchaseExcel } from './lib/salesPurchaseParser';
import {
  MissingLedgerItem,
  MissingStockItem,
  detectMissingLedgers,
  detectMissingStockItems,
  generateGroupMasterXML,
  generateUnitMasterXML,
  generateStockGroupMasterXML,
  generateMissingLedgerMastersXML,
  generateMissingStockItemMastersXML,
  generateCombinedImportXML
} from './lib/missingMasters';

// Set up PDF.js worker
if (getAppMode() !== 'desktop-offline') {
  pdfjsLib.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.mjs`;
} else {
  pdfjsLib.GlobalWorkerOptions.workerSrc = '';
}

// --- Types ---
interface ConversionRecord {
  id: string;
  fileName: string;
  status: 'pending' | 'processing' | 'completed' | 'failed';
  timestamp: any;
  error?: string;
  xmlContent?: string;
  voucherType?: string;
}

interface TallyContext {
  ledgers: string[];
  groups: string[];
  stockGroups?: string[];
  stockItems?: string[];
  units?: string[];
  ledgerGroupMap?: Record<string, string>; // Maps ledger name to its parent group
  stockItemStockGroupMap?: Record<string, string>; // Stock item to stock group mapping
  historicalMappings?: { narration: string; ledger: string }[];
  groupParentMap?: Record<string, string>; // Maps group name to parent group name
  ledgerDetails?: LedgerMasterRow[];
  stockItemDetails?: StockMasterRow[];
}

interface LedgerMasterRow {
  rowNum: number;
  ledgerName: string;
  underGroup: string;
  openingBalance?: string;
  drCr?: string;
  mailingName?: string;
  address1?: string;
  address2?: string;
  state?: string;
  country?: string;
  pincode?: string;
  pan?: string;
  gstin?: string;
  registrationType?: string;
  taxability?: string;
  isBillwiseOn?: string;
  isCostCentreOn?: string;
  email?: string;
  mobileNumber?: string;
  
  isValid: boolean;
  errors: string[];
  warnings: string[];
  isDuplicate: boolean;
  isPossibleDuplicate: boolean;
  duplicateMessage?: string;
  excluded: boolean;
}

interface StockMasterRow {
  rowNum: number;
  itemName: string;
  underGroup: string;
  unit: string;
  openingQty?: string;
  openingRate?: string;
  openingValue?: string;
  hsn?: string;
  gstApplicable?: string;
  taxability?: string;
  gstRate?: string;
  cgstRate?: string;
  sgstRate?: string;
  igstRate?: string;
  description?: string;
  
  isValid: boolean;
  errors: string[];
  warnings: string[];
  isDuplicate: boolean;
  isPossibleDuplicate: boolean;
  duplicateMessage?: string;
  excluded: boolean;
}

interface StockGroupRow {
  rowNum: number;
  groupName: string;
  underGroup?: string;
  isValid: boolean;
  errors: string[];
  excluded: boolean;
}

interface UnitRow {
  rowNum: number;
  symbol: string;
  formalName?: string;
  uqc?: string;
  decimalPlaces?: string;
  isValid: boolean;
  errors: string[];
  excluded: boolean;
}

const isBankOrCashGroup = (groupName: string): boolean => {
  const norm = groupName.toLowerCase().trim();
  const bankCashKeywords = [
    'bank accounts',
    'cash-in-hand',
    'bank od a/c',
    'bank occ a/c',
    'bank cc a/c',
    'bank overdraft',
    'cash account',
    'bank account',
    'cash in hand',
    'bank od',
    'bank cc',
    'bank occ',
  ];
  return bankCashKeywords.some(kw => norm === kw || norm.startsWith(kw));
};

const isExcludedGroup = (groupName: string): boolean => {
  const norm = groupName.toLowerCase().trim();
  const excludedKeywords = [
    'indirect expenses',
    'direct expenses',
    'indirect incomes',
    'direct incomes',
    'secured loans',
    'unsecured loans',
    'loans & advances',
    'current liabilities',
    'duties & taxes',
    'sundry debtors',
    'sundry creditors',
    'purchase accounts',
    'sales accounts',
    'fixed assets',
    'deposits',
    'expense',
    'income',
    'loan',
  ];
  return excludedKeywords.some(kw => norm === kw || norm.startsWith(kw) || norm.includes(kw));
};

const isStrictBankCashLedger = (ledgerName: string, context: TallyContext | null): boolean => {
  if (!context) return false;
  
  // Get direct parent group
  let currentGroup = context.ledgerGroupMap?.[ledgerName];
  if (!currentGroup) return false;

  const visited = new Set<string>();
  
  while (currentGroup) {
    const groupLower = currentGroup.toLowerCase().trim();
    
    // Check exclusion first
    if (isExcludedGroup(currentGroup)) {
      return false;
    }

    // Check inclusion
    if (isBankOrCashGroup(currentGroup)) {
      return true;
    }

    // Detect cycle
    if (visited.has(groupLower)) {
      break;
    }
    visited.add(groupLower);

    // Go to parent
    currentGroup = context.groupParentMap?.[currentGroup];
  }

  return false;
};

export interface JournalLine {
  rowNo: number;
  ledgerName: string;
  drCr: 'Dr' | 'Cr';
  amount: number;
  narration?: string;
  reference?: string;
  costCentre?: string;
  billReference?: string;
  remarks?: string;
  excluded?: boolean;
  validationMessage?: string;
}

export interface JournalVoucherGroup {
  voucherNo: string;
  voucherDate: string;
  totalDebit: number;
  totalCredit: number;
  difference: number;
  status: 'Balanced' | 'Error';
  lines: JournalLine[];
  errors: string[];
  warnings: string[];
  isValid: boolean;
}

export default function App() {
  const [user, setUser] = useState<User | null>(null);
  const [loading, setLoading] = useState(true);
  const [conversions, setConversions] = useState<ConversionRecord[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [mappings, setMappings] = useState<ColumnMapping[]>([]);
  const [currentStep, setCurrentStep] = useState<'upload' | 'mapping' | 'complete' | 'context' | 'master-review' | 'bank-statement-review' | 'bank-statement-detection-review' | 'verification' | 'sales-purchase-verification'>('upload');
  const [companyState, setCompanyState] = useState<string>('Maharashtra');
  const [salesPurchaseInvoices, setSalesPurchaseInvoices] = useState<SalesPurchaseInvoice[]>([]);
  const [selectedInvoiceIdx, setSelectedInvoiceIdx] = useState<number>(0);

  const updateActiveInvoice = (updater: (inv: SalesPurchaseInvoice) => void) => {
    if (!salesPurchaseInvoices[selectedInvoiceIdx]) return;
    const updated = [...salesPurchaseInvoices];
    const copy = JSON.parse(JSON.stringify(updated[selectedInvoiceIdx]));
    updater(copy);
    recalculateInvoice(copy, companyState);
    updated[selectedInvoiceIdx] = copy;
    setSalesPurchaseInvoices(updated);
  };

  // --- GST and Invoice Auto calculations helper ---
  const findGSTLedger = (ledgers: string[], voucherType: 'Sales' | 'Purchase', gstType: 'CGST' | 'SGST' | 'IGST', rate: number): string => {
    const isSales = voucherType === 'Sales';
    const rateHalf = gstType === 'IGST' ? rate : (rate / 2);
    const searchTerms = [
      `${gstType.toLowerCase()} ${isSales ? 'output' : 'input'} ${rateHalf}%`,
      `${gstType.toLowerCase()} ${isSales ? 'output' : 'input'} ${rateHalf}`,
      `${gstType.toLowerCase()} ${rateHalf}%`,
      `${gstType.toLowerCase()} ${rateHalf}`,
      `${gstType} ${isSales ? 'Output' : 'Input'}`,
      gstType,
    ];

    for (const term of searchTerms) {
      const found = ledgers.find(l => l.toLowerCase().replace(/\s+/g, ' ').includes(term.toLowerCase()));
      if (found) return found;
    }

    const anyMatch = ledgers.find(l => l.toLowerCase().includes(gstType.toLowerCase()));
    if (anyMatch) return anyMatch;

    return `${gstType} ${isSales ? 'Output' : 'Input'} ${rateHalf}%`;
  };

  interface SalesPurchaseBalancingResult {
    invoiceNo: string;
    debitTotal: number;
    creditTotal: number;
    expectedPartyAmount: number;
    actualPartyAmount: number;
    difference: number;
    isBalanced: boolean;
    ledgerEntries: { ledgerName: string; amount: number; isParty: boolean }[];
  }

  const verifySalesPurchaseBalancing = (inv: SalesPurchaseInvoice): SalesPurchaseBalancingResult => {
    const isSales = inv.voucherType === 'Sales';
    const ledgerEntries: { ledgerName: string; amount: number; isParty: boolean }[] = [];

    // 1. Party ledger (Actual Party Amount)
    const actualPartyAmount = isSales ? -inv.invoiceTotal : inv.invoiceTotal;
    ledgerEntries.push({
      ledgerName: inv.partyLedger || 'Missing Party Ledger',
      amount: actualPartyAmount,
      isParty: true
    });

    // 2. Sales/Purchase Ledger
    const totalTaxable = inv.items.reduce((sum, item) => sum + item.taxableValue, 0);
    const salesPurchaseAmount = isSales ? totalTaxable : -totalTaxable;
    ledgerEntries.push({
      ledgerName: inv.salesPurchaseLedger || (isSales ? 'Sales Account' : 'Purchase Account'),
      amount: salesPurchaseAmount,
      isParty: false
    });

    // 3. GST Ledgers
    const gstSums: Record<string, number> = {};
    inv.items.forEach(item => {
      if (item.cgstLedger && item.cgstAmount) {
        gstSums[item.cgstLedger] = (gstSums[item.cgstLedger] || 0) + item.cgstAmount;
      }
      if (item.sgstLedger && item.sgstAmount) {
        gstSums[item.sgstLedger] = (gstSums[item.sgstLedger] || 0) + item.sgstAmount;
      }
      if (item.igstLedger && item.igstAmount) {
        gstSums[item.igstLedger] = (gstSums[item.igstLedger] || 0) + item.igstAmount;
      }
    });
    Object.entries(gstSums).forEach(([ledgerName, amount]) => {
      ledgerEntries.push({
        ledgerName,
        amount: isSales ? amount : -amount,
        isParty: false
      });
    });

    // 4. Additional Charges
    const charges = [
      { ledger: inv.freightLedger, amount: inv.freightAmount, isDiscount: false },
      { ledger: inv.packingLedger, amount: inv.packingAmount, isDiscount: false },
      { ledger: inv.loadingLedger, amount: inv.loadingAmount, isDiscount: false },
      { ledger: inv.insuranceLedger, amount: inv.insuranceAmount, isDiscount: false },
      { ledger: inv.otherLedger1, amount: inv.otherAmount1, isDiscount: false },
      { ledger: inv.otherLedger2, amount: inv.otherAmount2, isDiscount: false },
      { ledger: inv.discountLedger, amount: inv.billDiscountAmount, isDiscount: true },
      { ledger: inv.roundOffLedger, amount: inv.roundOffAmount, isDiscount: false }
    ];

    charges.forEach(c => {
      if (c.ledger && c.amount !== undefined && c.amount !== 0) {
        let signValue = c.amount;
        if (c.isDiscount) {
          signValue = -c.amount;
        }
        const chargeAmountSign = isSales ? signValue : -signValue;
        ledgerEntries.push({
          ledgerName: c.ledger,
          amount: chargeAmountSign,
          isParty: false
        });
      }
    });

    // Calculate totals and difference
    let debitTotal = 0;
    let creditTotal = 0;
    let othersSum = 0;

    ledgerEntries.forEach(entry => {
      if (entry.amount < 0) {
        debitTotal += Math.abs(entry.amount);
      } else {
        creditTotal += entry.amount;
      }

      if (!entry.isParty) {
        othersSum += entry.amount;
      }
    });

    const expectedPartyAmount = -othersSum;
    const difference = actualPartyAmount - expectedPartyAmount;
    // Round to 2 decimal places to avoid float precision comparison issues
    const isBalanced = Math.abs(Math.round(difference * 100) / 100) < 0.01;

    return {
      invoiceNo: inv.invoiceNo,
      debitTotal: Math.round(debitTotal * 100) / 100,
      creditTotal: Math.round(creditTotal * 100) / 100,
      expectedPartyAmount: Math.round(expectedPartyAmount * 100) / 100,
      actualPartyAmount: Math.round(actualPartyAmount * 100) / 100,
      difference: Math.round(difference * 100) / 100,
      isBalanced,
      ledgerEntries
    };
  };

  const processSalesPurchaseExcel = (jsonData: any[]): SalesPurchaseInvoice[] => {
    return parseSalesPurchaseExcel(jsonData, 'standard_itemwise') as unknown as SalesPurchaseInvoice[];
  };

  const recalculateInvoice = (inv: SalesPurchaseInvoice, currentCompanyState: string) => {
    let totalTaxableValue = 0;
    let totalCGST = 0;
    let totalSGST = 0;
    let totalIGST = 0;

    const ledgersList = tallyContext?.ledgers || [];

    inv.items.forEach(item => {
      totalTaxableValue += item.taxableValue;

      if (inv.gstMode === 'Auto') {
        const isSameState = String(inv.placeOfSupply || '').toLowerCase().trim() === String(currentCompanyState).toLowerCase().trim();
        if (isSameState) {
          item.cgstAmount = item.taxableValue * (item.gstRate / 100) / 2;
          item.sgstAmount = item.taxableValue * (item.gstRate / 100) / 2;
          item.igstAmount = 0;
          item.cgstLedger = findGSTLedger(ledgersList, inv.voucherType, 'CGST', item.gstRate);
          item.sgstLedger = findGSTLedger(ledgersList, inv.voucherType, 'SGST', item.gstRate);
          item.igstLedger = '';
        } else {
          item.cgstAmount = 0;
          item.sgstAmount = 0;
          item.igstAmount = item.taxableValue * (item.gstRate / 100);
          item.cgstLedger = '';
          item.sgstLedger = '';
          item.igstLedger = findGSTLedger(ledgersList, inv.voucherType, 'IGST', item.gstRate);
        }
      }

      totalCGST += item.cgstAmount;
      totalSGST += item.sgstAmount;
      totalIGST += item.igstAmount;
    });

    inv.totalTaxableValue = totalTaxableValue;
    inv.totalCGST = totalCGST;
    inv.totalSGST = totalSGST;
    inv.totalIGST = totalIGST;

    inv.totalAdditionalCharges = inv.freightAmount + inv.packingAmount + inv.loadingAmount + inv.insuranceAmount + inv.otherAmount1 + inv.otherAmount2 - inv.billDiscountAmount;

    const totalBeforeRoundOff = totalTaxableValue + totalCGST + totalSGST + totalIGST + inv.totalAdditionalCharges;

    if (inv.roundOffLedger && inv.roundOffAmount === 0) {
      const rounded = Math.round(totalBeforeRoundOff);
      inv.roundOffAmount = rounded - totalBeforeRoundOff;
    }

    inv.invoiceTotal = totalBeforeRoundOff + inv.roundOffAmount;

    const errors: string[] = [];
    const warnings: string[] = [];

    const dNorm = normalizeTallyDate(inv.invoiceDate);
    if (!dNorm.isValid) {
      errors.push(`Invalid Invoice Date: ${dNorm.error}`);
    } else {
      inv.invoiceDate = dNorm.value;
    }

    if (!inv.invoiceNo) errors.push("Invoice Number is blank.");
    if (!inv.partyLedger) errors.push("Party Ledger is blank.");

    if (tallyContext) {
      const pExists = tallyContext.ledgers.some(l => l.toLowerCase() === inv.partyLedger.toLowerCase());
      if (!pExists) {
        warnings.push(`Party Ledger "${inv.partyLedger}" not found in Tally masters.`);
      }
      const sExists = tallyContext.ledgers.some(l => l.toLowerCase() === inv.salesPurchaseLedger.toLowerCase());
      if (!sExists) {
        warnings.push(`${inv.voucherType} Ledger "${inv.salesPurchaseLedger}" not found in Tally masters.`);
      }

      inv.items.forEach((item, itemIdx) => {
        const isAccountingMode = inv.voucherMode === 'Accounting';
        const itemPrefix = isAccountingMode ? `Ledger Line ${itemIdx + 1}:` : `Item ${itemIdx + 1} (${item.stockItem}):`;
        
        if (!isAccountingMode) {
          const iExists = tallyContext.stockItems?.some(s => s.toLowerCase() === item.stockItem.toLowerCase());
          if (!iExists) {
            warnings.push(`${itemPrefix} Stock Item not found in Tally masters.`);
          }

          // Auto-resolve empty HSN from stock item details in context
          if (!item.hsn && tallyContext?.stockItemDetails) {
            const sMaster = tallyContext.stockItemDetails.find(sm => sm.itemName.toLowerCase() === item.stockItem.toLowerCase());
            if (sMaster && sMaster.hsn) {
              item.hsn = sMaster.hsn;
            }
          }

          if (!item.hsn || !item.hsn.trim()) {
            errors.push(`${itemPrefix} HSN/SAC is blank. A valid HSN/SAC is mandatory for GST reporting.`);
          }
        }

        if (item.cgstAmount > 0 && !item.cgstLedger) {
          errors.push(`${itemPrefix} CGST Ledger required but not selected.`);
        } else if (item.cgstAmount > 0 && !tallyContext.ledgers.some(l => l.toLowerCase() === item.cgstLedger?.toLowerCase())) {
          warnings.push(`${itemPrefix} CGST Ledger "${item.cgstLedger}" not found in Tally masters.`);
        }

        if (item.sgstAmount > 0 && !item.sgstLedger) {
          errors.push(`${itemPrefix} SGST Ledger required but not selected.`);
        } else if (item.sgstAmount > 0 && !tallyContext.ledgers.some(l => l.toLowerCase() === item.sgstLedger?.toLowerCase())) {
          warnings.push(`${itemPrefix} SGST Ledger "${item.sgstLedger}" not found in Tally masters.`);
        }

        if (item.igstAmount > 0 && !item.igstLedger) {
          errors.push(`${itemPrefix} IGST Ledger required but not selected.`);
        } else if (item.igstAmount > 0 && !tallyContext.ledgers.some(l => l.toLowerCase() === item.igstLedger?.toLowerCase())) {
          warnings.push(`${itemPrefix} IGST Ledger "${item.igstLedger}" not found in Tally masters.`);
        }
      });

      if (inv.freightAmount !== 0 && !inv.freightLedger) errors.push("Freight amount entered but ledger is blank.");
      if (inv.packingAmount !== 0 && !inv.packingLedger) errors.push("Packing amount entered but ledger is blank.");
      if (inv.loadingAmount !== 0 && !inv.loadingLedger) errors.push("Loading amount entered but ledger is blank.");
      if (inv.insuranceAmount !== 0 && !inv.insuranceLedger) errors.push("Insurance amount entered but ledger is blank.");
      if (inv.otherAmount1 !== 0 && !inv.otherLedger1) errors.push("Other Amount 1 entered but ledger is blank.");
      if (inv.otherAmount2 !== 0 && !inv.otherLedger2) errors.push("Other Amount 2 entered but ledger is blank.");
      if (inv.billDiscountAmount !== 0 && !inv.discountLedger) errors.push("Bill discount entered but ledger is blank.");
      if (inv.roundOffAmount !== 0 && !inv.roundOffLedger) errors.push("Round Off adjustment exists but ledger is blank.");
    }

    inv.errors = errors;
    inv.warnings = warnings;
    inv.isValid = errors.length === 0;
  };

  const [pendingData, setPendingData] = useState<any[]>([]);
  const [pendingFileName, setPendingFileName] = useState('');
  const [tallyContext, setTallyContext] = useState<TallyContext | null>(null);
  const [isContextLoading, setIsContextLoading] = useState(false);
  const [selectedVoucherType, setSelectedVoucherType] = useState('Payment');
  const [selectedBankLedger, setSelectedBankLedger] = useState('');
  const [aiMappedTransactions, setAiMappedTransactions] = useState<MappedTransaction[]>([]);
  const [verificationRows, setVerificationRows] = useState<any[]>([]);
  const [verificationSourceStep, setVerificationSourceStep] = useState<'bank-statement-review' | 'mapping'>('mapping');
  const [useLedgerAsNarration, setUseLedgerAsNarration] = useState(false);
  const [bankSearchTerm, setBankSearchTerm] = useState('');
  const [bankFilterType, setBankFilterType] = useState('All');
  const [bankShowAttentionRequiredFirst, setBankShowAttentionRequiredFirst] = useState(true);
  const [confirmProceedWithSuspense, setConfirmProceedWithSuspense] = useState(false);
  const [bankCompactRowHeight, setBankCompactRowHeight] = useState(false);
  const [bankStatementValidationErrors, setBankStatementValidationErrors] = useState<{ rowNo: number; issue: string; suggestedFix: string }[]>([]);
  const [bankSuccessMessage, setBankSuccessMessage] = useState<string | null>(null);

  // Direct Tally Connection states
  const [tallyHost, setTallyHost] = useState('localhost');
  const [tallyPort, setTallyPort] = useState(9000);
  const [tallyStatus, setTallyStatus] = useState<'Disconnected' | 'Connecting' | 'Connected' | 'Error'>('Disconnected');
  const [tallyError, setTallyError] = useState('');
  const [tallyCompany, setTallyCompany] = useState('');
  const [tallyFromDate, setTallyFromDate] = useState('2026-04-01');
  const [tallyToDate, setTallyToDate] = useState('2026-07-19');
  const [recentPushLogs, setRecentPushLogs] = useState<TallyPushLog[]>([]);
  const [isPushModalOpen, setIsPushModalOpen] = useState(false);
  const [pushModalData, setPushModalData] = useState<{
    companyName: string;
    xmlType: string;
    voucherCount: number;
    masterCount: number;
    xmlContent: string;
    onSuccess?: () => void;
  } | null>(null);
  const [isPushing, setIsPushing] = useState(false);

  // Load push logs on mount
  useEffect(() => {
    getTallyPushLogs()
      .then(logs => setRecentPushLogs(logs))
      .catch(err => console.error("Error loading push logs:", err));
  }, []);

  const refreshPushLogs = async () => {
    try {
      const logs = await getTallyPushLogs();
      setRecentPushLogs(logs);
    } catch (err) {
      console.error("Error refreshing push logs:", err);
    }
  };

  const handleTestConnection = async () => {
    setTallyStatus('Connecting');
    setTallyError('');
    setTallyCompany('');
    try {
      const config = { host: tallyHost, port: tallyPort };
      const res = await testTallyConnection(config);
      if (res.success) {
        setTallyStatus('Connected');
        const summary = parseTallyResponse(res.response || '');
        if (summary.companyName) {
          setTallyCompany(summary.companyName);
        } else {
          const compRes = await fetchTallyCompany(config);
          if (compRes.success) {
            const compSummary = parseTallyResponse(compRes.response || '');
            if (compSummary.companyName) {
              setTallyCompany(compSummary.companyName);
            } else {
              setTallyCompany('Active Company Detected');
            }
          } else {
            setTallyCompany('Active Company Detected');
          }
        }
      } else {
        setTallyStatus('Error');
        setTallyError(res.error || 'Connection failed. Please verify host, port, and ensure TallyPrime is open.');
      }
    } catch (err: any) {
      setTallyStatus('Error');
      setTallyError(err.message || 'Connection failed.');
    }
  };

  const handleDisconnect = () => {
    setTallyStatus('Disconnected');
    setTallyError('');
    setTallyCompany('');
  };

  const handleFetchMasters = async () => {
    if (tallyStatus !== 'Connected') {
      setError('Please connect to Tally first.');
      return;
    }
    if (!user) return;
    setIsContextLoading(true);
    setError(null);
    try {
      const config = { host: tallyHost, port: tallyPort };
      const parsed = await fetchTallyMasters(config);
      
      const existingMappings = tallyContext?.historicalMappings || [];
      const contextPayload: any = {
        uid: user.uid,
        ledgers: Array.from(new Set(parsed.ledgers)),
        groups: Array.from(new Set(parsed.groups)),
        stockGroups: Array.from(new Set(parsed.stockGroups)),
        stockItems: Array.from(new Set(parsed.stockItems)),
        units: Array.from(new Set(parsed.units)),
        ledgerGroupMap: parsed.ledgerGroupMap,
        stockItemStockGroupMap: parsed.stockItemStockGroupMap,
        groupParentMap: parsed.groupParentMap || {},
        historicalMappings: existingMappings,
        ledgerDetails: parsed.ledgerDetails || [],
        stockItemDetails: parsed.stockItemDetails || [],
      };
      
      if (getAppMode() === 'web') {
        contextPayload.lastUpdated = serverTimestamp();
      }
      await saveTallyContext(user.uid, contextPayload);
      setIsContextLoading(false);
    } catch (err: any) {
      console.error("Failed to fetch Tally masters", err);
      setError(err.message || "Failed to fetch masters from local Tally.");
      setIsContextLoading(false);
    }
  };

  const handleFetchDaybook = async () => {
    if (tallyStatus !== 'Connected') {
      setError('Please connect to Tally first.');
      return;
    }
    if (!user) return;
    setIsContextLoading(true);
    setError(null);
    try {
      const config = { host: tallyHost, port: tallyPort };
      const txs = await fetchTallyDaybook(config, tallyFromDate, tallyToDate);
      
      const ledgerNamesFromEntries: string[] = [];
      const historicalMappings: { narration: string; ledger: string }[] = [];

      txs.forEach((tx) => {
        if (tx.ledger) {
          ledgerNamesFromEntries.push(tx.ledger);
        }
        if (tx.narration && tx.ledger) {
          const lLower = tx.ledger.toLowerCase();
          if (!lLower.includes('bank') && !lLower.includes('cash')) {
            historicalMappings.push({ narration: tx.narration, ledger: tx.ledger });
          }
        }
      });

      const currentLedgers = tallyContext?.ledgers || [];
      const mergedLedgers = Array.from(new Set([...currentLedgers, ...ledgerNamesFromEntries]));

      const currentGroups = tallyContext?.groups || [];
      const currentStockGroups = tallyContext?.stockGroups || [];
      const currentStockItems = tallyContext?.stockItems || [];
      const currentUnits = tallyContext?.units || [];
      const currentLedgerGroupMap = tallyContext?.ledgerGroupMap || {};
      const currentStockItemStockGroupMap = tallyContext?.stockItemStockGroupMap || {};
      const currentGroupParentMap = tallyContext?.groupParentMap || {};
      const currentHistorical = tallyContext?.historicalMappings || [];
      
      const mergedHistorical = [...currentHistorical, ...historicalMappings].slice(0, 500);

      const contextPayload: any = {
        uid: user.uid,
        ledgers: mergedLedgers,
        groups: currentGroups,
        stockGroups: currentStockGroups,
        stockItems: currentStockItems,
        units: currentUnits,
        ledgerGroupMap: currentLedgerGroupMap,
        stockItemStockGroupMap: currentStockItemStockGroupMap,
        groupParentMap: currentGroupParentMap,
        historicalMappings: mergedHistorical,
        ledgerDetails: tallyContext?.ledgerDetails || [],
        stockItemDetails: tallyContext?.stockItemDetails || [],
      };
      
      if (getAppMode() === 'web') {
        contextPayload.lastUpdated = serverTimestamp();
      }
      await saveTallyContext(user.uid, contextPayload);
      setIsContextLoading(false);
    } catch (err: any) {
      console.error("Failed to fetch Tally Daybook", err);
      setError(err.message || "Failed to fetch Daybook transactions from local Tally.");
      setIsContextLoading(false);
    }
  };

  const initiateTallyPush = (
    xmlContent: string,
    xmlType: string,
    voucherCount: number,
    masterCount: number,
    onSuccess?: () => void
  ) => {
    if (!isDirectTallyAvailable() || tallyStatus !== 'Connected') {
      alert('Direct Tally push is only available when Connected to local Tally Prime.');
      return;
    }
    setPushModalData({
      companyName: tallyCompany || 'Active Tally Company',
      xmlType,
      voucherCount,
      masterCount,
      xmlContent,
      onSuccess
    });
    setIsPushModalOpen(true);
  };

  const handleDirectPushToTally = async () => {
    if (!pushModalData) return;
    setIsPushing(true);
    try {
      const config = { host: tallyHost, port: tallyPort };
      const res = await pushXmlToTally(config, pushModalData.xmlContent);
      
      let status: 'success' | 'failed' = 'failed';
      let tallyResponse = '';
      let errorMessage = '';
      let created = 0;
      let altered = 0;
      let errorCount = 0;
      
      if (res.success && res.response) {
        tallyResponse = res.response;
        const parsed = parseTallyResponse(res.response);
        created = parsed.createdCount || 0;
        altered = parsed.alteredCount || 0;
        errorCount = parsed.errorCount || 0;
        
        if (errorCount === 0 && (created > 0 || altered > 0 || parsed.success)) {
          status = 'success';
        } else {
          status = 'failed';
          errorMessage = parsed.errors?.join('\n') || 'Tally returned import errors.';
        }
      } else {
        errorMessage = res.error || 'Failed to communicate with Tally.';
      }
      
      await saveTallyPushLog({
        companyName: pushModalData.companyName,
        host: tallyHost,
        port: tallyPort,
        xmlType: pushModalData.xmlType,
        voucherCount: pushModalData.voucherCount,
        masterCount: pushModalData.masterCount,
        status,
        tallyResponse,
        errorMessage
      });
      
      await refreshPushLogs();
      
      setIsPushing(false);
      setIsPushModalOpen(false);
      
      if (status === 'success') {
        alert(`Successfully imported into Tally!\nCreated: ${created}\nAltered: ${altered}`);
        if (pushModalData.onSuccess) {
          pushModalData.onSuccess();
        }
      } else {
        alert(`Failed to import into Tally.\nError Count: ${errorCount}\nDetails:\n${errorMessage}`);
      }
    } catch (err: any) {
      console.error("Direct push error:", err);
      setIsPushing(false);
      alert("Error pushing to Tally: " + (err.message || String(err)));
    }
  };

  const handlePushCombinedXmlOnTheFly = async () => {
    const validationErrors = validateReview();
    if (validationErrors.length > 0) {
      alert('Please fix validation errors first:\n' + validationErrors.join('\n'));
      return;
    }
    
    setIsProcessing(true);
    try {
      let xml = '';
      let voucherCount = 0;
      let masterCount = missingLedgers.filter(m => m.action === 'Create').length + missingStockItems.filter(m => m.action === 'Create').length;
      
      if (pendingExportType === 'Vouchers') {
        const finalVouchers = getReviewedVouchers();
        voucherCount = finalVouchers.length;
        const vouchers: TallyVoucher[] = finalVouchers.map(row => {
          const vType = row.voucherType;
          const amt = Math.abs(row.amount);
          const voucher: TallyVoucher = {
            date: row.normalizedDate,
            voucherType: vType,
            partyName: row.finalLedger,
            bankLedger: row.bankLedger,
            voucherNumber: row.voucherNumber || undefined,
            narration: row.description || undefined,
            reference: row.reference || undefined,
            ledgerEntries: []
          };
          if (vType === 'Receipt') {
            voucher.ledgerEntries.push({
              ledgerName: row.finalLedger,
              isDeemedPositive: 'No',
              isLastDeemedPositive: 'No',
              isPartyLedger: 'No',
              amount: amt
            });
            voucher.ledgerEntries.push({
              ledgerName: row.bankLedger || selectedBankLedger || 'Bank Account',
              isDeemedPositive: 'Yes',
              isLastDeemedPositive: 'Yes',
              isPartyLedger: 'Yes',
              amount: -amt
            });
          } else {
            voucher.ledgerEntries.push({
              ledgerName: row.finalLedger,
              isDeemedPositive: 'Yes',
              isLastDeemedPositive: 'Yes',
              isPartyLedger: 'No',
              amount: -amt
            });
            voucher.ledgerEntries.push({
              ledgerName: row.bankLedger || selectedBankLedger || 'Bank Account',
              isDeemedPositive: 'No',
              isLastDeemedPositive: 'No',
              isPartyLedger: 'Yes',
              amount: amt
            });
          }
          return voucher;
        });
        
        xml = generateTallyXML(vouchers, useLedgerAsNarration);
        xml = generateCombinedImportXML({
          voucherXml: xml,
          ...buildMissingMastersXMLs()
        });
      } else if (pendingExportType === 'Journal') {
        const nameMap: Record<string, string> = {};
        missingLedgers.forEach(ml => {
          const key = ml.name.toLowerCase();
          if (ml.action === 'Replace' && ml.replacementName) {
            nameMap[key] = ml.replacementName;
          } else if (ml.action === 'Create' && ml.name) {
            nameMap[key] = ml.name;
          }
        });
        const updatedGroups = journalGroups.map(g => {
          const updatedLines = g.lines.map(l => {
            const lKey = l.ledgerName.toLowerCase();
            if (nameMap[lKey]) {
              return { ...l, ledgerName: nameMap[lKey] };
            }
            return l;
          });
          return { ...g, lines: updatedLines };
        });
        voucherCount = updatedGroups.length;
        xml = generateJournalXML(updatedGroups, true);
      } else {
        const finalInvoices = getReviewedInvoices();
        voucherCount = finalInvoices.length;
        const ledgerNameMap: Record<string, string> = {};
        missingLedgers.forEach(ml => {
          if (ml.action === 'Replace' && ml.replacementName) {
            ledgerNameMap[ml.name.toLowerCase()] = ml.replacementName;
          } else if (ml.action === 'Create' && ml.name) {
            ledgerNameMap[ml.name.toLowerCase()] = ml.name;
          }
        });
        const stockItemNameMap: Record<string, string> = {};
        missingStockItems.forEach(ms => {
          if (ms.action === 'Replace' && ms.replacementName) {
            stockItemNameMap[ms.name.toLowerCase()] = ms.replacementName;
          } else if (ms.action === 'Create' && ms.name) {
            stockItemNameMap[ms.name.toLowerCase()] = ms.name;
          }
        });

        const mappedInvoices = finalInvoices.map(inv => {
          const updated = { ...inv };
          const partyKey = updated.partyName.toLowerCase();
          if (ledgerNameMap[partyKey]) {
            updated.partyName = ledgerNameMap[partyKey];
          }
          if (updated.ledgerEntries) {
            updated.ledgerEntries = updated.ledgerEntries.map(ent => {
              const entKey = ent.ledgerName.toLowerCase();
              if (ledgerNameMap[entKey]) {
                return { ...ent, ledgerName: ledgerNameMap[entKey] };
              }
              return ent;
            });
          }
          if (updated.inventoryEntries) {
            updated.inventoryEntries = updated.inventoryEntries.map(ent => {
              const entKey = ent.stockItemName.toLowerCase();
              const updatedItem = { ...ent };
              if (stockItemNameMap[entKey]) {
                updatedItem.stockItemName = stockItemNameMap[entKey];
              }
              if (updatedItem.ledgerEntries) {
                updatedItem.ledgerEntries = updatedItem.ledgerEntries.map(l => {
                  const lKey = l.ledgerName.toLowerCase();
                  if (ledgerNameMap[lKey]) {
                    return { ...l, ledgerName: ledgerNameMap[lKey] };
                  }
                  return l;
                });
              }
              return updatedItem;
            });
          }
          return updated;
        });

        xml = generateSalesPurchaseXML(mappedInvoices, companyState);
        xml = generateCombinedImportXML({
          voucherXml: xml,
          ...buildMissingMastersXMLs()
        });
      }

      setIsProcessing(false);
      initiateTallyPush(
        xml,
        'Combined Masters + Vouchers',
        voucherCount,
        masterCount,
        () => {
          setCurrentStep('complete');
        }
      );
    } catch (err: any) {
      console.error("Failed to generate combined XML for direct push", err);
      alert("Failed to generate combined XML for direct push: " + err.message);
      setIsProcessing(false);
    }
  };

  const updateVerificationRowLedger = (idx: number, ledgerName: string) => {
    // 1. Update verificationRows state
    const updatedRows = [...verificationRows];
    updatedRows[idx].finalLedger = ledgerName;
    setVerificationRows(updatedRows);

    // 2. Sync back to the source data
    const originalIdx = updatedRows[idx].sourceIdx;
    if (verificationSourceStep === 'bank-statement-review') {
      if (originalIdx !== undefined && originalIdx >= 0 && originalIdx < bankStatementRows.length) {
        const updatedBankRows = [...bankStatementRows];
        updatedBankRows[originalIdx].userLedger = ledgerName;
        setBankStatementRows(updatedBankRows);
      }
    } else {
      if (originalIdx !== undefined && originalIdx >= 0 && originalIdx < aiMappedTransactions.length) {
        const updatedMappedTx = [...aiMappedTransactions];
        updatedMappedTx[originalIdx].tallyLedger = ledgerName;
        setAiMappedTransactions(updatedMappedTx);
      }
    }
  };

  // Master States
  const [importType, setImportType] = useState<'Voucher' | 'Ledger' | 'StockItem' | 'StockGroup' | 'Unit'>('Voucher');
  const [parsedLedgers, setParsedLedgers] = useState<LedgerMasterRow[]>([]);
  const [parsedStockItems, setParsedStockItems] = useState<StockMasterRow[]>([]);
  const [parsedStockGroups, setParsedStockGroups] = useState<StockGroupRow[]>([]);
  const [parsedUnits, setParsedUnits] = useState<UnitRow[]>([]);
  const [masterReviewSearch, setMasterReviewSearch] = useState('');
  const [masterReviewFilter, setMasterReviewFilter] = useState<'all' | 'valid' | 'invalid' | 'warning' | 'duplicate' | 'excluded'>('all');
  const [showRestartConfirm, setShowRestartConfirm] = useState(false);
  const [skippedContext, setSkippedContext] = useState(false);
  const [proceededWithContext, setProceededWithContext] = useState(false);
  const isResettingWorkspaceRef = useRef(false);
  const [showAllLedgers, setShowAllLedgers] = useState(false);
  const [activeGuideField, setActiveGuideField] = useState<string | null>(null);

  // Missing Masters States
  const [missingLedgers, setMissingLedgers] = useState<MissingLedgerItem[]>([]);
  const [missingStockItems, setMissingStockItems] = useState<MissingStockItem[]>([]);
  const [showMissingMastersReview, setShowMissingMastersReview] = useState(false);
  const [pendingExportType, setPendingExportType] = useState<'Vouchers' | 'SalesPurchase' | 'Journal' | null>(null);
  const [salesPurchaseBalancingErrors, setSalesPurchaseBalancingErrors] = useState<SalesPurchaseBalancingResult[]>([]);

  // Journal Voucher States
  const [journalGroups, setJournalGroups] = useState<JournalVoucherGroup[]>([]);
  const [selectedJournalGroupIdx, setSelectedJournalGroupIdx] = useState(0);

  const handleUpdateLedger = (id: string, updates: Partial<MissingLedgerItem>) => {
    setMissingLedgers(prev => prev.map(l => l.id === id ? { ...l, ...updates } : l));
  };

  const handleUpdateStockItem = (id: string, updates: Partial<MissingStockItem>) => {
    setMissingStockItems(prev => prev.map(i => i.id === id ? { ...i, ...updates } : i));
  };

  // --- Journal Voucher Helpers ---
  const parseJournalExcel = (jsonData: any[]): JournalVoucherGroup[] => {
    const getValueByHeader = (row: any, header?: string) => {
      if (!header) return null;
      if (row[header] !== undefined) return row[header];
      const key = Object.keys(row).find(k => k.toLowerCase().trim() === header.toLowerCase().trim());
      return key ? row[key] : null;
    };

    const rawLines: JournalLine[] = jsonData.map((row: any, index: number) => {
      const rowNo = index + 2;
      const rawDate = getValueByHeader(row, 'Voucher Date');
      const rawVNo = getValueByHeader(row, 'Voucher No');
      const rawLedger = getValueByHeader(row, 'Ledger Name');
      const rawDrCr = getValueByHeader(row, 'Dr/Cr');
      const rawAmount = getValueByHeader(row, 'Amount');
      const rawNarration = getValueByHeader(row, 'Narration');
      const rawRef = getValueByHeader(row, 'Reference');
      const rawCostCentre = getValueByHeader(row, 'Cost Centre');
      const rawBillRef = getValueByHeader(row, 'Bill Reference');
      const rawRemarks = getValueByHeader(row, 'Remarks');

      const parsedAmt = parseFloat(String(rawAmount || '0').replace(/,/g, ''));
      let drCrClean: 'Dr' | 'Cr' = 'Dr';
      if (rawDrCr) {
        const dStr = String(rawDrCr).trim().toLowerCase();
        if (dStr === 'cr' || dStr === 'credit' || dStr === 'c') {
          drCrClean = 'Cr';
        }
      }

      return {
        rowNo,
        ledgerName: rawLedger ? String(rawLedger).trim() : '',
        drCr: drCrClean,
        amount: isNaN(parsedAmt) ? 0 : parsedAmt,
        narration: rawNarration ? String(rawNarration).trim() : '',
        reference: rawRef ? String(rawRef).trim() : '',
        costCentre: rawCostCentre ? String(rawCostCentre).trim() : '',
        billReference: rawBillRef ? String(rawBillRef).trim() : '',
        remarks: rawRemarks ? String(rawRemarks).trim() : '',
        excluded: false,
        voucherNo: rawVNo ? String(rawVNo).trim() : '',
        voucherDate: rawDate
      } as any;
    });

    const groupsMap: Record<string, JournalLine[]> = {};
    rawLines.forEach((line: any) => {
      const vNo = line.voucherNo || 'BLANK_VOUCHER';
      if (!groupsMap[vNo]) {
        groupsMap[vNo] = [];
      }
      groupsMap[vNo].push(line);
    });

    const groups: JournalVoucherGroup[] = Object.keys(groupsMap).map(vNo => {
      const lines = groupsMap[vNo];
      const errors: string[] = [];
      const warnings: string[] = [];

      const dates = lines.map((l: any) => l.voucherDate).filter(Boolean);
      const uniqueDates = Array.from(new Set(dates.map(d => {
        const norm = normalizeTallyDate(d);
        return norm.isValid ? norm.value : String(d);
      })));

      let normalizedDate = '';
      if (dates.length > 0) {
        const primaryDateNorm = normalizeTallyDate(dates[0]);
        if (primaryDateNorm.isValid) {
          normalizedDate = primaryDateNorm.value;
        } else {
          errors.push(`Invalid Voucher Date value "${dates[0]}". Error: ${primaryDateNorm.error}`);
        }
      } else {
        errors.push("Voucher Date is missing.");
      }

      if (uniqueDates.length > 1) {
        errors.push(`Same Voucher No "${vNo}" has multiple different dates: ${uniqueDates.join(', ')}.`);
      }

      lines.forEach(l => {
        if (!l.ledgerName) {
          errors.push(`Row ${l.rowNo}: Ledger Name cannot be blank.`);
          l.validationMessage = "Ledger Name cannot be blank.";
        }
        if (l.amount <= 0) {
          errors.push(`Row ${l.rowNo}: Amount must be a positive numeric value.`);
          l.validationMessage = "Amount must be a positive numeric value.";
        }
      });

      const debits = lines.filter(l => !l.excluded && l.drCr === 'Dr');
      const credits = lines.filter(l => !l.excluded && l.drCr === 'Cr');

      const totalDebit = debits.reduce((sum, l) => sum + l.amount, 0);
      const totalCredit = credits.reduce((sum, l) => sum + l.amount, 0);
      const difference = totalDebit - totalCredit;

      const activeLines = lines.filter(l => !l.excluded);
      if (activeLines.length < 2) {
        errors.push(`Voucher "${vNo}" requires minimum 2 ledger lines.`);
      }
      if (debits.length === 0) {
        errors.push(`Voucher "${vNo}" has no Debit (Dr) line.`);
      }
      if (credits.length === 0) {
        errors.push(`Voucher "${vNo}" has no Credit (Cr) line.`);
      }
      if (Math.abs(difference) > 0.001) {
        errors.push(`Debit total (${totalDebit.toFixed(2)}) and Credit total (${totalCredit.toFixed(2)}) do not match. Difference: ${difference.toFixed(2)}.`);
      }

      if (tallyContext) {
        lines.forEach(l => {
          if (l.ledgerName) {
            const exists = tallyContext.ledgers.some(m => m.toLowerCase() === l.ledgerName.toLowerCase());
            if (!exists) {
              warnings.push(`Row ${l.rowNo}: Ledger "${l.ledgerName}" not found in Tally masters.`);
            }
          }
        });
      }

      return {
        voucherNo: vNo === 'BLANK_VOUCHER' ? '' : vNo,
        voucherDate: normalizedDate,
        totalDebit,
        totalCredit,
        difference,
        status: Math.abs(difference) <= 0.001 && debits.length > 0 && credits.length > 0 ? 'Balanced' : 'Error',
        lines,
        errors,
        warnings,
        isValid: errors.length === 0
      };
    });

    return groups;
  };

  const generateJournalErrorExcel = (groups: JournalVoucherGroup[]) => {
    const errorGroups = groups.filter(g => !g.isValid);
    if (errorGroups.length === 0) return;

    const wb = new ExcelJS.Workbook();
    const sheet = wb.addWorksheet('Errors');

    const headers = [
      'Voucher No',
      'Voucher Date',
      'Debit Total',
      'Credit Total',
      'Difference',
      'Error Type',
      'Error Message',
      'Affected Rows'
    ];

    sheet.getRow(1).values = headers;

    errorGroups.forEach(g => {
      const rows = g.lines.map(l => l.rowNo);
      const affectedRows = rows.length > 0 ? `Rows ${Math.min(...rows)}-${Math.max(...rows)}` : '';

      let errorType = 'Validation Error';
      const errMsg = g.errors.join(' | ');
      if (errMsg.includes('do not match') || errMsg.includes('Difference')) {
        errorType = 'Debit/Credit Mismatch';
      } else if (errMsg.includes('no Debit')) {
        errorType = 'Missing Debit';
      } else if (errMsg.includes('no Credit')) {
        errorType = 'Missing Credit';
      } else if (errMsg.includes('minimum 2')) {
        errorType = 'Insufficient Lines';
      } else if (errMsg.includes('Date')) {
        errorType = 'Invalid Date';
      } else if (errMsg.includes('Ledger Name')) {
        errorType = 'Missing Ledger Name';
      }

      sheet.addRow([
        g.voucherNo || 'Blank',
        g.voucherDate || 'Blank',
        g.totalDebit,
        g.totalCredit,
        g.difference,
        errorType,
        errMsg,
        affectedRows
      ]);
    });

    sheet.columns.forEach(col => {
      let maxLen = 0;
      col.eachCell({ includeEmpty: true }, (cell) => {
        const valStr = cell.value ? String(cell.value) : '';
        if (valStr.length > maxLen) maxLen = valStr.length;
      });
      col.width = Math.max(maxLen + 3, 12);
    });

    wb.xlsx.writeBuffer().then(buffer => {
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Journal_Voucher_Errors.xlsx`;
      a.click();
      window.URL.revokeObjectURL(url);
    });
  };

  const generateSalesPurchaseBalancingErrorExcel = (errors: SalesPurchaseBalancingResult[]) => {
    if (errors.length === 0) return;

    const wb = new ExcelJS.Workbook();
    const sheet = wb.addWorksheet('Balancing_Errors');

    const headers = [
      'Invoice No',
      'Debit Total',
      'Credit Total',
      'Expected Party Amount',
      'Actual Party Amount',
      'Difference'
    ];

    sheet.getRow(1).values = headers;
    sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFF' } };
    sheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'DC2626' } // Red background for errors
    };

    errors.forEach(err => {
      sheet.addRow([
        err.invoiceNo || 'Blank',
        err.debitTotal,
        err.creditTotal,
        err.expectedPartyAmount,
        err.actualPartyAmount,
        err.difference
      ]);
    });

    sheet.columns.forEach(col => {
      let maxLen = 0;
      col.eachCell({ includeEmpty: true }, (cell) => {
        const valStr = cell.value ? String(cell.value) : '';
        if (valStr.length > maxLen) maxLen = valStr.length;
      });
      col.width = Math.max(maxLen + 3, 15);
    });

    wb.xlsx.writeBuffer().then(buffer => {
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Sales_Purchase_Balancing_Errors.xlsx`;
      a.click();
      window.URL.revokeObjectURL(url);
    });
  };

  const generateJournalXML = (groupsToUse: JournalVoucherGroup[], includeMasters: boolean = false) => {
    const vouchers: TallyVoucher[] = groupsToUse.map(g => {
      const ledgerEntries = g.lines.filter(l => !l.excluded).map(l => {
        const amt = l.drCr === 'Dr' ? -l.amount : l.amount;
        return {
          ledgerName: l.ledgerName,
          isDeemedPositive: (l.drCr === 'Dr' ? 'Yes' : 'No') as 'Yes' | 'No',
          isLastDeemedPositive: (l.drCr === 'Dr' ? 'Yes' : 'No') as 'Yes' | 'No',
          isPartyLedger: 'No' as 'Yes' | 'No',
          amount: amt
        };
      });

      return {
        date: g.voucherDate,
        voucherType: 'Journal',
        partyName: '',
        bankLedger: '',
        voucherNumber: g.voucherNo,
        narration: g.lines[0]?.narration || undefined,
        reference: g.lines[0]?.reference || undefined,
        ledgerEntries
      };
    });

    let xml = generateTallyXML(vouchers);
    if (includeMasters) {
      xml = generateCombinedImportXML({
        voucherXml: xml,
        ...buildMissingMastersXMLs()
      });
    }
    return xml;
  };

  const checkMissingMastersAndProceedJournal = (groups: JournalVoucherGroup[]) => {
    setError(null);
    const validGroups = groups.filter(g => g.isValid);
    if (validGroups.length === 0) {
      setError("No valid, balanced Journal vouchers available to export. Please correct errors.");
      setIsProcessing(false);
      return;
    }

    const ledgersList = tallyContext?.ledgers || [];
    const usedLedgers: { name: string; source: string; type: string }[] = [];

    validGroups.forEach(g => {
      const sourceId = g.voucherNo ? `Journal No: ${g.voucherNo}` : `Journal Voucher`;
      g.lines.forEach(l => {
        if (!l.excluded && l.ledgerName) {
          usedLedgers.push({ name: l.ledgerName, source: sourceId, type: 'Journal Ledger' });
        }
      });
    });

    const detectedLedgers = detectMissingLedgers(usedLedgers, ledgersList);

    if (detectedLedgers.length > 0) {
      setMissingLedgers(detectedLedgers);
      setMissingStockItems([]);
      setPendingExportType('Journal');
      setShowMissingMastersReview(true);
      setIsProcessing(false);
    } else {
      generateFinalJournalXMLDirect(validGroups);
    }
  };

  const generateFinalJournalXMLDirect = async (groupsToUse: JournalVoucherGroup[], includeMasters: boolean = false) => {
    if (!user || groupsToUse.length === 0) return;
    setIsProcessing(true);
    setError(null);

    try {
      const conversionPayload: any = {
        fileName: pendingFileName,
        status: 'processing',
        voucherType: 'Journal'
      };
      if (getAppMode() === 'web') {
        conversionPayload.timestamp = serverTimestamp();
      }
      const conversionId = await saveConversion(user.uid, conversionPayload);

      const nameMap: Record<string, string> = {};
      missingLedgers.forEach(ml => {
        const key = ml.name.toLowerCase();
        if (ml.action === 'Replace' && ml.replacementName) {
          nameMap[key] = ml.replacementName;
        } else if (ml.action === 'Create' && ml.name) {
          nameMap[key] = ml.name;
        }
      });

      const updatedGroups = groupsToUse.map(g => {
        const updatedLines = g.lines.map(l => {
          const lKey = l.ledgerName.toLowerCase();
          if (nameMap[lKey]) {
            return { ...l, ledgerName: nameMap[lKey] };
          }
          return l;
        });
        return { ...g, lines: updatedLines };
      });

      const xml = generateJournalXML(updatedGroups, includeMasters);

      await updateConversion(user.uid, conversionId, {
        status: 'completed',
        xmlContent: xml
      });

      const newRecord: ConversionRecord = {
        id: conversionId,
        fileName: pendingFileName,
        timestamp: { seconds: Math.floor(Date.now() / 1000) },
        status: 'completed',
        xmlContent: xml,
        voucherType: 'Journal'
      };

      setConversions(prev => [newRecord, ...prev]);
      setCurrentStep('complete');
      setIsProcessing(false);
      setShowMissingMastersReview(false);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Conversion failed");
      setIsProcessing(false);
    }
  };

  const updateJournalLine = (groupIndex: number, lineIndex: number, updates: Partial<JournalLine>) => {
    setJournalGroups(prev => {
      const copy = [...prev];
      const grp = { ...copy[groupIndex] };
      const lines = [...grp.lines];
      lines[lineIndex] = { ...lines[lineIndex], ...updates };

      const errors: string[] = [];
      const warnings: string[] = [];

      const dates = lines.map((l: any) => l.voucherDate).filter(Boolean);
      const uniqueDates = Array.from(new Set(dates.map(d => {
        const norm = normalizeTallyDate(d);
        return norm.isValid ? norm.value : String(d);
      })));

      let normalizedDate = grp.voucherDate;
      if (dates.length > 0) {
        const primaryDateNorm = normalizeTallyDate(dates[0]);
        if (primaryDateNorm.isValid) {
          normalizedDate = primaryDateNorm.value;
        } else {
          errors.push(`Invalid Voucher Date value "${dates[0]}". Error: ${primaryDateNorm.error}`);
        }
      }

      if (uniqueDates.length > 1) {
        errors.push(`Same Voucher No "${grp.voucherNo}" has multiple different dates.`);
      }

      lines.forEach((l, idx) => {
        l.validationMessage = undefined;
        if (!l.excluded) {
          if (!l.ledgerName) {
            errors.push(`Row ${l.rowNo || idx + 1}: Ledger Name cannot be blank.`);
            l.validationMessage = "Ledger Name cannot be blank.";
          }
          if (l.amount <= 0) {
            errors.push(`Row ${l.rowNo || idx + 1}: Amount must be a positive numeric value.`);
            l.validationMessage = "Amount must be a positive numeric value.";
          }
        }
      });

      const debits = lines.filter(l => !l.excluded && l.drCr === 'Dr');
      const credits = lines.filter(l => !l.excluded && l.drCr === 'Cr');

      const totalDebit = debits.reduce((sum, l) => sum + l.amount, 0);
      const totalCredit = credits.reduce((sum, l) => sum + l.amount, 0);
      const difference = totalDebit - totalCredit;

      const activeLines = lines.filter(l => !l.excluded);
      if (activeLines.length < 2) {
        errors.push("Requires minimum 2 ledger lines.");
      }
      if (debits.length === 0) {
        errors.push("Has no Debit (Dr) line.");
      }
      if (credits.length === 0) {
        errors.push("Has no Credit (Cr) line.");
      }
      if (Math.abs(difference) > 0.001) {
        errors.push(`Debit total (${totalDebit.toFixed(2)}) and Credit total (${totalCredit.toFixed(2)}) do not match. Diff: ${difference.toFixed(2)}.`);
      }

      if (tallyContext) {
        lines.forEach((l, idx) => {
          if (!l.excluded && l.ledgerName) {
            const exists = tallyContext.ledgers.some(m => m.toLowerCase() === l.ledgerName.toLowerCase());
            if (!exists) {
              warnings.push(`Row ${l.rowNo || idx + 1}: Ledger "${l.ledgerName}" not found in Tally masters.`);
            }
          }
        });
      }

      grp.lines = lines;
      grp.totalDebit = totalDebit;
      grp.totalCredit = totalCredit;
      grp.difference = difference;
      grp.errors = errors;
      grp.warnings = warnings;
      grp.status = Math.abs(difference) <= 0.001 && debits.length > 0 && credits.length > 0 ? 'Balanced' : 'Error';
      grp.isValid = errors.length === 0;

      copy[groupIndex] = grp;
      return copy;
    });
  };

  const updateJournalGroupHeader = (groupIndex: number, fields: { voucherNo?: string; voucherDate?: string }) => {
    setJournalGroups(prev => {
      const copy = [...prev];
      const grp = { ...copy[groupIndex] };
      if (fields.voucherNo !== undefined) grp.voucherNo = fields.voucherNo;
      if (fields.voucherDate !== undefined) {
        const norm = normalizeTallyDate(fields.voucherDate);
        grp.voucherDate = norm.isValid ? norm.value : fields.voucherDate;
      }
      copy[groupIndex] = grp;
      return copy;
    });
  };

  const addJournalLine = (groupIndex: number) => {
    setJournalGroups(prev => {
      const copy = [...prev];
      const grp = { ...copy[groupIndex] };
      const lines = [...grp.lines];
      lines.push({
        rowNo: lines.length > 0 ? Math.max(...lines.map(l => l.rowNo)) + 1 : 1,
        ledgerName: '',
        drCr: 'Dr',
        amount: 0,
        narration: '',
        reference: '',
        costCentre: '',
        billReference: '',
        remarks: '',
        excluded: false
      });
      grp.lines = lines;
      copy[groupIndex] = grp;
      return copy;
    });
    setTimeout(() => {
      updateJournalLine(groupIndex, journalGroups[groupIndex]?.lines.length || 0, {});
    }, 10);
  };

  // Bank Statement Import States
  const [voucherImportMethod, setVoucherImportMethod] = useState<'template' | 'bankStatement'>('template');
  const [voucherMode, setVoucherMode] = useState<'auto' | 'payment' | 'receipt'>('auto');
  const [bankStatementRows, setBankStatementRows] = useState<any[]>([]);

  // Advanced Bank Statement States
  const [rawGrid, setRawGrid] = useState<any[][]>([]);
  const [columnMappings, setColumnMappings] = useState<{
    date: number | null;
    narration: number | null;
    debit: number | null;
    credit: number | null;
    amount: number | null;
    drCr: number | null;
    balance: number | null;
    reference: number | null;
  }>({
    date: null,
    narration: null,
    debit: null,
    credit: null,
    amount: null,
    drCr: null,
    balance: null,
    reference: null,
  });
  const [headerRowIdx, setHeaderRowIdx] = useState<number>(0);
  const [dataStartRowIdx, setDataStartRowIdx] = useState<number>(1);
  const [dataEndRowIdx, setDataEndRowIdx] = useState<number>(0);
  const [detectedVoucherMode, setDetectedVoucherMode] = useState<'auto' | 'payment' | 'receipt'>('auto');

  // Field Aliases for deterministic mapping
  const LEDGER_FIELDS = {
    ledgerName: ['ledger name', 'name', 'ledgername'],
    underGroup: ['under group', 'group', 'parent', 'undergroup'],
    openingBalance: ['opening balance', 'openingbalance', 'balance'],
    drCr: ['dr/cr', 'drcr', 'dr cr', 'type'],
    mailingName: ['mailing name', 'mailingname'],
    address1: ['address line 1', 'addressline1', 'address 1', 'address1'],
    address2: ['address line 2', 'addressline2', 'address 2', 'address2'],
    state: ['state'],
    country: ['country'],
    pincode: ['pincode', 'pin code', 'pin'],
    pan: ['pan', 'pan number', 'pan no'],
    gstin: ['gstin', 'gst number', 'gstin/u登録', 'gst in', 'gstno', 'gst'],
    registrationType: ['registration type', 'registrationtype', 'gst registration type'],
    taxability: ['taxability'],
    isBillwiseOn: ['is billwise on', 'billwise', 'bill wise'],
    isCostCentreOn: ['is cost centre on', 'cost centre', 'cost center'],
    email: ['email', 'e-mail'],
    mobileNumber: ['mobile number', 'mobile', 'phone']
  };

  const STOCK_FIELDS = {
    itemName: ['stock item name', 'item name', 'name', 'item', 'stockitemname'],
    underGroup: ['under stock group', 'stock group', 'group', 'parent', 'understockgroup'],
    unit: ['unit', 'uom', 'base unit', 'units'],
    openingQty: ['opening quantity', 'opening qty', 'quantity', 'qty'],
    openingRate: ['opening rate', 'rate'],
    openingValue: ['opening value', 'value', 'opening amount', 'amount'],
    hsn: ['hsn/sac', 'hsn', 'sac', 'hsn code'],
    gstApplicable: ['gst applicable', 'gstapplicable', 'gst'],
    taxability: ['taxability'],
    gstRate: ['gst rate', 'gst %', 'rate %'],
    cgstRate: ['cgst rate', 'cgst %', 'cgst'],
    sgstRate: ['sgst rate', 'sgst %', 'sgst'],
    igstRate: ['igst rate', 'igst %', 'igst'],
    description: ['description', 'item description']
  };

  const editDistance = (s1: string, s2: string): number => {
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
  };

  const getSimilarity = (s1: string, s2: string): number => {
    const norm1 = s1.toLowerCase().replace(/[^a-z0-9]/g, '');
    const norm2 = s2.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (norm1 === norm2) return 1.0;
    
    const longer = norm1.length > norm2.length ? norm1 : norm2;
    const shorter = norm1.length > norm2.length ? norm2 : norm1;
    
    if (longer.length === 0) return 1.0;
    
    const distance = editDistance(longer, shorter);
    return (longer.length - distance) / longer.length;
  };

  const validateGSTIN = (gstin: string): { isValid: boolean; stateName?: string; error?: string } => {
    const cleanGst = String(gstin).trim().toUpperCase();
    if (cleanGst.length !== 15) {
      return { isValid: false, error: 'GSTIN must be exactly 15 characters long.' };
    }

    const stateCode = cleanGst.substring(0, 2);
    const stateMap: Record<string, string> = {
      '01': 'Jammu & Kashmir', '02': 'Himachal Pradesh', '03': 'Punjab', '04': 'Chandigarh', '05': 'Uttarakhand',
      '06': 'Haryana', '07': 'Delhi', '08': 'Rajasthan', '09': 'Uttar Pradesh', '10': 'Bihar',
      '11': 'Sikkim', '12': 'Arunachal Pradesh', '13': 'Nagaland', '14': 'Manipur', '15': 'Mizoram',
      '16': 'Tripura', '17': 'Meghalaya', '18': 'Assam', '19': 'West Bengal', '20': 'Jharkhand',
      '21': 'Odisha', '22': 'Chhattisgarh', '23': 'Madhya Pradesh', '24': 'Gujarat', '25': 'Daman & Diu',
      '26': 'Dadra & Nagar Haveli', '27': 'Maharashtra', '28': 'Andhra Pradesh', '29': 'Karnataka', '30': 'Goa',
      '31': 'Lakshadweep', '32': 'Kerala', '33': 'Tamil Nadu', '34': 'Puducherry', '35': 'Andaman & Nicobar Islands',
      '36': 'Telangana', '37': 'Andhra Pradesh', '38': 'Ladakh', '97': 'Other Territory'
    };

    const stateName = stateMap[stateCode];
    if (!stateName) {
      return { isValid: false, error: `Invalid State Code: "${stateCode}". Must be a valid Indian state code.` };
    }

    const pan = cleanGst.substring(2, 12);
    const panRegex = /^[A-Z]{5}[0-9]{4}[A-Z]{1}$/;
    if (!panRegex.test(pan)) {
      return { isValid: false, error: `Characters 3 to 12 must match PAN structure (e.g., ABCDE1234F). Got "${pan}".` };
    }

    const entityCode = cleanGst.charAt(12);
    if (!/^[0-9A-Z]$/.test(entityCode)) {
      return { isValid: false, error: `13th character (Entity code) must be alphanumeric. Got "${entityCode}".` };
    }

    const zChar = cleanGst.charAt(13);
    if (!/^[0-9A-Z]$/.test(zChar)) {
      return { isValid: false, error: `14th character must be alphanumeric. Got "${zChar}".` };
    }

    const checksum = cleanGst.charAt(14);
    if (!/^[0-9A-Z]$/.test(checksum)) {
      return { isValid: false, error: `15th character (checksum) must be alphanumeric. Got "${checksum}".` };
    }

    return { isValid: true, stateName };
  };

  const mapRowToFields = (row: any, fieldMappings: Record<string, string[]>) => {
    const result: Record<string, any> = {};
    const rowKeys = Object.keys(row);
    
    Object.entries(fieldMappings).forEach(([field, aliases]) => {
      const matchedKey = rowKeys.find(key => {
        const cleanKey = key.toLowerCase().trim();
        return aliases.includes(cleanKey) || aliases.some(alias => cleanKey.includes(alias));
      });
      result[field] = matchedKey ? row[matchedKey] : undefined;
    });
    
    return result;
  };

  const DEFAULT_LEDGERS = [
    'Suspense Account',
    'Cash',
    'Profit & Loss A/c',
    'Bank Account',
    'Sales A/c',
    'Purchase A/c',
    'Indirect Expenses',
    'Indirect Incomes',
    'Salary Expense',
    'Rent Expense',
    'GST Ledger',
  ];

  const DEFAULT_GROUPS = [
    'Primary',
    'Sundry Debtors',
    'Sundry Creditors',
    'Bank Accounts',
    'Cash-in-hand',
    'Indirect Expenses',
    'Indirect Incomes',
    'Direct Expenses',
    'Direct Incomes',
    'Sales Accounts',
    'Purchase Accounts',
    'Capital Account',
    'Loans (Liability)',
    'Secured Loans',
    'Unsecured Loans',
    'Current Liabilities',
    'Duties & Taxes',
    'Provisions',
    'Current Assets',
    'Stock-in-hand',
    'Loans & Advances (Asset)',
    'Investments',
    'Fixed Assets',
    'Suspense A/c'
  ];

  const DEFAULT_STOCK_GROUPS = [
    'Primary'
  ];

  const DEFAULT_UNITS = [
    'NOS',
    'PCS',
    'BOX',
    'KGS',
    'MTR',
    'SET',
    'BAG',
    'DOZ',
    'HRS',
    'MIN',
  ];

  const buildMastersSheet = (data: Record<string, string[]>) => {
    const keys = Object.keys(data);
    const maxLen = Math.max(...keys.map(k => data[k].length));
    const rows = [];
    for (let i = 0; i < maxLen; i++) {
      const row: any = {};
      keys.forEach(k => {
        row[k] = data[k][i] !== undefined ? data[k][i] : '';
      });
      rows.push(row);
    }
    return XLSX.utils.json_to_sheet(rows, { header: keys });
  };

  const colNumToLetter = (col: number): string => {
    let temp = '';
    let letter = '';
    while (col > 0) {
      temp = String.fromCharCode(((col - 1) % 26) + 65);
      letter = temp + letter;
      col = Math.floor((col - 1) / 26);
    }
    return letter;
  };

  const applyDropdownByHeader = (
    worksheet: ExcelJS.Worksheet,
    headers: string[],
    headerName: string,
    validation: ExcelJS.DataValidation
  ) => {
    const colIndex = headers.findIndex(h => h.toLowerCase().trim() === headerName.toLowerCase().trim()) + 1;
    if (colIndex <= 0) return;
    const colLetter = colNumToLetter(colIndex);
    for (let r = 2; r <= 1000; r++) {
      worksheet.getCell(`${colLetter}${r}`).dataValidation = validation;
    }
  };

  const downloadTemplate = async (type: string) => {
    const isSalesMode = type.toLowerCase().includes('sales');
    const isPurchaseMode = type.toLowerCase().includes('purchase');
    const isSalesPurchase = isSalesMode || isPurchaseMode;

    if (isSalesPurchase) {
      const realVoucherType = isSalesMode ? 'Sales' : 'Purchase';
      const isVoucherwise = type.toLowerCase().includes('voucherwise');
      const isSales = isSalesMode;
      const headers = isVoucherwise ? [
        'Invoice Date', 'Invoice No', 'Voucher Type', 'Party Ledger', 'Party GSTIN', 'Place of Supply',
        'Sales/Purchase Ledger', 'Taxable Value (Purchase Value)', 'GST Mode', 'GST Rate %',
        'CGST Ledger', 'CGST Amount', 'SGST Ledger', 'SGST Amount', 'IGST Ledger', 'IGST Amount',
        'Other Ledger 1', 'Other Amount 1', 'Round Off Ledger', 'Round Off Amount',
        'Stock Item', 'Description', 'Quantity', 'Unit', 'Rate', 'Item Amount', 'Discount %', 'HSN/SAC',
        'Party Address 1', 'Party Address 2', 'Party State',
        'Dispatch Date', 'Delivery Note No', 'Dispatch Doc No', 'Bilty LR No', 'Transporter Name', 'Transporter GSTIN', 'Vehicle No', 'Destination', 'Mode of Transport', 'Eway Bill No',
        'Freight Ledger', 'Freight Amount', 'Packing Ledger', 'Packing Amount', 'Loading Ledger', 'Loading Amount', 'Insurance Ledger', 'Insurance Amount',
        'Other Ledger 2', 'Other Amount 2', 'Discount Ledger', 'Bill Discount Amount',
        'Narration', 'Reference', 'Voucher Mode', 'Inventory Mode'
      ] : [
        'Invoice Date', 'Invoice No', 'Voucher Type', 'Party Ledger', 'Sales/Purchase Ledger',
        'Stock Item', 'Description', 'Quantity', 'Unit', 'Rate', 'Item Amount', 'Discount %', 'Taxable Value', 'HSN/SAC',
        'Party GSTIN', 'Party Address 1', 'Party Address 2', 'Party State', 'Place of Supply',
        'Dispatch Date', 'Delivery Note No', 'Dispatch Doc No', 'Bilty LR No', 'Transporter Name', 'Transporter GSTIN', 'Vehicle No', 'Destination', 'Mode of Transport', 'Eway Bill No',
        'GST Mode', 'GST Rate %', 'CGST Ledger', 'CGST Amount', 'SGST Ledger', 'SGST Amount', 'IGST Ledger', 'IGST Amount',
        'Freight Ledger', 'Freight Amount', 'Packing Ledger', 'Packing Amount', 'Loading Ledger', 'Loading Amount', 'Insurance Ledger', 'Insurance Amount', 'Other Ledger 1', 'Other Amount 1', 'Other Ledger 2', 'Other Amount 2', 'Discount Ledger', 'Bill Discount Amount', 'Round Off Ledger', 'Round Off Amount',
        'Narration', 'Reference', 'Voucher Mode', 'Inventory Mode'
      ];

      const sampleData = isVoucherwise ? [
        {
          'Invoice Date': new Date().toLocaleDateString('en-GB'),
          'Invoice No': 'INV-2001',
          'Voucher Type': realVoucherType,
          'Party Ledger': isSales ? 'Cash-in-hand' : 'Sundry Creditors',
          'Party GSTIN': '27ABCDE1234F1Z5',
          'Place of Supply': 'Maharashtra',
          'Sales/Purchase Ledger': isSales ? 'Sales Account' : 'Purchase Account',
          'Taxable Value (Purchase Value)': 10000,
          'GST Mode': 'Auto',
          'GST Rate %': 18,
          'CGST Ledger': isSales ? 'CGST Output 9%' : 'CGST Input 9%',
          'CGST Amount': '',
          'SGST Ledger': isSales ? 'SGST Output 9%' : 'SGST Input 9%',
          'SGST Amount': '',
          'IGST Ledger': '',
          'IGST Amount': '',
          'Other Ledger 1': '',
          'Other Amount 1': '',
          'Round Off Ledger': 'Round Off A/c',
          'Round Off Amount': '0.00',
          'Stock Item': '',
          'Description': '',
          'Quantity': '',
          'Unit': '',
          'Rate': '',
          'Item Amount': '',
          'Discount %': '',
          'HSN/SAC': '',
          'Party Address 1': '456 Commercial Street',
          'Party Address 2': 'Business Park',
          'Party State': 'Maharashtra',
          'Dispatch Date': '',
          'Delivery Note No': '',
          'Dispatch Doc No': '',
          'Bilty LR No': '',
          'Transporter Name': '',
          'Transporter GSTIN': '',
          'Vehicle No': '',
          'Destination': '',
          'Mode of Transport': '',
          'Eway Bill No': '',
          'Freight Ledger': '',
          'Freight Amount': '',
          'Packing Ledger': '',
          'Packing Amount': '',
          'Loading Ledger': '',
          'Loading Amount': '',
          'Insurance Ledger': '',
          'Insurance Amount': '',
          'Other Ledger 2': '',
          'Other Amount 2': '',
          'Discount Ledger': '',
          'Bill Discount Amount': '',
          'Narration': `${realVoucherType} voucher entry generated via TallyGen Pro (Voucherwise)`,
          'Reference': 'REF-2001',
          'Voucher Mode': 'Accounting',
          'Inventory Mode': 'Inventory Optional'
        }
      ] : [
        {
          'Invoice Date': new Date().toLocaleDateString('en-GB'),
          'Invoice No': 'INV-1001',
          'Voucher Type': realVoucherType,
          'Party Ledger': isSales ? 'Cash-in-hand' : 'Sundry Creditors',
          'Sales/Purchase Ledger': isSales ? 'Sales Account' : 'Purchase Account',
          'Stock Item': 'Sample Stock Item',
          'Description': 'Item 1 description',
          'Quantity': '10',
          'Unit': 'NOS',
          'Rate': '500',
          'Item Amount': '5000',
          'Discount %': '5',
          'Taxable Value': '4750',
          'HSN/SAC': '8471',
          'Party GSTIN': '27ABCDE1234F1Z5',
          'Party Address 1': '123 Tech Park',
          'Party Address 2': 'Industrial Area',
          'Party State': 'Maharashtra',
          'Place of Supply': 'Maharashtra',
          'Dispatch Date': '',
          'Delivery Note No': '',
          'Dispatch Doc No': '',
          'Bilty LR No': '',
          'Transporter Name': 'Express Cargo',
          'Transporter GSTIN': '27ABCDE1234F1Z5',
          'Vehicle No': 'MH-12-PQ-9999',
          'Destination': 'Warehouse A',
          'Mode of Transport': 'Road',
          'Eway Bill No': '123456789012',
          'GST Mode': 'Auto',
          'GST Rate %': '18',
          'CGST Ledger': isSales ? 'CGST Output 9%' : 'CGST Input 9%',
          'CGST Amount': '',
          'SGST Ledger': isSales ? 'SGST Output 9%' : 'SGST Input 9%',
          'SGST Amount': '',
          'IGST Ledger': '',
          'IGST Amount': '',
          'Freight Ledger': 'Freight Charges',
          'Freight Amount': '100',
          'Packing Ledger': 'Packing Expenses',
          'Packing Amount': '50',
          'Loading Ledger': '',
          'Loading Amount': '',
          'Insurance Ledger': '',
          'Insurance Amount': '',
          'Other Ledger 1': '',
          'Other Amount 1': '',
          'Other Ledger 2': '',
          'Other Amount 2': '',
          'Discount Ledger': '',
          'Bill Discount Amount': '',
          'Round Off Ledger': 'Round Off A/c',
          'Round Off Amount': '0.50',
          'Narration': `${realVoucherType} invoice generated via TallyGen Pro (Itemwise)`,
          'Reference': 'REF-1001',
          'Voucher Mode': 'Item Invoice',
          'Inventory Mode': 'Inventory Optional'
        },
        {
          'Invoice Date': new Date().toLocaleDateString('en-GB'),
          'Invoice No': 'INV-1001',
          'Voucher Type': realVoucherType,
          'Party Ledger': isSales ? 'Cash-in-hand' : 'Sundry Creditors',
          'Sales/Purchase Ledger': isSales ? 'Sales Account' : 'Purchase Account',
          'Stock Item': 'Another Stock Item',
          'Description': 'Item 2 description',
          'Quantity': '5',
          'Unit': 'NOS',
          'Rate': '300',
          'Item Amount': '1500',
          'Discount %': '10',
          'Taxable Value': '1350',
          'HSN/SAC': '8518',
          'Party GSTIN': '27ABCDE1234F1Z5',
          'Party Address 1': '123 Tech Park',
          'Party Address 2': 'Industrial Area',
          'Party State': 'Maharashtra',
          'Place of Supply': 'Maharashtra',
          'Dispatch Date': '',
          'Delivery Note No': '',
          'Dispatch Doc No': '',
          'Bilty LR No': '',
          'Transporter Name': 'Express Cargo',
          'Transporter GSTIN': '27ABCDE1234F1Z5',
          'Vehicle No': 'MH-12-PQ-9999',
          'Destination': 'Warehouse A',
          'Mode of Transport': 'Road',
          'Eway Bill No': '123456789012',
          'GST Mode': 'Auto',
          'GST Rate %': '18',
          'CGST Ledger': isSales ? 'CGST Output 9%' : 'CGST Input 9%',
          'CGST Amount': '',
          'SGST Ledger': isSales ? 'SGST Output 9%' : 'SGST Input 9%',
          'SGST Amount': '',
          'IGST Ledger': '',
          'IGST Amount': '',
          'Freight Ledger': 'Freight Charges',
          'Freight Amount': '100',
          'Packing Ledger': 'Packing Expenses',
          'Packing Amount': '50',
          'Loading Ledger': '',
          'Loading Amount': '',
          'Insurance Ledger': '',
          'Insurance Amount': '',
          'Other Ledger 1': '',
          'Other Amount 1': '',
          'Other Ledger 2': '',
          'Other Amount 2': '',
          'Discount Ledger': '',
          'Bill Discount Amount': '',
          'Round Off Ledger': 'Round Off A/c',
          'Round Off Amount': '0.50',
          'Narration': `${realVoucherType} invoice generated via TallyGen Pro (Itemwise)`,
          'Reference': 'REF-1001',
          'Voucher Mode': 'Item Invoice',
          'Inventory Mode': 'Inventory Optional'
        }
      ];

      const wb = new ExcelJS.Workbook();
      const templateSheet = wb.addWorksheet('Template');
      const mastersSheet = wb.addWorksheet('Tally_Masters');

      mastersSheet.views = [{ showGridLines: false }];
      mastersSheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFF' } };
      mastersSheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '414141' }
      };

      const DEFAULT_LEDGERS = ['Sundry Debtors', 'Sundry Creditors', 'Sales Account', 'Purchase Account', 'Cash-in-hand', 'CGST Output 9%', 'SGST Output 9%', 'CGST Input 9%', 'SGST Input 9%', 'Freight Charges', 'Packing Expenses', 'Round Off A/c'];
      const DEFAULT_UNITS = ['NOS', 'PCS', 'KG', 'BOX', 'SET', 'MTR'];

      const ledgersList = (tallyContext && tallyContext.ledgers && tallyContext.ledgers.length > 0)
        ? tallyContext.ledgers
        : DEFAULT_LEDGERS;

      const salesPurchaseLedgers = ledgersList.filter(l => 
        l.toLowerCase().includes('sale') || 
        l.toLowerCase().includes('purchase') || 
        l.toLowerCase().includes('account') ||
        l.toLowerCase().includes('a/c')
      );
      const finalSalesPurchaseLedgers = salesPurchaseLedgers.length > 0 ? salesPurchaseLedgers : ledgersList;

      const stockList = (tallyContext && tallyContext.stockItems && tallyContext.stockItems.length > 0)
        ? tallyContext.stockItems
        : ['Sample Stock Item', 'Another Stock Item'];

      const unitsList = (tallyContext && tallyContext.units && tallyContext.units.length > 0)
        ? tallyContext.units
        : DEFAULT_UNITS;

      const gstLedgers = ledgersList.filter(l => 
        l.toLowerCase().includes('cgst') || 
        l.toLowerCase().includes('sgst') || 
        l.toLowerCase().includes('igst') || 
        l.toLowerCase().includes('utgst') || 
        l.toLowerCase().includes('tax') || 
        l.toLowerCase().includes('duty') || 
        l.toLowerCase().includes('gst')
      );
      const finalGstLedgers = gstLedgers.length > 0 ? gstLedgers : ledgersList;

      const finalAdditionalLedgers = ledgersList;

      const roundOffLedgers = ledgersList.filter(l => l.toLowerCase().includes('round'));
      const finalRoundOffLedgers = roundOffLedgers.length > 0 ? roundOffLedgers : ledgersList;

      const finalStates = [
        'Andhra Pradesh', 'Arunachal Pradesh', 'Assam', 'Bihar', 'Chhattisgarh',
        'Goa', 'Gujarat', 'Haryana', 'Himachal Pradesh', 'Jharkhand', 'Karnataka',
        'Kerala', 'Madhya Pradesh', 'Maharashtra', 'Manipur', 'Meghalaya', 'Mizoram',
        'Nagaland', 'Odisha', 'Punjab', 'Rajasthan', 'Sikkim', 'Tamil Nadu',
        'Telangana', 'Tripura', 'Uttar Pradesh', 'Uttarakhand', 'West Bengal',
        'Andaman and Nicobar Islands', 'Chandigarh', 'Dadra and Nagar Haveli and Daman and Diu',
        'Delhi', 'Jammu and Kashmir', 'Ladakh', 'Lakshadweep', 'Puducherry'
      ];

      const finalTransportModes = ['Road', 'Rail', 'Air', 'Ship'];
      const voucherTypes = ['Payment', 'Receipt', 'Contra', 'Journal', 'Sales', 'Purchase'];

      // Set headers on Tally_Masters
      mastersSheet.getRow(1).values = [
        'Voucher_Type',
        'Party_Ledger',
        'Sales_Purchase_Ledger',
        'Stock_Item',
        'Unit',
        'GST_Ledgers',
        'Additional_Ledgers',
        'Round_Off_Ledgers',
        'Place_of_Supply',
        'Mode_of_Transport'
      ];

      const maxRows = Math.max(
        voucherTypes.length,
        ledgersList.length,
        finalSalesPurchaseLedgers.length,
        stockList.length,
        unitsList.length,
        finalGstLedgers.length,
        finalAdditionalLedgers.length,
        finalRoundOffLedgers.length,
        finalStates.length,
        finalTransportModes.length
      );

      for (let i = 0; i < maxRows; i++) {
        mastersSheet.getRow(i + 2).values = [
          voucherTypes[i] || '',
          ledgersList[i] || '',
          finalSalesPurchaseLedgers[i] || '',
          stockList[i] || '',
          unitsList[i] || '',
          finalGstLedgers[i] || '',
          finalAdditionalLedgers[i] || '',
          finalRoundOffLedgers[i] || '',
          finalStates[i] || '',
          finalTransportModes[i] || ''
        ];
      }

      // Populate Template sheet
      templateSheet.getRow(1).values = headers;
      templateSheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFF' } };
      templateSheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '1F2937' }
      };

      sampleData.forEach(row => {
        const rowValues = headers.map(h => row[h as keyof typeof row] !== undefined ? row[h as keyof typeof row] : '');
        templateSheet.addRow(rowValues);
      });

      // Apply dropdown validations using applyDropdownByHeader
      applyDropdownByHeader(templateSheet, headers, 'Voucher Type', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$A$2:$A$${voucherTypes.length + 1}`]
      });
      applyDropdownByHeader(templateSheet, headers, 'Party Ledger', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$B$2:$B$${ledgersList.length + 1}`]
      });
      applyDropdownByHeader(templateSheet, headers, 'Sales/Purchase Ledger', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$C$2:$C$${finalSalesPurchaseLedgers.length + 1}`]
      });
      applyDropdownByHeader(templateSheet, headers, 'Stock Item', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$D$2:$D$${stockList.length + 1}`]
      });
      applyDropdownByHeader(templateSheet, headers, 'Unit', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$E$2:$E$${unitsList.length + 1}`]
      });
      applyDropdownByHeader(templateSheet, headers, 'Party State', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$I$2:$I$${finalStates.length + 1}`]
      });
      applyDropdownByHeader(templateSheet, headers, 'Place of Supply', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$I$2:$I$${finalStates.length + 1}`]
      });
      applyDropdownByHeader(templateSheet, headers, 'Mode of Transport', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$J$2:$J$${finalTransportModes.length + 1}`]
      });
      applyDropdownByHeader(templateSheet, headers, 'GST Mode', {
        type: 'list',
        allowBlank: true,
        formulae: ['"Auto,Manual"']
      });
      applyDropdownByHeader(templateSheet, headers, 'CGST Ledger', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$F$2:$F$${finalGstLedgers.length + 1}`]
      });
      applyDropdownByHeader(templateSheet, headers, 'SGST Ledger', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$F$2:$F$${finalGstLedgers.length + 1}`]
      });
      applyDropdownByHeader(templateSheet, headers, 'IGST Ledger', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$F$2:$F$${finalGstLedgers.length + 1}`]
      });
      applyDropdownByHeader(templateSheet, headers, 'Freight Ledger', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$G$2:$G$${finalAdditionalLedgers.length + 1}`]
      });
      applyDropdownByHeader(templateSheet, headers, 'Packing Ledger', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$G$2:$G$${finalAdditionalLedgers.length + 1}`]
      });
      applyDropdownByHeader(templateSheet, headers, 'Loading Ledger', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$G$2:$G$${finalAdditionalLedgers.length + 1}`]
      });
      applyDropdownByHeader(templateSheet, headers, 'Insurance Ledger', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$G$2:$G$${finalAdditionalLedgers.length + 1}`]
      });
      applyDropdownByHeader(templateSheet, headers, 'Other Ledger 1', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$G$2:$G$${finalAdditionalLedgers.length + 1}`]
      });
      applyDropdownByHeader(templateSheet, headers, 'Other Ledger 2', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$G$2:$G$${finalAdditionalLedgers.length + 1}`]
      });
      applyDropdownByHeader(templateSheet, headers, 'Discount Ledger', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$G$2:$G$${finalAdditionalLedgers.length + 1}`]
      });
      applyDropdownByHeader(templateSheet, headers, 'Round Off Ledger', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$H$2:$H$${finalRoundOffLedgers.length + 1}`]
      });
      applyDropdownByHeader(templateSheet, headers, 'Voucher Mode', {
        type: 'list',
        allowBlank: true,
        formulae: ['"Auto,Item Invoice,Accounting"']
      });
      applyDropdownByHeader(templateSheet, headers, 'Inventory Mode', {
        type: 'list',
        allowBlank: true,
        formulae: ['"No Inventory,Inventory Optional,Inventory Mandatory"']
      });

      // Auto-adjust column widths for Template Sheet
      templateSheet.columns.forEach(col => {
        let maxLen = 0;
        col.eachCell({ includeEmpty: true }, (cell) => {
          const valStr = cell.value ? String(cell.value) : '';
          if (valStr.length > maxLen) maxLen = valStr.length;
        });
        col.width = Math.max(maxLen + 3, 12);
      });

      // Auto-adjust column widths for Tally_Masters
      mastersSheet.columns.forEach(col => {
        let maxLen = 0;
        col.eachCell({ includeEmpty: true }, (cell) => {
          const valStr = cell.value ? String(cell.value) : '';
          if (valStr.length > maxLen) maxLen = valStr.length;
        });
        col.width = Math.max(maxLen + 3, 15);
      });

      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Tally_${realVoucherType}_${isVoucherwise ? 'Voucherwise' : 'Itemwise'}_Template.xlsx`;
      a.click();
      window.URL.revokeObjectURL(url);
      return;
    }

    if (type.toLowerCase() === 'journal') {
      const headers = [
        'Voucher Date', 'Voucher No', 'Ledger Name', 'Dr/Cr', 'Amount', 'Narration', 'Reference', 'Cost Centre', 'Bill Reference', 'Remarks'
      ];
      const sampleData = [
        {
          'Voucher Date': new Date().toLocaleDateString('en-GB'),
          'Voucher No': 'JV-001',
          'Ledger Name': 'Rent Expense',
          'Dr/Cr': 'Dr',
          'Amount': 25000,
          'Narration': 'Rent entry',
          'Reference': 'Ref1',
          'Cost Centre': '',
          'Bill Reference': '',
          'Remarks': ''
        },
        {
          'Voucher Date': new Date().toLocaleDateString('en-GB'),
          'Voucher No': 'JV-001',
          'Ledger Name': 'HDFC Bank',
          'Dr/Cr': 'Cr',
          'Amount': 25000,
          'Narration': 'Rent entry',
          'Reference': 'Ref1',
          'Cost Centre': '',
          'Bill Reference': '',
          'Remarks': ''
        }
      ];

      const wb = new ExcelJS.Workbook();
      const templateSheet = wb.addWorksheet('Template');
      templateSheet.getRow(1).values = headers;
      templateSheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFF' } };
      templateSheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '1F2937' }
      };

      sampleData.forEach(row => {
        const rowValues = headers.map(h => row[h as keyof typeof row] !== undefined ? row[h as keyof typeof row] : '');
        templateSheet.addRow(rowValues);
      });

      const DEFAULT_LEDGERS = ['Rent Expense', 'HDFC Bank', 'Sundry Debtors', 'Sundry Creditors', 'Sales Account', 'Purchase Account', 'Cash-in-hand', 'CGST Output 9%', 'SGST Output 9%', 'CGST Input 9%', 'SGST Input 9%', 'Freight Charges', 'Packing Expenses', 'Round Off A/c'];
      const ledgersList = (tallyContext && tallyContext.ledgers && tallyContext.ledgers.length > 0)
        ? tallyContext.ledgers
        : DEFAULT_LEDGERS;

      const mastersSheet = wb.addWorksheet('Tally_Masters');
      mastersSheet.state = 'hidden';
      mastersSheet.getRow(1).values = ['Ledgers', 'DrCr'];
      ledgersList.forEach((l, idx) => {
        mastersSheet.getCell(`A${idx + 2}`).value = l;
      });
      mastersSheet.getCell('B2').value = 'Dr';
      mastersSheet.getCell('B3').value = 'Cr';

      applyDropdownByHeader(templateSheet, headers, 'Ledger Name', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$A$2:$A$${ledgersList.length + 1}`]
      });

      applyDropdownByHeader(templateSheet, headers, 'Dr/Cr', {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$B$2:$B$3`]
      });

      // Auto-adjust column widths
      templateSheet.columns.forEach(col => {
        let maxLen = 0;
        col.eachCell({ includeEmpty: true }, (cell) => {
          const valStr = cell.value ? String(cell.value) : '';
          if (valStr.length > maxLen) maxLen = valStr.length;
        });
        col.width = Math.max(maxLen + 3, 12);
      });

      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Tally_Journal_Template.xlsx`;
      a.click();
      window.URL.revokeObjectURL(url);
      return;
    }

    const headers = ['Date', 'Particulars', 'Voucher Type', 'Voucher No', 'Amount', 'Narration', 'Reference'];
    const sampleData = [
      {
        'Date': new Date().toLocaleDateString('en-GB'),
        'Particulars': 'Sample Party Name',
        'Voucher Type': type,
        'Voucher No': '1',
        'Amount': '1000',
        'Narration': `Sample ${type} entry`,
        'Reference': 'REF001'
      }
    ];

    const wb = new ExcelJS.Workbook();
    const templateSheet = wb.addWorksheet('Template');
    templateSheet.getRow(1).values = headers;
    sampleData.forEach(row => {
      const rowValues = headers.map(h => row[h as keyof typeof row] !== undefined ? row[h as keyof typeof row] : '');
      templateSheet.addRow(rowValues);
    });

    const ledgersList = (tallyContext && tallyContext.ledgers && tallyContext.ledgers.length > 0)
      ? tallyContext.ledgers
      : DEFAULT_LEDGERS;

    const mastersSheet = wb.addWorksheet('Tally_Masters');
    mastersSheet.state = 'hidden';
    mastersSheet.getRow(1).values = ['Ledgers'];
    ledgersList.forEach((l, idx) => {
      mastersSheet.getRow(idx + 2).values = [l];
    });

    for (let r = 2; r <= 100; r++) {
      // B: Particulars -> Col A of Tally_Masters
      templateSheet.getCell(`B${r}`).dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: [`Tally_Masters!$A$2:$A$${ledgersList.length + 1}`]
      };
      // C: Voucher Type -> list
      templateSheet.getCell(`C${r}`).dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: ['"Payment,Receipt,Contra,Journal"']
      };
    }

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Tally_${type}_Template.xlsx`;
    a.click();
    window.URL.revokeObjectURL(url);
  };

  const downloadMasterTemplate = (type: string) => {
    let headers: string[] = [];
    let sampleData: any[] = [];
    let sheetName = "Template";

    if (type === 'Ledger') {
      headers = [
        'Ledger Name', 'Under Group', 'Opening Balance', 'Dr/Cr', 'Mailing Name', 
        'Address Line 1', 'Address Line 2', 'State', 'Country', 'Pincode', 
        'PAN', 'GSTIN', 'Registration Type', 'Taxability', 'Is Billwise On', 
        'Is Cost Centre On', 'Email', 'Mobile Number'
      ];
      sampleData = [
        {
          'Ledger Name': 'Acme Corporation',
          'Under Group': 'Sundry Debtors',
          'Opening Balance': '5000',
          'Dr/Cr': 'Dr',
          'Mailing Name': 'Acme Corporation Pvt Ltd',
          'Address Line 1': '123 Business Park',
          'Address Line 2': 'Sector 62',
          'State': 'Maharashtra',
          'Country': 'INDIA',
          'Pincode': '400001',
          'PAN': 'ABCDE1234F',
          'GSTIN': '27ABCDE1234F1Z5',
          'Registration Type': 'Regular',
          'Taxability': 'Taxable',
          'Is Billwise On': 'Yes',
          'Is Cost Centre On': 'No',
          'Email': 'info@acme.com',
          'Mobile Number': '9876543210'
        }
      ];
      sheetName = "Ledger_Template";
    } else if (type === 'StockItem') {
      headers = [
        'Stock Item Name', 'Under Stock Group', 'Unit', 'Opening Quantity', 
        'Opening Rate', 'Opening Value', 'HSN/SAC', 'GST Applicable', 
        'Taxability', 'GST Rate', 'CGST Rate', 'SGST Rate', 'IGST Rate', 'Description'
      ];
      sampleData = [
        {
          'Stock Item Name': 'Widget A',
          'Under Stock Group': 'Primary',
          'Unit': 'NOS',
          'Opening Quantity': '100',
          'Opening Rate': '50',
          'Opening Value': '5000',
          'HSN/SAC': '84713010',
          'GST Applicable': 'Applicable',
          'Taxability': 'Taxable',
          'GST Rate': '18',
          'CGST Rate': '9',
          'SGST Rate': '9',
          'IGST Rate': '18',
          'Description': 'High quality widget'
        }
      ];
      sheetName = "Stock_Item_Template";
    } else if (type === 'StockGroup') {
      headers = ['Stock Group Name', 'Under Parent Group'];
      sampleData = [
        { 'Stock Group Name': 'Electronics', 'Under Parent Group': 'Primary' }
      ];
      sheetName = "Stock_Group_Template";
    } else if (type === 'Unit') {
      headers = ['Symbol', 'Formal Name', 'Unit Quantity Code (UQC)', 'Number of Decimal Places'];
      sampleData = [
        { 'Symbol': 'NOS', 'Formal Name': 'Numbers', 'Unit Quantity Code (UQC)': 'NOS-NUMBERS', 'Number of Decimal Places': '0' }
      ];
      sheetName = "Unit_Template";
    }

    const ws = XLSX.utils.json_to_sheet(sampleData, { header: headers });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    if (type === 'Ledger') {
      const groupsList = (tallyContext && tallyContext.groups && tallyContext.groups.length > 0)
        ? tallyContext.groups
        : DEFAULT_GROUPS;

      const mastersWS = buildMastersSheet({ 'Groups': groupsList });
      XLSX.utils.book_append_sheet(wb, mastersWS, "Tally_Masters");

      ws['!dataValidation'] = [
        {
          sqref: 'B2:B100',
          type: 'list',
          allowBlank: true,
          formula1: `Tally_Masters!$A$2:$A$${groupsList.length + 1}`
        }
      ];
    } else if (type === 'StockItem') {
      const stockGroupsList = (tallyContext && tallyContext.stockGroups && tallyContext.stockGroups.length > 0)
        ? tallyContext.stockGroups
        : DEFAULT_STOCK_GROUPS;

      const unitsList = (tallyContext && tallyContext.units && tallyContext.units.length > 0)
        ? tallyContext.units
        : DEFAULT_UNITS;

      const mastersWS = buildMastersSheet({
        'StockGroups': stockGroupsList,
        'Units': unitsList
      });
      XLSX.utils.book_append_sheet(wb, mastersWS, "Tally_Masters");

      ws['!dataValidation'] = [
        {
          sqref: 'B2:B100',
          type: 'list',
          allowBlank: true,
          formula1: `Tally_Masters!$A$2:$A$${stockGroupsList.length + 1}`
        },
        {
          sqref: 'C2:C100',
          type: 'list',
          allowBlank: true,
          formula1: `Tally_Masters!$B$2:$B$${unitsList.length + 1}`
        }
      ];
    } else if (type === 'StockGroup') {
      const stockGroupsList = (tallyContext && tallyContext.stockGroups && tallyContext.stockGroups.length > 0)
        ? tallyContext.stockGroups
        : DEFAULT_STOCK_GROUPS;

      const mastersWS = buildMastersSheet({ 'StockGroups': stockGroupsList });
      XLSX.utils.book_append_sheet(wb, mastersWS, "Tally_Masters");

      ws['!dataValidation'] = [
        {
          sqref: 'B2:B100',
          type: 'list',
          allowBlank: true,
          formula1: `Tally_Masters!$A$2:$A$${stockGroupsList.length + 1}`
        }
      ];
    }

    XLSX.writeFile(wb, `Tally_${type}_Template.xlsx`);
  };

  const downloadMappedExcel = (transactions: MappedTransaction[], fileName: string) => {
    const headers = ['Date', 'Particulars', 'Voucher Type', 'Voucher No', 'Amount', 'Narration', 'Reference'];
    const data = transactions.map((tx, idx) => ({
      'Date': tx.date,
      'Particulars': tx.tallyLedger,
      'Voucher Type': selectedVoucherType,
      'Voucher No': (idx + 1).toString(),
      'Amount': tx.amount.toString(),
      'Narration': tx.description,
      'Reference': tx.reference || ''
    }));

    const ws = XLSX.utils.json_to_sheet(data, { header: headers });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Mapped Data");
    XLSX.writeFile(wb, fileName.replace(/\.[^/.]+$/, "") + "_Mapped.xlsx");
  };


  // --- Auth ---
  useEffect(() => {
    if (getAppMode() === 'desktop-offline') {
      setUser({
        uid: OFFLINE_USER.uid,
        displayName: OFFLINE_USER.displayName,
        email: OFFLINE_USER.email,
        emailVerified: true,
        isAnonymous: false,
        metadata: {},
        providerData: [],
        tenantId: null,
        delete: async () => {},
        getIdToken: async () => '',
        getIdTokenResult: async () => ({} as any),
        reload: async () => {},
        toJSON: () => ({}),
        phoneNumber: null,
        photoURL: null,
        providerId: 'firebase',
      } as unknown as User);
      setLoading(false);
      return;
    }

    const unsubscribe = onAuthStateChanged(auth, (u) => {
      setUser(u);
      setLoading(false);
    });
    return () => unsubscribe();
  }, []);

  const handleLogin = async () => {
    const provider = new GoogleAuthProvider();
    try {
      await signInWithPopup(auth, provider);
    } catch (err) {
      console.error("Login failed", err);
    }
  };

  const handleLogout = () => signOut(auth);

  const handleRestartWorkspace = async () => {
    if (!user) return;
    
    // Set guard flag BEFORE clearing state to block stale snapshots
    isResettingWorkspaceRef.current = true;

    // Clear all local states immediately
    setPendingData([]);
    setPendingFileName('');
    setMappings([]);
    setError(null);
    setSelectedVoucherType('Payment');
    setSelectedBankLedger('');
    setAiMappedTransactions([]);
    setImportType('Voucher');
    setParsedLedgers([]);
    setParsedStockItems([]);
    setParsedStockGroups([]);
    setParsedUnits([]);
    setMasterReviewSearch('');
    setMasterReviewFilter('all');
    setTallyContext(null);
    setSkippedContext(false);
    setProceededWithContext(false);
    setShowAllLedgers(false);
    setCurrentStep('upload');
    setShowRestartConfirm(false);

    try {
      setIsProcessing(true);
      if (getAppMode() === 'desktop-offline') {
        await clearOfflineWorkspace(user.uid);
        isResettingWorkspaceRef.current = false;
        setError("Offline workspace cleared successfully.");
        return;
      }
      // Delete the active Tally context document from Firestore
      await deleteDoc(doc(db, 'tally_context', user.uid));
      isResettingWorkspaceRef.current = false;
    } catch (err: any) {
      console.error("Failed to delete cloud Tally context from Firestore", {
        code: err?.code,
        message: err?.message,
        error: err
      });
      
      // Attempt fallback cloud reset using setDoc with empty safe context
      try {
        const fallbackContext = {
          uid: user.uid,
          ledgers: [],
          groups: [],
          ledgerGroupMap: {},
          historicalMappings: [],
          stockGroups: [],
          stockItems: [],
          units: [],
          lastUpdated: serverTimestamp()
        };
        await setDoc(doc(db, 'tally_context', user.uid), fallbackContext);
        
        // Fallback overwrite succeeded
        setError("Workspace reset completed. Cloud context was cleared using fallback reset.");
        isResettingWorkspaceRef.current = false;
      } catch (fallbackErr: any) {
        console.error("Fallback cloud reset also failed", {
          code: fallbackErr?.code,
          message: fallbackErr?.message,
          error: fallbackErr
        });
        
        setError("Local workspace cleared, but cloud context could not be cleared. Please deploy updated Firestore rules.");
        // Note: isResettingWorkspaceRef.current remains true to prevent onSnapshot from reloading stale data!
      }
    } finally {
      setIsProcessing(false);
    }
  };

  // --- Real-time Data ---
  useEffect(() => {
    if (!user) {
      setConversions([]);
      setTallyContext(null);
      return;
    }

    // Conversions
    const unsubscribeConversions = subscribeToConversions(user.uid, (docs) => {
      setConversions(docs as ConversionRecord[]);
    });

    // Tally Context
    const unsubscribeContext = subscribeToTallyContext(user.uid, (data) => {
      if (isResettingWorkspaceRef.current) {
        if (!data) {
          setTallyContext(null);
          setProceededWithContext(false);
          isResettingWorkspaceRef.current = false;
        } else if (!data.ledgers || data.ledgers.length === 0) {
          setTallyContext(data as TallyContext);
          setProceededWithContext(false);
          isResettingWorkspaceRef.current = false;
        }
        return;
      }

      if (data) {
        setTallyContext(data as TallyContext);
        setProceededWithContext(true);
      } else {
        setTallyContext(null);
        setProceededWithContext(false);
      }
    });

    return () => {
      unsubscribeConversions();
      unsubscribeContext();
    };
  }, [user]);

  // --- Tally XML Parsing Helper ---
  const extractAllMastersFromXml = (result: any) => {
    const ledgers: string[] = [];
    const groups: string[] = [];
    const stockGroups: string[] = [];
    const stockItems: string[] = [];
    const units: string[] = [];
    const ledgerGroupMap: Record<string, string> = {};
    const stockItemStockGroupMap: Record<string, string> = {};
    const groupParentMap: Record<string, string> = {};
    const ledgerDetails: LedgerMasterRow[] = [];
    const stockItemDetails: StockMasterRow[] = [];

    const getParentName = (obj: any): string => {
      if (!obj) return '';
      if (typeof obj === 'string') return obj;
      if (typeof obj === 'object') {
        return obj['#text'] || obj['NAME'] || obj['@_NAME'] || '';
      }
      return String(obj);
    };

    const traverse = (obj: any) => {
      if (!obj) return;

      if (obj.LEDGER) {
        const list = Array.isArray(obj.LEDGER) ? obj.LEDGER : [obj.LEDGER];
        list.forEach((l: any, idx: number) => {
          const name = l['@_NAME'] || l.NAME;
          if (name) {
            const nameStr = String(name);
            ledgers.push(nameStr);
            const parent = l.PARENT || l['@_PARENT'];
            const parentStr = getParentName(parent);
            if (parentStr) ledgerGroupMap[nameStr] = parentStr;

            const registrationType = String(l.GSTREGISTRATIONTYPE || l['@_GSTREGISTRATIONTYPE'] || l.REGISTRATIONTYPE || '').trim();
            const gstin = String(l.PARTYGSTIN || l['@_PARTYGSTIN'] || l.GSTIN || l['@_GSTIN'] || '').trim();
            const state = String(l.STATENAME || l['@_STATENAME'] || l.STATE || '').trim();
            
            let address1 = '';
            let address2 = '';
            if (l.ADDRESS) {
              const addrList = Array.isArray(l.ADDRESS) ? l.ADDRESS : [l.ADDRESS];
              const resolvedLines: string[] = addrList.map((addr: any) => {
                if (typeof addr === 'string') return addr;
                if (typeof addr === 'object' && addr) return addr['#text'] || '';
                return '';
              }).filter(Boolean);
              address1 = resolvedLines[0] || '';
              address2 = resolvedLines[1] || resolvedLines.slice(1).join(', ') || '';
            }

            ledgerDetails.push({
              rowNum: idx + 1,
              ledgerName: nameStr,
              underGroup: parentStr || 'Sundry Debtors',
              address1,
              address2,
              state,
              gstin,
              registrationType,
              isValid: true,
              errors: [],
              warnings: [],
              isDuplicate: false,
              isPossibleDuplicate: false,
              excluded: false
            });
          }
        });
      }
      if (obj.GROUP) {
        const list = Array.isArray(obj.GROUP) ? obj.GROUP : [obj.GROUP];
        list.forEach((g: any) => {
          const name = g['@_NAME'] || g.NAME;
          if (name) {
            const nameStr = String(name);
            groups.push(nameStr);
            const parent = g.PARENT || g['@_PARENT'];
            const parentStr = getParentName(parent);
            if (parentStr) {
              groupParentMap[nameStr] = parentStr;
            }
          }
        });
      }
      if (obj.STOCKGROUP) {
        const list = Array.isArray(obj.STOCKGROUP) ? obj.STOCKGROUP : [obj.STOCKGROUP];
        list.forEach((sg: any) => {
          const name = sg['@_NAME'] || sg.NAME;
          if (name) stockGroups.push(String(name));
        });
      }
      if (obj.STOCKITEM) {
        const list = Array.isArray(obj.STOCKITEM) ? obj.STOCKITEM : [obj.STOCKITEM];
        list.forEach((si: any, idx: number) => {
          const name = si['@_NAME'] || si.NAME;
          if (name) {
            const nameStr = String(name);
            stockItems.push(nameStr);
            const parentGroup = si.PARENT ? String(si.PARENT) : '';
            if (parentGroup) stockItemStockGroupMap[nameStr] = parentGroup;

            const unit = String(si.BASEUNITS || si['@_BASEUNITS'] || si.UNIT || 'NOS').trim();
            
            // Recursive key finder helper
            const findKeyInObject = (o: any, keyName: string): any => {
              if (!o) return undefined;
              if (typeof o !== 'object') return undefined;
              if (o[keyName] !== undefined) return o[keyName];
              for (const k of Object.keys(o)) {
                const val = o[k];
                if (typeof val === 'object') {
                  const found = findKeyInObject(val, keyName);
                  if (found !== undefined) return found;
                }
              }
              return undefined;
            };

            const hsn = String(findKeyInObject(si, 'HSNCODE') || findKeyInObject(si, 'HSN') || findKeyInObject(si, 'GSTHSNNAME') || '').trim();
            let gstRateVal = '';
            const rateFound = findKeyInObject(si, 'GSTRATE') || findKeyInObject(si, 'RATE');
            if (rateFound !== undefined) {
              if (typeof rateFound === 'string') {
                gstRateVal = rateFound;
              } else if (typeof rateFound === 'number') {
                gstRateVal = String(rateFound);
              } else if (typeof rateFound === 'object' && rateFound) {
                gstRateVal = rateFound['#text'] || '';
              }
            }

            stockItemDetails.push({
              rowNum: idx + 1,
              itemName: nameStr,
              underGroup: parentGroup,
              unit,
              hsn,
              gstRate: gstRateVal,
              isValid: true,
              errors: [],
              warnings: [],
              isDuplicate: false,
              isPossibleDuplicate: false,
              excluded: false
            });
          }
        });
      }
      if (obj.UNIT) {
        const list = Array.isArray(obj.UNIT) ? obj.UNIT : [obj.UNIT];
        list.forEach((u: any) => {
          const name = u['@_NAME'] || u.NAME;
          if (name) units.push(String(name));
        });
      }

      if (Array.isArray(obj)) {
        obj.forEach(item => traverse(item));
      } else if (typeof obj === 'object') {
        Object.values(obj).forEach(val => traverse(val));
      }
    };

    traverse(result);

    return {
      ledgers,
      groups,
      stockGroups,
      stockItems,
      units,
      ledgerGroupMap,
      stockItemStockGroupMap,
      groupParentMap,
      ledgerDetails,
      stockItemDetails,
    };
  };

  // --- Tally XML Parsing ---
  const handleTallyMastersUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file || !user) return;

    isResettingWorkspaceRef.current = false;
    setIsContextLoading(true);
    setError(null);
    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const xml = event.target?.result as string;
        const parser = new XMLParser({ ignoreAttributes: false });
        const result = parser.parse(xml);
        
        const parsed = extractAllMastersFromXml(result);

        const existingMappings = tallyContext?.historicalMappings || [];

        const contextPayload: any = {
          uid: user.uid,
          ledgers: Array.from(new Set(parsed.ledgers)),
          groups: Array.from(new Set(parsed.groups)),
          stockGroups: Array.from(new Set(parsed.stockGroups)),
          stockItems: Array.from(new Set(parsed.stockItems)),
          units: Array.from(new Set(parsed.units)),
          ledgerGroupMap: parsed.ledgerGroupMap,
          stockItemStockGroupMap: parsed.stockItemStockGroupMap,
          groupParentMap: parsed.groupParentMap || {},
          historicalMappings: existingMappings,
          ledgerDetails: parsed.ledgerDetails || [],
          stockItemDetails: parsed.stockItemDetails || [],
        };
        if (getAppMode() === 'web') {
          contextPayload.lastUpdated = serverTimestamp();
        }
        await saveTallyContext(user.uid, contextPayload);

        setIsContextLoading(false);
      } catch (err) {
        console.error("Failed to parse Tally Masters XML", err);
        setError("Failed to parse Tally Masters XML. Please ensure it is a valid Tally XML file.");
        setIsContextLoading(false);
      }
    };
    reader.readAsText(file);
  };

  const handleDaybookUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file || !user) return;

    isResettingWorkspaceRef.current = false;
    setIsContextLoading(true);
    setError(null);
    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const xml = event.target?.result as string;
        const parser = new XMLParser({ ignoreAttributes: false });
        const result = parser.parse(xml);
        
        const ledgerNamesFromEntries: string[] = [];
        const historicalMappings: { narration: string; ledger: string }[] = [];

        const traverseForVouchers = (obj: any) => {
          if (!obj) return;

          if (obj.VOUCHER) {
            const list = Array.isArray(obj.VOUCHER) ? obj.VOUCHER : [obj.VOUCHER];
            list.forEach((v: any) => {
              const narration = v.NARRATION || '';
              const entriesList = v.ALLLEDGERENTRIES_LIST || v['ALLLEDGERENTRIES.LIST'];
              if (entriesList) {
                const entries = Array.isArray(entriesList) ? entriesList : [entriesList];
                entries.forEach((ent: any) => {
                  if (ent.LEDGERNAME) {
                    ledgerNamesFromEntries.push(String(ent.LEDGERNAME));
                  }
                });

                if (narration) {
                  const primaryEntry = entries.find((e: any) => e.LEDGERNAME && !e.LEDGERNAME.toLowerCase().includes('bank') && !e.LEDGERNAME.toLowerCase().includes('cash'));
                  if (primaryEntry) {
                    historicalMappings.push({ narration: String(narration), ledger: String(primaryEntry.LEDGERNAME) });
                  }
                }
              }
            });
          }

          if (Array.isArray(obj)) {
            obj.forEach(item => traverseForVouchers(item));
          } else if (typeof obj === 'object') {
            Object.values(obj).forEach(val => traverseForVouchers(val));
          }
        };

        traverseForVouchers(result);

        const currentLedgers = tallyContext?.ledgers || [];
        const mergedLedgers = Array.from(new Set([...currentLedgers, ...ledgerNamesFromEntries]));

        const currentGroups = tallyContext?.groups || [];
        const currentStockGroups = tallyContext?.stockGroups || [];
        const currentStockItems = tallyContext?.stockItems || [];
        const currentUnits = tallyContext?.units || [];
        const currentLedgerGroupMap = tallyContext?.ledgerGroupMap || {};
        const currentStockItemStockGroupMap = tallyContext?.stockItemStockGroupMap || {};
        const currentGroupParentMap = tallyContext?.groupParentMap || {};
        const currentHistorical = tallyContext?.historicalMappings || [];
        
        const mergedHistorical = [...currentHistorical, ...historicalMappings].slice(0, 500);

        const contextPayload: any = {
          uid: user.uid,
          ledgers: mergedLedgers,
          groups: currentGroups,
          stockGroups: currentStockGroups,
          stockItems: currentStockItems,
          units: currentUnits,
          ledgerGroupMap: currentLedgerGroupMap,
          stockItemStockGroupMap: currentStockItemStockGroupMap,
          groupParentMap: currentGroupParentMap,
          historicalMappings: mergedHistorical,
          ledgerDetails: tallyContext?.ledgerDetails || [],
          stockItemDetails: tallyContext?.stockItemDetails || [],
        };
        if (getAppMode() === 'web') {
          contextPayload.lastUpdated = serverTimestamp();
        }
        await saveTallyContext(user.uid, contextPayload);

        setIsContextLoading(false);
      } catch (err) {
        console.error("Failed to parse Daybook XML", err);
        setError("Failed to parse Daybook XML. Please ensure it is a valid Tally XML file.");
        setIsContextLoading(false);
      }
    };
    reader.readAsText(file);
  };

  const handleCombinedTallyUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0 || !user) return;

    isResettingWorkspaceRef.current = false;
    setIsContextLoading(true);
    setError(null);
    
    const ledgers: string[] = [];
    const groups: string[] = [];
    const stockGroups: string[] = [];
    const stockItems: string[] = [];
    const units: string[] = [];
    const ledgerGroupMap: Record<string, string> = {};
    const stockItemStockGroupMap: Record<string, string> = {};
    const groupParentMap: Record<string, string> = {};
    const historicalMappings: { narration: string; ledger: string }[] = [];

    const processFile = (file: File): Promise<void> => {
      return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (event) => {
          try {
            const xml = event.target?.result as string;
            const parser = new XMLParser({ ignoreAttributes: false });
            const result = parser.parse(xml);

            // Parse Masters
            const parsed = extractAllMastersFromXml(result);
            ledgers.push(...parsed.ledgers);
            groups.push(...parsed.groups);
            stockGroups.push(...parsed.stockGroups);
            stockItems.push(...parsed.stockItems);
            units.push(...parsed.units);
            Object.assign(ledgerGroupMap, parsed.ledgerGroupMap);
            Object.assign(stockItemStockGroupMap, parsed.stockItemStockGroupMap);
            if (parsed.groupParentMap) {
              Object.assign(groupParentMap, parsed.groupParentMap);
            }

            // Recursively extract Vouchers
            const traverseVouchers = (obj: any) => {
              if (!obj) return;
              if (obj.VOUCHER) {
                const list = Array.isArray(obj.VOUCHER) ? obj.VOUCHER : [obj.VOUCHER];
                list.forEach((v: any) => {
                  const narration = v.NARRATION || '';
                  const entriesList = v.ALLLEDGERENTRIES_LIST || v['ALLLEDGERENTRIES.LIST'];
                  if (entriesList) {
                    const entries = Array.isArray(entriesList) ? entriesList : [entriesList];
                    entries.forEach((ent: any) => {
                      if (ent.LEDGERNAME) {
                        ledgers.push(String(ent.LEDGERNAME));
                      }
                    });
                    if (narration) {
                      const primaryEntry = entries.find((e: any) => e.LEDGERNAME && !e.LEDGERNAME.toLowerCase().includes('bank') && !e.LEDGERNAME.toLowerCase().includes('cash'));
                      if (primaryEntry) {
                        historicalMappings.push({ narration: String(narration), ledger: String(primaryEntry.LEDGERNAME) });
                      }
                    }
                  }
                });
              }
              if (Array.isArray(obj)) {
                obj.forEach(item => traverseVouchers(item));
              } else if (typeof obj === 'object') {
                Object.values(obj).forEach(val => traverseVouchers(val));
              }
            };

            traverseVouchers(result);
            resolve();
          } catch (err) {
            console.error("Failed to parse combined Tally XML", err);
            resolve();
          }
        };
        reader.readAsText(file);
      });
    };

    await Promise.all(Array.from(files).map(processFile));

    const contextPayload: any = {
      uid: user.uid,
      ledgers: Array.from(new Set(ledgers)),
      groups: Array.from(new Set(groups)),
      stockGroups: Array.from(new Set(stockGroups)),
      stockItems: Array.from(new Set(stockItems)),
      units: Array.from(new Set(units)),
      ledgerGroupMap,
      stockItemStockGroupMap,
      groupParentMap,
      historicalMappings: historicalMappings.slice(0, 500),
      ledgerDetails: tallyContext?.ledgerDetails || [],
      stockItemDetails: tallyContext?.stockItemDetails || [],
    };
    if (getAppMode() === 'web') {
      contextPayload.lastUpdated = serverTimestamp();
    }
    await saveTallyContext(user.uid, contextPayload);

    setIsContextLoading(false);
  };

  // --- Optimized PDF Parsing ---
  const extractTextFromPdf = async (file: File): Promise<string> => {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ 
      data: arrayBuffer,
      verbosity: 0 // Reduce logging for speed
    }).promise;
    
    // Process pages in parallel for speed
    const pagePromises = [];
    for (let i = 1; i <= Math.min(pdf.numPages, 20); i++) { // Limit to 20 pages for performance
      pagePromises.push((async () => {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        return textContent.items.map((item: any) => item.str).join(' ');
      })());
    }
    const pages = await Promise.all(pagePromises);
    return pages.join('\n');
  };

  // --- Local Deterministic Bank Statement Helpers ---
  const suggestLedgerForNarrationEnhanced = (
    narration: string,
    context: TallyContext | null
  ): { ledgerName: string; confidence: number; reasoning: string } => {
    if (!context || !narration) {
      return { ledgerName: '', confidence: 0, reasoning: 'No narration or context' };
    }

    const cleanNarr = narration.toLowerCase().trim();
    const normNarr = cleanNarr.replace(/[^a-z0-9]/g, '');

    const isValidLedger = (name: string) => context.ledgers.includes(name);

    let candidateLedger = '';
    let candidateConfidence = 0;
    let candidateReasoning = 'No match found.';

    // 1. Exact historical mapping from Daybook
    if (context.historicalMappings) {
      const exactHist = context.historicalMappings.find(h => h.narration.toLowerCase().trim() === cleanNarr);
      if (exactHist && isValidLedger(exactHist.ledger)) {
        candidateLedger = exactHist.ledger;
        candidateConfidence = 1.0;
        candidateReasoning = 'Matched exact historical narration from Daybook';
      }
    }

    // 2. Normalized historical match
    if (!candidateLedger && context.historicalMappings) {
      const normHist = context.historicalMappings.find(h => h.narration.toLowerCase().replace(/[^a-z0-9]/g, '') === normNarr);
      if (normHist && isValidLedger(normHist.ledger)) {
        candidateLedger = normHist.ledger;
        candidateConfidence = 0.95;
        candidateReasoning = 'Matched normalized historical narration';
      }
    }

    // 3. Keyword match (exact ledger name in narration)
    if (!candidateLedger) {
      const sortedLedgers = [...context.ledgers].sort((a, b) => b.length - a.length);
      for (const ledger of sortedLedgers) {
        const ledgerLower = ledger.toLowerCase().trim();
        if (ledgerLower.length > 3 && cleanNarr.includes(ledgerLower)) {
          candidateLedger = ledger;
          candidateConfidence = 0.90;
          candidateReasoning = `Found exact ledger name "${ledger}" in narration`;
          break;
        }
      }
    }

    // 4. Fuzzy match (getSimilarity score >= 0.8)
    if (!candidateLedger && context.historicalMappings) {
      let bestMatch: { ledger: string; score: number } | null = null;
      for (const h of context.historicalMappings) {
        const score = getSimilarity(h.narration, narration);
        if (score >= 0.8) {
          if (!bestMatch || score > bestMatch.score) {
            bestMatch = { ledger: h.ledger, score };
          }
        }
      }
      if (bestMatch && isValidLedger(bestMatch.ledger)) {
        candidateLedger = bestMatch.ledger;
        candidateConfidence = Math.round(bestMatch.score * 100) / 100;
        candidateReasoning = `Fuzzy matched historical narration with ${(bestMatch.score * 100).toFixed(0)}% similarity`;
      }
    }

    // 5. Ledger name words found inside narration
    if (!candidateLedger) {
      const sortedLedgers = [...context.ledgers].sort((a, b) => b.length - a.length);
      for (const ledger of sortedLedgers) {
        const ledgerWords = ledger.toLowerCase().split(/[^a-zA-Z0-9]/).filter(w => w.length > 3);
        if (ledgerWords.length > 0 && ledgerWords.every(w => cleanNarr.includes(w))) {
          candidateLedger = ledger;
          candidateConfidence = 0.75;
          candidateReasoning = `All descriptive words of "${ledger}" found in narration`;
          break;
        }
      }
    }

    const suspenseLedger = context.ledgers.find(l => /^suspense$/i.test(l) || /^suspense account$/i.test(l));

    // Rule: Where system is less than 90% sure about ledger name, select Suspense Account (if available).
    if (candidateLedger && candidateConfidence >= 0.90) {
      return {
        ledgerName: candidateLedger,
        confidence: candidateConfidence,
        reasoning: candidateReasoning
      };
    } else {
      if (suspenseLedger) {
        return {
          ledgerName: suspenseLedger,
          confidence: candidateLedger ? candidateConfidence : 0,
          reasoning: "Low confidence. Suspense selected. Please review."
        };
      } else {
        return {
          ledgerName: '',
          confidence: candidateLedger ? candidateConfidence : 0,
          reasoning: "Low confidence. Please review. (Needs Review)"
        };
      }
    }
  };

  const suggestLedgerForNarration = (narration: string, context: TallyContext | null): string => {
    return suggestLedgerForNarrationEnhanced(narration, context).ledgerName;
  };

  const autoDetectBankStatement = (rows: any[][]) => {
    const DATE_ALIASES = ['date', 'txn date', 'transaction date', 'value date', 'posting date', 'entry date'];
    const NARRATION_ALIASES = ['narration', 'description', 'transaction details', 'details', 'remarks', 'description/narration'];
    const DEBIT_ALIASES = ['debit', 'withdrawal', 'withdrawals', 'paid out', 'dr', 'debit amount', 'withdrawal amount', 'payment', 'payment amount'];
    const CREDIT_ALIASES = ['credit', 'deposit', 'deposits', 'paid in', 'cr', 'credit amount', 'deposit amount', 'receipt', 'receipt amount'];
    const AMOUNT_ALIASES = ['amount', 'transaction amount', 'txn amount', 'value', 'amt'];
    const DR_CR_ALIASES = ['dr/cr', 'debit/credit', 'cr/dr', 'type', 'transaction type', 'amount type', 'd/c'];
    const BALANCE_ALIASES = ['balance', 'closing balance', 'running balance', 'available balance', 'ledger balance'];
    const REFERENCE_ALIASES = ['reference', 'ref no', 'utr', 'cheque no', 'instrument no', 'transaction id', 'txn id', 'bank ref', 'chq/ref no'];

    let bestHeaderRowIdx = 0;
    let maxScore = -1;

    const scanLimit = Math.min(50, rows.length);
    for (let r = 0; r < scanLimit; r++) {
      const row = rows[r];
      if (!row || row.length === 0) continue;

      let score = 0;
      let hasDateLike = false;
      let hasNarrLike = false;
      let hasDebitCreditOrAmt = false;
      let hasBalance = false;
      let hasRef = false;

      row.forEach(val => {
        if (val === undefined || val === null) return;
        const str = String(val).toLowerCase().trim();
        if (!str) return;

        if (DATE_ALIASES.includes(str)) {
          score += 10;
          hasDateLike = true;
        } else if (NARRATION_ALIASES.includes(str)) {
          score += 10;
          hasNarrLike = true;
        } else if (DEBIT_ALIASES.includes(str) || CREDIT_ALIASES.includes(str) || AMOUNT_ALIASES.includes(str)) {
          score += 10;
          hasDebitCreditOrAmt = true;
        } else if (BALANCE_ALIASES.includes(str)) {
          score += 5;
          hasBalance = true;
        } else if (REFERENCE_ALIASES.includes(str)) {
          score += 5;
          hasRef = true;
        } else if (DR_CR_ALIASES.includes(str)) {
          score += 5;
        }
      });

      // Special literal checks
      if (hasDateLike && hasNarrLike && (hasDebitCreditOrAmt || hasBalance)) {
        score += 50;
      }

      if (score > maxScore) {
        maxScore = score;
        bestHeaderRowIdx = r;
      }
    }

    const headerRow = rows[bestHeaderRowIdx] || [];
    const mappings = {
      date: null as number | null,
      narration: null as number | null,
      debit: null as number | null,
      credit: null as number | null,
      amount: null as number | null,
      drCr: null as number | null,
      balance: null as number | null,
      reference: null as number | null,
    };

    headerRow.forEach((val, c) => {
      if (val === undefined || val === null) return;
      const str = String(val).toLowerCase().trim();
      if (!str) return;

      if (DATE_ALIASES.includes(str) && mappings.date === null) {
        mappings.date = c;
      } else if (NARRATION_ALIASES.includes(str) && mappings.narration === null) {
        mappings.narration = c;
      } else if (DEBIT_ALIASES.includes(str) && mappings.debit === null) {
        mappings.debit = c;
      } else if (CREDIT_ALIASES.includes(str) && mappings.credit === null) {
        mappings.credit = c;
      } else if (AMOUNT_ALIASES.includes(str) && mappings.amount === null) {
        mappings.amount = c;
      } else if (DR_CR_ALIASES.includes(str) && mappings.drCr === null) {
        mappings.drCr = c;
      } else if (BALANCE_ALIASES.includes(str) && mappings.balance === null) {
        mappings.balance = c;
      } else if (REFERENCE_ALIASES.includes(str) && mappings.reference === null) {
        mappings.reference = c;
      }
    });

    headerRow.forEach((val, c) => {
      if (val === undefined || val === null) return;
      const str = String(val).toLowerCase().trim();
      if (!str) return;

      if (mappings.date === null && str.includes('date')) {
        mappings.date = c;
      } else if (mappings.narration === null && (str.includes('narr') || str.includes('desc') || str.includes('detail') || str.includes('remark'))) {
        mappings.narration = c;
      } else if (mappings.debit === null && (str.includes('debit') || str.includes('withdr') || str.includes('paid out'))) {
        mappings.debit = c;
      } else if (mappings.credit === null && (str.includes('credit') || str.includes('deposit') || str.includes('paid in'))) {
        mappings.credit = c;
      } else if (mappings.amount === null && str.includes('amount') && !str.includes('balance')) {
        mappings.amount = c;
      } else if (mappings.reference === null && (str.includes('ref') || str.includes('utr') || str.includes('cheque') || str.includes('chq') || str.includes('instrument') || str.includes('txn id') || str.includes('transaction id'))) {
        mappings.reference = c;
      } else if (mappings.drCr === null && (str.includes('dr/cr') || str.includes('d/c') || str.includes('type'))) {
        mappings.drCr = c;
      } else if (mappings.balance === null && str.includes('balance')) {
        mappings.balance = c;
      }
    });

    if (mappings.narration === null) {
      headerRow.forEach((val, c) => {
        if (val === undefined || val === null) return;
        const str = String(val).toLowerCase().trim();
        if (str.includes('particulars') || str.includes('part')) {
          mappings.narration = c;
        }
      });
    }

    let dataStartRowIdx = bestHeaderRowIdx + 1;
    while (dataStartRowIdx < rows.length) {
      const row = rows[dataStartRowIdx];
      const isRowBlank = !row || row.every(val => val === undefined || val === null || String(val).trim() === '');
      if (!isRowBlank) {
        break;
      }
      dataStartRowIdx++;
    }

    let dataEndRowIdx = rows.length - 1;
    const summaryKeywords = [
      'opening balance', 'closing balance', 'total', 'grand total', 'statement summary',
      'page total', 'carried forward', 'brought forward', 'opening bal', 'closing bal'
    ];

    for (let r = dataStartRowIdx; r < rows.length; r++) {
      const row = rows[r];
      if (!row) continue;

      let isSummaryRow = false;
      for (const val of row) {
        if (val === undefined || val === null) continue;
        const strVal = String(val).toLowerCase().trim();
        if (summaryKeywords.some(kw => strVal === kw || strVal.startsWith(kw) || strVal.includes(' ' + kw) || strVal.includes(kw + ' '))) {
          isSummaryRow = true;
          break;
        }
      }

      if (isSummaryRow) {
        dataEndRowIdx = r - 1;
        break;
      }
    }

    if (dataEndRowIdx < dataStartRowIdx) {
      dataEndRowIdx = rows.length - 1;
    }

    while (dataEndRowIdx >= dataStartRowIdx) {
      const row = rows[dataEndRowIdx];
      const isRowBlank = !row || row.every(val => val === undefined || val === null || String(val).trim() === '');
      if (!isRowBlank) {
        break;
      }
      dataEndRowIdx--;
    }

    return {
      headerRowIdx: bestHeaderRowIdx,
      dataStartRowIdx,
      dataEndRowIdx,
      mappings
    };
  };

  const parseBankStatementTextLocally = (text: string): any[] => {
    const lines = text.split('\n');
    const results: any[] = [];
    
    // Date pattern: DD/MM/YYYY or DD-MM-YYYY or DD/MM/YY or DD-MM-YY
    const dateRegex = /\b(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})\b/g;

    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed) continue;
      
      dateRegex.lastIndex = 0;
      const dateMatch = dateRegex.exec(trimmed);
      if (dateMatch) {
        const dateStr = dateMatch[0];
        const rest = trimmed.replace(dateStr, '').trim();
        
        const numberMatches = rest.match(/\b\d[\d,]*\.\d{2}\b|\b\d[\d,]*\b/g);
        
        let debitVal: number | null = null;
        let creditVal: number | null = null;
        let amountVal = 0;
        let description = rest;

        if (numberMatches && numberMatches.length > 0) {
          const numbers = numberMatches.map(n => parseFloat(n.replace(/,/g, ''))).filter(n => !isNaN(n));
          
          for (const numStr of numberMatches) {
            description = description.replace(numStr, '');
          }
          
          if (numbers.length >= 3) {
            debitVal = numbers[0];
            creditVal = numbers[1];
          } else if (numbers.length === 2) {
            const lowerRest = rest.toLowerCase();
            const isDebitKeyword = ['charge', 'tax', 'fee', 'withdrawal', 'paid', 'to', 'rtgs dr', 'neft dr', 'debit', 'chg'].some(k => lowerRest.includes(k));
            const isCreditKeyword = ['interest', 'deposit', 'received', 'from', 'rtgs cr', 'neft cr', 'credit', 'refund'].some(k => lowerRest.includes(k));
            
            if (isDebitKeyword) {
              debitVal = numbers[0];
              creditVal = null;
            } else if (isCreditKeyword) {
              debitVal = null;
              creditVal = numbers[0];
            } else {
              debitVal = numbers[0];
              creditVal = numbers[1];
            }
          } else if (numbers.length === 1) {
            const lowerRest = rest.toLowerCase();
            const isCreditKeyword = ['interest', 'deposit', 'received', 'from', 'rtgs cr', 'neft cr', 'credit', 'refund'].some(k => lowerRest.includes(k));
            if (isCreditKeyword) {
              creditVal = numbers[0];
            } else {
              debitVal = numbers[0];
            }
            amountVal = numbers[0];
          }
        }

        description = description.replace(/\s+/g, ' ').trim();
        if (!description) description = 'Transaction';

        results.push({
          date: dateStr,
          description: description,
          debit: debitVal,
          credit: creditVal,
          amount: amountVal
        });
      }
    }
    return results;
  };

  const parseBankStatementExcelOrCsv = (jsonData: any[], currentVoucherMode: 'auto' | 'payment' | 'receipt') => {
    if (jsonData.length === 0) return [];

    const headers = Array.from(new Set(jsonData.flatMap(row => Object.keys(row as object))));

    const getHeaderKey = (candidates: string[]) => {
      for (const candidate of candidates) {
        const key = headers.find(h => h.toLowerCase().trim() === candidate.toLowerCase().trim());
        if (key !== undefined) {
          return key;
        }
      }
      return null;
    };

    const DATE_HEADERS = ['date', 'txn date', 'transaction date', 'value date', 'posting date'];
    const NARRATION_HEADERS = ['narration', 'description', 'transaction details', 'details', 'remarks'];
    const DEBIT_HEADERS = ['debit', 'withdrawal', 'withdrawals', 'paid out', 'dr', 'debit amount'];
    const CREDIT_HEADERS = ['credit', 'deposit', 'deposits', 'paid in', 'cr', 'credit amount'];
    const AMOUNT_HEADERS = ['amount', 'transaction amount'];
    const REF_HEADERS = ['reference', 'ref no', 'utr', 'cheque no', 'instrument no', 'transaction id'];

    const dateCol = getHeaderKey(DATE_HEADERS);
    let descCol = getHeaderKey(NARRATION_HEADERS);
    if (!descCol) {
      descCol = headers.find(h => h.toLowerCase().trim() === 'particulars') || null;
    }
    const debitCol = getHeaderKey(DEBIT_HEADERS);
    const creditCol = getHeaderKey(CREDIT_HEADERS);
    const amtCol = getHeaderKey(AMOUNT_HEADERS);
    const refCol = getHeaderKey(REF_HEADERS);

    return jsonData.map((row: any, idx: number) => {
      let dateVal = dateCol ? row[dateCol] : undefined;
      let descVal = descCol ? row[descCol] : undefined;
      let debitVal = debitCol ? row[debitCol] : undefined;
      let creditVal = creditCol ? row[creditCol] : undefined;
      let amtVal = amtCol ? row[amtCol] : undefined;
      let refVal = refCol ? row[refCol] : undefined;

      if (dateVal === undefined) {
        const dateKey = headers.find(h => h.toLowerCase().includes('date'));
        if (dateKey) dateVal = row[dateKey];
      }
      if (descVal === undefined) {
        const descKey = headers.find(h => h.toLowerCase().includes('desc') || h.toLowerCase().includes('narr') || h.toLowerCase().includes('part'));
        if (descKey) descVal = row[descKey];
      }

      let dateStr = '';
      if (dateVal instanceof Date) {
        const d = dateVal.getDate();
        const m = dateVal.getMonth() + 1;
        const y = dateVal.getFullYear();
        dateStr = `${String(d).padStart(2, '0')}/${String(m).padStart(2, '0')}/${y}`;
      } else if (typeof dateVal === 'number' && dateVal > 40000 && dateVal < 100000) {
        const date = new Date((dateVal - 25569) * 86400 * 1000);
        const d = date.getDate();
        const m = date.getMonth() + 1;
        const y = date.getFullYear();
        dateStr = `${String(d).padStart(2, '0')}/${String(m).padStart(2, '0')}/${y}`;
      } else if (dateVal !== undefined && dateVal !== null) {
        dateStr = String(dateVal).trim();
      }

      let description = descVal !== undefined && descVal !== null ? String(descVal).trim() : 'No Narration Found';

      let debitNum: number | null = null;
      let creditNum: number | null = null;

      if (debitVal !== undefined && debitVal !== null && String(debitVal).trim() !== '') {
        const parsed = parseFloat(String(debitVal).replace(/,/g, ''));
        if (!isNaN(parsed) && parsed > 0) debitNum = parsed;
      }
      if (creditVal !== undefined && creditVal !== null && String(creditVal).trim() !== '') {
        const parsed = parseFloat(String(creditVal).replace(/,/g, ''));
        if (!isNaN(parsed) && parsed > 0) creditNum = parsed;
      }

      let amountNum = 0;
      if (amtVal !== undefined && amtVal !== null && String(amtVal).trim() !== '') {
        const parsed = parseFloat(String(amtVal).replace(/,/g, ''));
        if (!isNaN(parsed)) amountNum = parsed;
      }

      const reference = refVal !== undefined && refVal !== null ? String(refVal).trim() : '';

      let detectedVoucherType: 'Payment' | 'Receipt' | 'Unknown' = 'Unknown';
      let status: 'valid' | 'invalid' | 'warning' = 'valid';
      let errorMsg = '';
      let finalAmount = 0;

      const hasDebit = debitNum !== null && debitNum > 0;
      const hasCredit = creditNum !== null && creditNum > 0;

      if (hasDebit && hasCredit) {
        status = 'invalid';
        errorMsg = 'Both debit and credit found in same row. Please verify.';
      } else if (!hasDebit && !hasCredit && !amountNum) {
        status = 'invalid';
        errorMsg = 'Debit, Credit, and Amount are all blank.';
      } else if (hasDebit) {
        detectedVoucherType = 'Payment';
        finalAmount = debitNum!;
      } else if (hasCredit) {
        detectedVoucherType = 'Receipt';
        finalAmount = creditNum!;
      } else {
        if (selectedVoucherType === 'Payment') {
          detectedVoucherType = 'Payment';
          finalAmount = Math.abs(amountNum);
          status = 'warning';
          errorMsg = 'Single Amount column found. Voucher type has been applied using selected screen mode.';
        } else if (selectedVoucherType === 'Receipt') {
          detectedVoucherType = 'Receipt';
          finalAmount = Math.abs(amountNum);
          status = 'warning';
          errorMsg = 'Single Amount column found. Voucher type has been applied using selected screen mode.';
        } else {
          detectedVoucherType = amountNum < 0 ? 'Payment' : 'Receipt';
          finalAmount = Math.abs(amountNum);
          status = 'warning';
          errorMsg = 'Single Amount column found. Auto classified by amount sign.';
        }
      }

      let excluded = false;
      if (status === 'valid' || status === 'warning') {
        if (currentVoucherMode === 'payment') {
          if (detectedVoucherType === 'Receipt') {
            status = 'warning';
            errorMsg = 'Skipped: Credit row in Payment only mode';
            excluded = true;
          }
        } else if (currentVoucherMode === 'receipt') {
          if (detectedVoucherType === 'Payment') {
            status = 'warning';
            errorMsg = 'Skipped: Debit row in Receipt only mode';
            excluded = true;
          }
        }
      }

      const suggestedLedger = suggestLedgerForNarration(description, tallyContext);

      return {
        rowNo: idx + 1,
        rawRow: row,
        date: dateStr,
        rawDate: dateVal,
        description,
        debit: debitNum,
        credit: creditNum,
        amount: finalAmount,
        detectedVoucherType,
        suggestedLedger,
        userLedger: suggestedLedger,
        reference,
        status,
        errorMsg,
        excluded
      };
    });
  };

  const handleBankStatementGenerateXML = async () => {
    setIsProcessing(true);
    setError(null);
    setBankStatementValidationErrors([]);

    try {
      // 1. Get included rows
      const includedRows = bankStatementRows.filter(row => !row.excluded);
      if (includedRows.length === 0) {
        throw new Error("No non-excluded bank statement rows selected for XML generation.");
      }

      // 2. Perform validations
      const validationErrs: { rowNo: number; issue: string; suggestedFix: string }[] = [];

      // Requirement 8: Bank/Cash ledger not selected
      if (!selectedBankLedger) {
        validationErrs.push({
          rowNo: 0,
          issue: "Bank/Cash ledger not selected",
          suggestedFix: "Please select a Bank / Cash Ledger from the top of the mapping screen or current workspace."
        });
      }

      // Check each included row
      includedRows.forEach(row => {
        // Date validation
        const dateNorm = normalizeTallyDate(row.date);
        if (!dateNorm.isValid) {
          validationErrs.push({
            rowNo: row.rowNo,
            issue: `Invalid date: "${row.date}"`,
            suggestedFix: "Correct the date format to DD-MM-YYYY or DD/MM/YYYY."
          });
        }

        // Voucher type validation
        if (!row.detectedVoucherType || row.detectedVoucherType === 'Unknown') {
          validationErrs.push({
            rowNo: row.rowNo,
            issue: "Voucher Type missing or unknown",
            suggestedFix: "Ensure the transaction is classified as either Payment or Receipt."
          });
        }

        // Amount validation
        if (row.amount === null || row.amount === undefined || row.amount <= 0) {
          validationErrs.push({
            rowNo: row.rowNo,
            issue: "Amount missing or zero",
            suggestedFix: "Enter a valid positive number for debit or credit amount."
          });
        }

        // Final Ledger blank validation
        if (!row.userLedger || !row.userLedger.trim()) {
          validationErrs.push({
            rowNo: row.rowNo,
            issue: "Final Ledger (Particulars) is blank",
            suggestedFix: "Select or type a valid ledger for this transaction."
          });
        }

        // Suspense confirmation validation
        const isSuspense = row.userLedger && /suspense/i.test(row.userLedger);
        if (isSuspense && !confirmProceedWithSuspense) {
          validationErrs.push({
            rowNo: row.rowNo,
            issue: "Suspense row exists and is not confirmed",
            suggestedFix: "Confirm 'Proceed with Suspense' at the bottom of the screen or map to a different ledger."
          });
        }
      });

      // Also check for blocked invalid rows in general if they are included
      const invalidIncluded = includedRows.find(row => row.status === 'invalid');
      if (invalidIncluded) {
        validationErrs.push({
          rowNo: invalidIncluded.rowNo,
          issue: "Row has critical invalid state: " + (invalidIncluded.errorMsg || "unrecognized format"),
          suggestedFix: "Please exclude this row or fix the columns before exporting."
        });
      }

      if (validationErrs.length > 0) {
        setBankStatementValidationErrors(validationErrs);
        throw new Error("Validation failed. Please review and correct the row-wise error summary below.");
      }

      // 3. Build verificationRows
      const rows = includedRows.map(row => {
        const dateNorm = normalizeTallyDate(row.date);
        return {
          rowNo: row.rowNo,
          voucherType: row.detectedVoucherType,
          originalDate: row.date,
          normalizedDate: dateNorm.value,
          finalLedger: row.userLedger,
          bankLedger: selectedBankLedger || 'Bank Account',
          amount: row.amount,
          description: row.description || '',
          reference: row.reference || '',
          sourceIdx: bankStatementRows.indexOf(row),
        };
      });

      setVerificationRows(rows);
      setVerificationSourceStep('bank-statement-review');

      // 4. Pass the newly built rows directly to checkMissingMastersAndProceed to prevent React state cycle lag
      checkMissingMastersAndProceed('Vouchers', rows);
    } catch (err) {
      setError(err instanceof Error ? err.message : "XML generation failed");
      setIsProcessing(false);
    }
  };

  const applyDetectionAndMapLedgers = () => {
    setIsProcessing(true);
    setError(null);

    try {
      const results: any[] = [];
      
      const dCol = columnMappings.date;
      const nCol = columnMappings.narration;
      const dbCol = columnMappings.debit;
      const crCol = columnMappings.credit;
      const aCol = columnMappings.amount;
      const dcCol = columnMappings.drCr;
      const rCol = columnMappings.reference;
      
      // Check that we have at least Date and Narration columns mapped!
      if (dCol === null) throw new Error("Please select Date column mapping.");
      if (nCol === null) throw new Error("Please select Narration column mapping.");
      
      // Loop from dataStartRowIdx to dataEndRowIdx
      const startIdx = Math.max(0, dataStartRowIdx);
      const endIdx = Math.min(rawGrid.length - 1, dataEndRowIdx);
      
      // Check if Suspense Account is available
      const hasSuspense = tallyContext?.ledgers.includes('Suspense Account') || false;

      for (let r = startIdx; r <= endIdx; r++) {
        const row = rawGrid[r];
        if (!row) continue;
        
        // If the row is completely empty, skip it
        const isRowBlank = row.every(val => val === undefined || val === null || String(val).trim() === '');
        if (isRowBlank) continue;
        
        // Date value parsing
        let dateVal = row[dCol];
        let dateStr = '';
        if (dateVal instanceof Date) {
          const d = dateVal.getDate();
          const m = dateVal.getMonth() + 1;
          const y = dateVal.getFullYear();
          dateStr = `${String(d).padStart(2, '0')}/${String(m).padStart(2, '0')}/${y}`;
        } else if (typeof dateVal === 'number' && dateVal > 40000 && dateVal < 100000) {
          const date = new Date((dateVal - 25569) * 86400 * 1000);
          const d = date.getDate();
          const m = date.getMonth() + 1;
          const y = date.getFullYear();
          dateStr = `${String(d).padStart(2, '0')}/${String(m).padStart(2, '0')}/${y}`;
        } else if (dateVal !== undefined && dateVal !== null) {
          dateStr = String(dateVal).trim();
        }
        
        // Narration
        let descVal = row[nCol];
        let description = descVal !== undefined && descVal !== null ? String(descVal).trim() : 'No Narration Found';
        
        // Debit & Credit
        let debitVal = dbCol !== null ? row[dbCol] : undefined;
        let creditVal = crCol !== null ? row[crCol] : undefined;
        let debitNum: number | null = null;
        let creditNum: number | null = null;
        
        if (debitVal !== undefined && debitVal !== null && String(debitVal).trim() !== '') {
          const parsed = parseFloat(String(debitVal).replace(/,/g, ''));
          if (!isNaN(parsed) && parsed > 0) debitNum = parsed;
        }
        if (creditVal !== undefined && creditVal !== null && String(creditVal).trim() !== '') {
          const parsed = parseFloat(String(creditVal).replace(/,/g, ''));
          if (!isNaN(parsed) && parsed > 0) creditNum = parsed;
        }
        
        // Single Amount
        let amountNum = 0;
        if (aCol !== null && row[aCol] !== undefined && row[aCol] !== null && String(row[aCol]).trim() !== '') {
          const parsed = parseFloat(String(row[aCol]).replace(/,/g, ''));
          if (!isNaN(parsed)) amountNum = parsed;
        }
        
        // Dr/Cr Indicator
        let dcVal = dcCol !== null ? row[dcCol] : undefined;
        let drCrStr = '';
        if (dcVal !== undefined && dcVal !== null) {
          const str = String(dcVal).toLowerCase().trim();
          if (['dr', 'debit'].includes(str)) {
            drCrStr = 'Dr';
          } else if (['cr', 'credit'].includes(str)) {
            drCrStr = 'Cr';
          } else if (str) {
            drCrStr = str;
          }
        }
        
        // Reference
        let refVal = rCol !== null ? row[rCol] : undefined;
        let reference = refVal !== undefined && refVal !== null ? String(refVal).trim() : '';
        
        // Classify transaction row-wise
        let detectedVoucherType: 'Payment' | 'Receipt' | 'Unknown' = 'Unknown';
        let status: 'valid' | 'invalid' | 'warning' = 'valid';
        let errorMsg = '';
        let finalAmount = 0;
        
        // If separate debit/credit columns exist:
        if (dbCol !== null || crCol !== null) {
          const hasDebit = debitNum !== null && debitNum > 0;
          const hasCredit = creditNum !== null && creditNum > 0;
          
          if (hasDebit && hasCredit) {
            status = 'invalid';
            errorMsg = 'Both debit and credit found in same row. Please verify.';
          } else if (!hasDebit && !hasCredit) {
            status = 'invalid';
            errorMsg = 'No debit/credit amount found.';
          } else if (hasDebit) {
            detectedVoucherType = 'Payment';
            finalAmount = debitNum!;
          } else if (hasCredit) {
            detectedVoucherType = 'Receipt';
            finalAmount = creditNum!;
          }
        } 
        // If single amount + Dr/Cr indicator exists:
        else if (aCol !== null && dcCol !== null) {
          if (drCrStr === 'Dr') {
            detectedVoucherType = 'Payment';
            finalAmount = Math.abs(amountNum);
          } else if (drCrStr === 'Cr') {
            detectedVoucherType = 'Receipt';
            finalAmount = Math.abs(amountNum);
          } else {
            status = 'invalid';
            errorMsg = 'Unable to identify Dr/Cr indicator.';
          }
        } 
        // If only single amount exists:
        else if (aCol !== null) {
          if (detectedVoucherMode === 'payment') {
            detectedVoucherType = 'Payment';
            finalAmount = Math.abs(amountNum);
          } else if (detectedVoucherMode === 'receipt') {
            detectedVoucherType = 'Receipt';
            finalAmount = Math.abs(amountNum);
          } else {
            status = 'invalid';
            errorMsg = 'Single amount column found without Dr/Cr indicator. Please select Payment Only or Receipt Only mode.';
          }
        } else {
          status = 'invalid';
          errorMsg = 'No amount or debit/credit column mapped.';
        }
        
        // Skip logic based on screen mode
        let excluded = false;
        if (status === 'valid') {
          if (voucherMode === 'payment' && detectedVoucherType === 'Receipt') {
            status = 'warning';
            errorMsg = 'Skipped: Credit row in Payment only mode';
            excluded = true;
          } else if (voucherMode === 'receipt' && detectedVoucherType === 'Payment') {
            status = 'warning';
            errorMsg = 'Skipped: Debit row in Receipt only mode';
            excluded = true;
          }
        }
        
        // Get enhanced ledger suggestion
        const suggestion = suggestLedgerForNarrationEnhanced(description, tallyContext);
        
        let rowStatus = status;
        let rowErrorMsg = errorMsg;
        if (rowStatus === 'valid' && !suggestion.ledgerName) {
          rowStatus = 'warning';
          rowErrorMsg = 'Needs Review';
        }

        results.push({
          rowNo: r + 1, // Excel row number (1-indexed)
          date: dateStr,
          description,
          debit: debitNum,
          credit: creditNum,
          amount: finalAmount,
          drCr: drCrStr,
          detectedVoucherType,
          suggestedLedger: suggestion.ledgerName,
          confidence: Math.round(suggestion.confidence * 100),
          userLedger: suggestion.ledgerName,
          reference,
          status: rowStatus,
          errorMsg: rowErrorMsg,
          excluded,
          reasoning: suggestion.reasoning
        });
      }
      
      setBankStatementRows(results);
      setCurrentStep('bank-statement-review');
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to apply detection mappings.");
    } finally {
      setIsProcessing(false);
    }
  };

  const downloadReviewedTemplateExcel = async () => {
    try {
      const validRows = bankStatementRows.filter(row => !row.excluded && row.status !== 'invalid');
      if (validRows.length === 0) {
        setError("No valid, non-excluded rows available to download.");
        return;
      }

      const wb = new ExcelJS.Workbook();
      const sheet = wb.addWorksheet('Reviewed_Vouchers');

      // Define columns/headers
      const headers = ['Date', 'Particulars', 'Voucher Type', 'Voucher No', 'Amount', 'Narration', 'Reference'];
      sheet.getRow(1).values = headers;
      sheet.getRow(1).font = { bold: true };

      // Set columns for formatting
      sheet.columns = [
        { header: 'Date', key: 'Date', width: 12 },
        { header: 'Particulars', key: 'Particulars', width: 25 },
        { header: 'Voucher Type', key: 'VoucherType', width: 15 },
        { header: 'Voucher No', key: 'VoucherNo', width: 12 },
        { header: 'Amount', key: 'Amount', width: 12 },
        { header: 'Narration', key: 'Narration', width: 40 },
        { header: 'Reference', key: 'Reference', width: 15 }
      ];

      // Add rows
      validRows.forEach(row => {
        sheet.addRow({
          Date: row.date,
          Particulars: row.userLedger || '',
          VoucherType: row.detectedVoucherType || 'Payment',
          VoucherNo: '',
          Amount: row.amount,
          Narration: row.description || '',
          Reference: row.reference || ''
        });
      });

      // Handle dropdown validation using Tally Masters
      const masterLedgers = tallyContext?.ledgers || [];
      if (masterLedgers.length > 0) {
        // Create hidden master sheet
        const masterSheet = wb.addWorksheet('Tally_Masters', { state: 'hidden' });
        masterLedgers.forEach((ledger, i) => {
          masterSheet.getCell(i + 1, 1).value = ledger;
        });

        // Add validation to the Particulars column (column B) for rows 2 to count+1
        const rowCount = validRows.length;
        const excelFormula = `Tally_Masters!$A$1:$A$${masterLedgers.length}`;
        for (let i = 2; i <= rowCount + 1; i++) {
          const cell = sheet.getCell(i, 2); // Column B is Particulars
          cell.dataValidation = {
            type: 'list',
            allowBlank: true,
            formulae: [excelFormula],
            showErrorMessage: true,
            errorTitle: 'Invalid Ledger',
            error: 'Please select a valid ledger from Tally Masters.'
          };
        }
      }

      // Download file
      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = `Reviewed_${pendingFileName || 'BankStatement'}.xlsx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      setBankSuccessMessage("Reviewed Excel file downloaded successfully!");
      setTimeout(() => setBankSuccessMessage(null), 5000);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to download reviewed template excel");
    }
  };

  // --- File Handling ---
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file || !user) return;

    if (file.type === 'application/pdf' && getAppMode() === 'desktop-offline') {
      setError("PDF parsing is not available offline yet. Please upload Excel/CSV bank statement.");
      return;
    }

    if (importType !== 'Voucher') {
      await handleMasterFileUpload(file);
      return;
    }

    if (voucherImportMethod === 'bankStatement') {
      if (!selectedBankLedger) {
        setError("Please select Bank / Cash Ledger before uploading bank statement.");
        return;
      }
    }

    setIsProcessing(true);
    setError(null);
    setPendingFileName(file.name);

    try {
      if (voucherImportMethod === 'bankStatement') {
        if (file.type === 'application/pdf') {
          const text = await extractTextFromPdf(file);
          const rawRows = parseBankStatementTextLocally(text);
          if (rawRows.length === 0) throw new Error("No transactions could be parsed from PDF.");

          const headers = ['Date', 'Narration', 'Debit', 'Credit', 'Amount'];
          const grid = [
            headers,
            ...rawRows.map(r => [r.date, r.description, r.debit, r.credit, r.amount])
          ];

          setRawGrid(grid);
          setHeaderRowIdx(0);
          setDataStartRowIdx(1);
          setDataEndRowIdx(grid.length - 1);
          setColumnMappings({
            date: 0,
            narration: 1,
            debit: 2,
            credit: 3,
            amount: 4,
            drCr: null,
            balance: null,
            reference: null
          });

          setCurrentStep('bank-statement-detection-review');
          setIsProcessing(false);
        } else {
          const reader = new FileReader();
          reader.onload = async (event) => {
            try {
              const data = new Uint8Array(event.target?.result as ArrayBuffer);
              const workbook = XLSX.read(data, { type: 'array', cellDates: true });
              const firstSheetName = workbook.SheetNames[0];
              const worksheet = workbook.Sheets[firstSheetName];
              
              // Get raw grid as a 2D array!
              const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
              if (rawRows.length === 0) throw new Error("Excel/CSV file is empty");

              // Save to state
              setRawGrid(rawRows);

              // Auto-detect structure!
              const detected = autoDetectBankStatement(rawRows);
              setHeaderRowIdx(detected.headerRowIdx);
              setDataStartRowIdx(detected.dataStartRowIdx);
              setDataEndRowIdx(detected.dataEndRowIdx);
              setColumnMappings(detected.mappings);

              // Go to detection review!
              setCurrentStep('bank-statement-detection-review');
              setIsProcessing(false);
            } catch (err) {
              setError(err instanceof Error ? err.message : "Failed to parse bank statement file");
              setIsProcessing(false);
            }
          };
          reader.readAsArrayBuffer(file);
        }
        return;
      }

      if (file.type === 'application/pdf') {
        const text = await extractTextFromPdf(file);
        let transactions: BankTransaction[] = [];
        if (isGeminiAvailable()) {
          transactions = await parseBankStatementText(text);
        } else {
          const rawRows = parseBankStatementTextLocally(text);
          transactions = rawRows.map(row => ({
            date: row.date,
            description: row.description,
            amount: row.amount || 0,
            reference: row.reference || ''
          }));
        }
        
        setPendingData(transactions);
        
        if (tallyContext) {
          const mapped = await mapBankTransactions(transactions, tallyContext.ledgers, tallyContext.historicalMappings);
          setAiMappedTransactions(mapped);
          
          setMappings([
            { excelColumn: 'date', tallyField: 'DATE', confidence: 1, reasoning: 'Extracted from PDF' },
            { excelColumn: 'description', tallyField: 'NARRATION', confidence: 1, reasoning: 'Extracted from PDF' },
            { excelColumn: 'amount', tallyField: 'AMOUNT', confidence: 1, reasoning: 'Extracted from PDF' }
          ]);
        } else {
          setMappings([
            { excelColumn: 'date', tallyField: 'DATE', confidence: 1, reasoning: 'Extracted from PDF' },
            { excelColumn: 'description', tallyField: 'PARTYNAME', confidence: 1, reasoning: 'Extracted from PDF' },
            { excelColumn: 'amount', tallyField: 'AMOUNT', confidence: 1, reasoning: 'Extracted from PDF' }
          ]);
        }
        
        setCurrentStep('mapping');
        setIsProcessing(false);
      } else {
        const reader = new FileReader();
        reader.onload = async (event) => {
          try {
            const data = new Uint8Array(event.target?.result as ArrayBuffer);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            
            if (jsonData.length === 0) throw new Error("Excel file is empty");

            // Get all unique keys from all rows to ensure we don't miss any headers
            const headers = Array.from(new Set(jsonData.flatMap(row => Object.keys(row as object))));

            const isSalesPurchase = headers.some(h => {
              const hLower = h.toLowerCase().trim();
              return hLower.includes('stock item') || hLower.includes('invoice no') || hLower.includes('party ledger') || hLower.includes('sales/purchase ledger');
            });

            if (isSalesPurchase) {
              const invoices = processSalesPurchaseExcel(jsonData);
              invoices.forEach(inv => recalculateInvoice(inv, companyState));
              setSalesPurchaseInvoices(invoices);
              setSalesPurchaseBalancingErrors([]);
              setCurrentStep('sales-purchase-verification');
              setIsProcessing(false);
              return;
            }

            const isJournal = headers.some(h => {
              const hLower = h.toLowerCase().trim();
              return hLower === 'dr/cr' || hLower === 'dr_cr' || hLower === 'dr_or_cr' || hLower === 'dr/cr selection' || hLower.includes('dr/cr');
            });

            if (isJournal) {
              const groups = parseJournalExcel(jsonData);
              setJournalGroups(groups);
              setSelectedJournalGroupIdx(0);
              
              const hasErrors = groups.some(g => !g.isValid);
              if (hasErrors) {
                generateJournalErrorExcel(groups);
              }
              
              setCurrentStep('journal-verification');
              setIsProcessing(false);
              return;
            }

            setPendingData(jsonData);

            // Get AI Column Mapping FIRST
            const aiMappings = await getAIColumnMapping(headers);
            setMappings(aiMappings);

            // Helper for case-insensitive and trimmed row access
            const getValueByHeader = (row: any, header?: string) => {
              if (!header) return null;
              if (row[header] !== undefined) return row[header];
              const key = Object.keys(row).find(k => k.toLowerCase().trim() === header.toLowerCase().trim());
              return key ? row[key] : null;
            };

            // Use AI mappings to extract bank transactions correctly
            const dateCol = aiMappings.find(m => m.tallyField === 'DATE')?.excelColumn;
            let descCol = aiMappings.find(m => m.tallyField === 'NARRATION')?.excelColumn;
            const amtCol = aiMappings.find(m => m.tallyField === 'AMOUNT')?.excelColumn;

            // Extract the PARTYNAME / Particulars column separately from NARRATION
            let partyCol = aiMappings.find(m => m.tallyField === 'PARTYNAME')?.excelColumn;
            if (!partyCol) {
              const partyAliases = ['party name', 'partyname', 'particulars', 'ledger', 'ledger name', 'account'];
              partyCol = headers.find(h => 
                partyAliases.some(alias => h.toLowerCase().trim() === alias)
              ) || headers.find(h => 
                partyAliases.some(alias => h.toLowerCase().trim().includes(alias))
              );
            }

            // If NARRATION is not found or is the same as DATE or PARTYNAME, find/fallback NARRATION
            if (!descCol || descCol === dateCol || descCol === partyCol) {
              descCol = headers.find(h => 
                h !== dateCol && h !== partyCol &&
                !['date', 'amount', 'balance', 'vch', 'no'].some(k => h.toLowerCase().includes(k)) &&
                ['particulars', 'description', 'narration', 'remarks', 'details'].some(k => h.toLowerCase().includes(k))
              ) || descCol || headers.find(h => 
                ['particulars', 'description', 'narration', 'remarks', 'details'].some(k => h.toLowerCase().includes(k))
              );
            }

            const bankTransactions: BankTransaction[] = jsonData.map((row: any) => {
              // Handle Excel date serial numbers if necessary
              let dateVal = getValueByHeader(row, dateCol) || getValueByHeader(row, 'Date') || getValueByHeader(row, 'date') || '';
              
              if (dateVal instanceof Date) {
                dateVal = dateVal.toLocaleDateString('en-GB');
              } else if (typeof dateVal === 'number' && dateVal > 40000) {
                // Basic conversion for Excel serial dates
                const date = new Date((dateVal - 25569) * 86400 * 1000);
                dateVal = date.toLocaleDateString('en-GB');
              }

              const rawAmount = getValueByHeader(row, amtCol) || getValueByHeader(row, 'Amount') || getValueByHeader(row, 'amount') || getValueByHeader(row, 'Credit') || getValueByHeader(row, 'Debit') || 0;
              const parsedAmount = parseFloat(String(rawAmount).replace(/,/g, ''));

              // Try to find narration with multiple fallbacks
              let narration = getValueByHeader(row, descCol);
              if (!narration || String(narration).trim() === '') {
                const commonNarrationHeaders = ['narration', 'particulars', 'description', 'remarks', 'details', 'txn details', 'transaction details'];
                for (const h of commonNarrationHeaders) {
                  if (h === partyCol) continue; // Prefer not using the party column for narration
                  const val = getValueByHeader(row, h);
                  if (val && String(val).trim() !== '') {
                    narration = val;
                    break;
                  }
                }
              }

              // Extract particulars value
              let particularsStr = '';
              if (partyCol) {
                const rawParticulars = getValueByHeader(row, partyCol);
                particularsStr = rawParticulars !== undefined && rawParticulars !== null ? String(rawParticulars).trim() : '';
              }

              const refCol = aiMappings.find(m => m.tallyField === 'REFERENCE')?.excelColumn;
              let referenceStr = '';
              if (refCol) {
                const rawRef = getValueByHeader(row, refCol);
                referenceStr = rawRef !== undefined && rawRef !== null ? String(rawRef).trim() : '';
              } else {
                const commonRefHeaders = ['reference', 'ref no', 'utr', 'cheque no', 'instrument no', 'transaction id', 'ref', 'chq no'];
                for (const h of commonRefHeaders) {
                  const val = getValueByHeader(row, h);
                  if (val !== undefined && val !== null && String(val).trim() !== '') {
                    referenceStr = String(val).trim();
                    break;
                  }
                }
              }

              return {
                date: String(dateVal),
                description: String(narration || 'No Narration Found'),
                amount: isNaN(parsedAmount) ? 0 : parsedAmount,
                particulars: particularsStr,
                reference: referenceStr
              };
            });

            if (tallyContext) {
              const mapped = await mapBankTransactions(bankTransactions, tallyContext.ledgers, tallyContext.historicalMappings);
              setAiMappedTransactions(mapped);
            } else {
              setAiMappedTransactions(bankTransactions.map(tx => ({
                ...tx,
                tallyLedger: tx.particulars || '',
                confidence: tx.particulars ? 1.0 : 0,
                reasoning: tx.particulars ? 'Preserved user-provided ledger.' : 'No Tally masters uploaded.'
              })));
            }
            
            setCurrentStep('mapping');
            setIsProcessing(false);
          } catch (err) {
            setError(err instanceof Error ? err.message : "Failed to parse Excel file");
            setIsProcessing(false);
          }
        };
        reader.readAsArrayBuffer(file);
      }
    } catch (err) {
      setError("Failed to read file");
      setIsProcessing(false);
    }
  };

  // --- Master Upload and Handling ---
  const handleMasterFileUpload = async (file: File) => {
    setIsProcessing(true);
    setError(null);
    setPendingFileName(file.name);

    try {
      const reader = new FileReader();
      reader.onload = async (event) => {
        try {
          const data = new Uint8Array(event.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet);

          if (jsonData.length === 0) throw new Error("Excel file is empty");

          if (importType === 'Ledger') {
            const rows: LedgerMasterRow[] = jsonData.map((row: any, index) => {
              const fields = mapRowToFields(row, LEDGER_FIELDS);
              const ledgerName = String(fields.ledgerName || '').trim();
              const underGroup = String(fields.underGroup || '').trim();
              
              const errors: string[] = [];
              const warnings: string[] = [];

              if (!ledgerName) errors.push("Ledger Name cannot be blank.");
              if (!underGroup) errors.push("Under Group cannot be blank.");

              // Country defaults to INDIA
              const country = String(fields.country || '').trim().toUpperCase() || 'INDIA';

              // GSTIN validation and State auto-detection
              let state = String(fields.state || '').trim();
              const gstin = String(fields.gstin || '').trim().toUpperCase();
              if (gstin) {
                const gstinValidation = validateGSTIN(gstin);
                if (!gstinValidation.isValid) {
                  errors.push(gstinValidation.error || "Invalid GSTIN.");
                } else if (gstinValidation.stateName) {
                  state = gstinValidation.stateName;
                }
              }

              // Group verification
              if (tallyContext) {
                const groupExists = tallyContext.groups.some(g => g.toLowerCase() === underGroup.toLowerCase());
                if (!groupExists) {
                  warnings.push(`Group "${underGroup}" not found in uploaded Tally masters. Please confirm whether to create/use this group.`);
                }
              } else {
                warnings.push("Tally Masters not loaded. Group validation is not available.");
              }

              // Duplicate checking
              let isDuplicate = false;
              let isPossibleDuplicate = false;
              let duplicateMessage = "";

              if (tallyContext && ledgerName) {
                const exactMatch = tallyContext.ledgers.some(l => l.toLowerCase() === ledgerName.toLowerCase());
                if (exactMatch) {
                  isDuplicate = true;
                  duplicateMessage = `Ledger "${ledgerName}" already exists in Tally context.`;
                } else {
                  // Fuzzy similarity checking
                  const closeMatch = tallyContext.ledgers.find(l => getSimilarity(l, ledgerName) >= 0.8);
                  if (closeMatch) {
                    isPossibleDuplicate = true;
                    duplicateMessage = `Similar ledger "${closeMatch}" exists in Tally context.`;
                  }
                }
              }

              return {
                rowNum: index + 1,
                ledgerName,
                underGroup,
                openingBalance: fields.openingBalance ? String(fields.openingBalance) : '',
                drCr: fields.drCr ? String(fields.drCr) : 'Dr',
                mailingName: fields.mailingName ? String(fields.mailingName) : ledgerName,
                address1: fields.address1 ? String(fields.address1) : '',
                address2: fields.address2 ? String(fields.address2) : '',
                state,
                country,
                pincode: fields.pincode ? String(fields.pincode) : '',
                pan: fields.pan ? String(fields.pan) : '',
                gstin,
                registrationType: fields.registrationType ? String(fields.registrationType) : 'Regular',
                taxability: fields.taxability ? String(fields.taxability) : 'Taxable',
                isBillwiseOn: fields.isBillwiseOn ? String(fields.isBillwiseOn) : 'No',
                isCostCentreOn: fields.isCostCentreOn ? String(fields.isCostCentreOn) : 'No',
                email: fields.email ? String(fields.email) : '',
                mobileNumber: fields.mobileNumber ? String(fields.mobileNumber) : '',
                
                isValid: errors.length === 0,
                errors,
                warnings,
                isDuplicate,
                isPossibleDuplicate,
                duplicateMessage,
                excluded: false
              };
            });
            setParsedLedgers(rows);
            setCurrentStep('master-review');
          } else if (importType === 'StockItem') {
            const rows: StockMasterRow[] = jsonData.map((row: any, index) => {
              const fields = mapRowToFields(row, STOCK_FIELDS);
              const itemName = String(fields.itemName || '').trim();
              const underGroup = String(fields.underGroup || '').trim();
              const unit = String(fields.unit || '').trim();

              const errors: string[] = [];
              const warnings: string[] = [];

              if (!itemName) errors.push("Stock Item Name cannot be blank.");
              if (!underGroup) errors.push("Under Stock Group cannot be blank.");
              if (!unit) errors.push("Unit cannot be blank.");

              // Validate numeric fields
              if (fields.gstRate && isNaN(Number(fields.gstRate))) {
                errors.push("GST Rate must be numeric.");
              }
              if (fields.openingQty && isNaN(Number(fields.openingQty))) {
                errors.push("Opening Quantity must be numeric.");
              }
              if (fields.openingRate && isNaN(Number(fields.openingRate))) {
                errors.push("Opening Rate must be numeric.");
              }
              if (fields.openingValue && isNaN(Number(fields.openingValue))) {
                errors.push("Opening Value must be numeric.");
              }

              // Verify stock group and unit
              if (tallyContext) {
                if (tallyContext.stockGroups && tallyContext.stockGroups.length > 0) {
                  const sgExists = tallyContext.stockGroups.some(g => g.toLowerCase() === underGroup.toLowerCase());
                  if (!sgExists) {
                    warnings.push(`Stock Group "${underGroup}" not found in Tally masters.`);
                  }
                }
                if (tallyContext.units && tallyContext.units.length > 0) {
                  const uExists = tallyContext.units.some(u => u.toLowerCase() === unit.toLowerCase());
                  if (!uExists) {
                    warnings.push(`Unit "${unit}" not found in Tally masters.`);
                  }
                }
              }

              // Duplicate checking
              let isDuplicate = false;
              let isPossibleDuplicate = false;
              let duplicateMessage = "";

              if (tallyContext && tallyContext.stockItems && itemName) {
                const exactMatch = tallyContext.stockItems.some(i => i.toLowerCase() === itemName.toLowerCase());
                if (exactMatch) {
                  isDuplicate = true;
                  duplicateMessage = `Stock Item "${itemName}" already exists in Tally context.`;
                } else {
                  const closeMatch = tallyContext.stockItems.find(i => getSimilarity(i, itemName) >= 0.8);
                  if (closeMatch) {
                    isPossibleDuplicate = true;
                    duplicateMessage = `Similar Stock Item "${closeMatch}" exists in Tally context.`;
                  }
                }
              }

              return {
                rowNum: index + 1,
                itemName,
                underGroup,
                unit,
                openingQty: fields.openingQty ? String(fields.openingQty) : '',
                openingRate: fields.openingRate ? String(fields.openingRate) : '',
                openingValue: fields.openingValue ? String(fields.openingValue) : '',
                hsn: fields.hsn ? String(fields.hsn) : '',
                gstApplicable: fields.gstApplicable ? String(fields.gstApplicable) : 'Applicable',
                taxability: fields.taxability ? String(fields.taxability) : 'Taxable',
                gstRate: fields.gstRate ? String(fields.gstRate) : '',
                cgstRate: fields.cgstRate ? String(fields.cgstRate) : '',
                sgstRate: fields.sgstRate ? String(fields.sgstRate) : '',
                igstRate: fields.igstRate ? String(fields.igstRate) : '',
                description: fields.description ? String(fields.description) : '',
                
                isValid: errors.length === 0,
                errors,
                warnings,
                isDuplicate,
                isPossibleDuplicate,
                duplicateMessage,
                excluded: false
              };
            });
            setParsedStockItems(rows);
            setCurrentStep('master-review');
          } else if (importType === 'StockGroup') {
            const rows: StockGroupRow[] = jsonData.map((row: any, index) => {
              const groupName = String(row['Stock Group Name'] || row['Group Name'] || row['Name'] || '').trim();
              const underGroup = String(row['Under Parent Group'] || row['Parent'] || 'Primary').trim();
              
              const errors: string[] = [];
              if (!groupName) errors.push("Stock Group Name cannot be blank.");

              return {
                rowNum: index + 1,
                groupName,
                underGroup,
                isValid: errors.length === 0,
                errors,
                excluded: false
              };
            });
            setParsedStockGroups(rows);
            setCurrentStep('master-review');
          } else if (importType === 'Unit') {
            const rows: UnitRow[] = jsonData.map((row: any, index) => {
              const symbol = String(row['Symbol'] || row['Unit'] || '').trim();
              const formalName = String(row['Formal Name'] || '').trim();
              const uqc = String(row['Unit Quantity Code (UQC)'] || row['UQC'] || '').trim();
              const decimalPlaces = String(row['Number of Decimal Places'] || '0').trim();

              const errors: string[] = [];
              if (!symbol) errors.push("Symbol cannot be blank.");

              return {
                rowNum: index + 1,
                symbol,
                formalName,
                uqc,
                decimalPlaces,
                isValid: errors.length === 0,
                errors,
                excluded: false
              };
            });
            setParsedUnits(rows);
            setCurrentStep('master-review');
          }

          setIsProcessing(false);
        } catch (err) {
          setError(err instanceof Error ? err.message : "Failed to parse Excel file");
          setIsProcessing(false);
        }
      };
      reader.readAsArrayBuffer(file);
    } catch (err) {
      setError("Failed to read file");
      setIsProcessing(false);
    }
  };

  const handleLedgerCellChange = (rowIndex: number, field: keyof LedgerMasterRow, value: any) => {
    const updated = [...parsedLedgers];
    const row = updated[rowIndex];
    
    (row as any)[field] = value;

    const errors: string[] = [];
    const warnings: string[] = [];

    if (!row.ledgerName.trim()) errors.push("Ledger Name cannot be blank.");
    if (!row.underGroup.trim()) errors.push("Under Group cannot be blank.");

    if (row.gstin) {
      const gstinValidation = validateGSTIN(row.gstin);
      if (!gstinValidation.isValid) {
        errors.push(gstinValidation.error || "Invalid GSTIN.");
      } else if (gstinValidation.stateName) {
        row.state = gstinValidation.stateName;
      }
    }

    if (tallyContext) {
      const groupExists = tallyContext.groups.some(g => g.toLowerCase() === row.underGroup.toLowerCase());
      if (!groupExists) {
        warnings.push(`Group "${row.underGroup}" not found in uploaded Tally masters. Please confirm whether to create/use this group.`);
      }
    } else {
      warnings.push("Tally Masters not loaded. Group validation is not available.");
    }

    let isDuplicate = false;
    let isPossibleDuplicate = false;
    let duplicateMessage = "";

    if (tallyContext && row.ledgerName) {
      const exactMatch = tallyContext.ledgers.some(l => l.toLowerCase() === row.ledgerName.toLowerCase());
      if (exactMatch) {
        isDuplicate = true;
        duplicateMessage = `Ledger "${row.ledgerName}" already exists in Tally context.`;
      } else {
        const closeMatch = tallyContext.ledgers.find(l => getSimilarity(l, row.ledgerName) >= 0.8);
        if (closeMatch) {
          isPossibleDuplicate = true;
          duplicateMessage = `Similar ledger "${closeMatch}" exists in Tally context.`;
        }
      }
    }

    row.isValid = errors.length === 0;
    row.errors = errors;
    row.warnings = warnings;
    row.isDuplicate = isDuplicate;
    row.isPossibleDuplicate = isPossibleDuplicate;
    row.duplicateMessage = duplicateMessage;

    setParsedLedgers(updated);
  };

  const handleStockItemCellChange = (rowIndex: number, field: keyof StockMasterRow, value: any) => {
    const updated = [...parsedStockItems];
    const row = updated[rowIndex];

    (row as any)[field] = value;

    const errors: string[] = [];
    const warnings: string[] = [];

    if (!row.itemName.trim()) errors.push("Stock Item Name cannot be blank.");
    if (!row.underGroup.trim()) errors.push("Under Stock Group cannot be blank.");
    if (!row.unit.trim()) errors.push("Unit cannot be blank.");

    if (row.gstRate && isNaN(Number(row.gstRate))) {
      errors.push("GST Rate must be numeric.");
    }
    if (row.openingQty && isNaN(Number(row.openingQty))) {
      errors.push("Opening Quantity must be numeric.");
    }
    if (row.openingRate && isNaN(Number(row.openingRate))) {
      errors.push("Opening Rate must be numeric.");
    }
    if (row.openingValue && isNaN(Number(row.openingValue))) {
      errors.push("Opening Value must be numeric.");
    }

    if (tallyContext) {
      if (tallyContext.stockGroups && tallyContext.stockGroups.length > 0) {
        const sgExists = tallyContext.stockGroups.some(g => g.toLowerCase() === row.underGroup.toLowerCase());
        if (!sgExists) {
          warnings.push(`Stock Group "${row.underGroup}" not found in Tally masters.`);
        }
      }
      if (tallyContext.units && tallyContext.units.length > 0) {
        const uExists = tallyContext.units.some(u => u.toLowerCase() === row.unit.toLowerCase());
        if (!uExists) {
          warnings.push(`Unit "${row.unit}" not found in Tally masters.`);
        }
      }
    }

    let isDuplicate = false;
    let isPossibleDuplicate = false;
    let duplicateMessage = "";

    if (tallyContext && tallyContext.stockItems && row.itemName) {
      const exactMatch = tallyContext.stockItems.some(i => i.toLowerCase() === row.itemName.toLowerCase());
      if (exactMatch) {
        isDuplicate = true;
        duplicateMessage = `Stock Item "${row.itemName}" already exists in Tally context.`;
      } else {
        const closeMatch = tallyContext.stockItems.find(i => getSimilarity(i, row.itemName) >= 0.8);
        if (closeMatch) {
          isPossibleDuplicate = true;
          duplicateMessage = `Similar Stock Item "${closeMatch}" exists in Tally context.`;
        }
      }
    }

    row.isValid = errors.length === 0;
    row.errors = errors;
    row.warnings = warnings;
    row.isDuplicate = isDuplicate;
    row.isPossibleDuplicate = isPossibleDuplicate;
    row.duplicateMessage = duplicateMessage;

    setParsedStockItems(updated);
  };

  const getGroupOptions = () => {
    const fallback = [
      'Sundry Debtors', 'Sundry Creditors', 'Bank Accounts', 'Cash-in-hand', 
      'Sales Accounts', 'Purchase Accounts', 'Direct Expenses', 'Indirect Expenses', 
      'Direct Incomes', 'Indirect Incomes', 'Duties & Taxes', 'Loans (Liability)', 
      'Current Assets', 'Current Liabilities', 'Capital Account'
    ];
    if (tallyContext && tallyContext.groups && tallyContext.groups.length > 0) {
      return Array.from(new Set([...tallyContext.groups, ...fallback])).sort();
    }
    return fallback.sort();
  };

  const getStockGroupOptions = () => {
    const fallback = ['Primary'];
    if (tallyContext && tallyContext.stockGroups && tallyContext.stockGroups.length > 0) {
      return Array.from(new Set([...tallyContext.stockGroups, ...fallback])).sort();
    }
    return fallback;
  };

  const getUnitOptions = () => {
    const fallback = ['NOS', 'PCS', 'KGS', 'BOX', 'LTR', 'MTR', 'BAG', 'BOX', 'CTN', 'DOZ', 'GMS', 'HRS', 'KGS', 'KLR', 'LTR', 'MTR', 'MTS', 'NOS', 'PAC', 'PCS', 'QTL', 'ROL', 'SET', 'SQF', 'SQM', 'TBS', 'TGM', 'THD', 'TON', 'UNT', 'YDS'];
    if (tallyContext && tallyContext.units && tallyContext.units.length > 0) {
      return Array.from(new Set([...tallyContext.units, ...fallback])).sort();
    }
    return fallback.sort();
  };

  const downloadValidationReport = () => {
    let reportData: any[] = [];
    let reportName = 'validation-report.xlsx';

    if (importType === 'Ledger') {
      reportData = parsedLedgers.map(l => ({
        'Row Number': l.rowNum,
        'Ledger Name': l.ledgerName,
        'Under Group': l.underGroup,
        'GSTIN': l.gstin || '',
        'State': l.state || '',
        'Country': l.country || '',
        'Opening Balance': l.openingBalance || '',
        'Status': l.excluded ? 'Excluded' : (l.errors.length > 0 ? 'Error' : 'Valid'),
        'Errors': l.errors.join('; '),
        'Warnings': l.warnings.join('; '),
        'Duplicate Status': l.duplicateMessage || 'No'
      }));
      reportName = 'ledger-validation-report.xlsx';
    } else if (importType === 'StockItem') {
      reportData = parsedStockItems.map(i => ({
        'Row Number': i.rowNum,
        'Stock Item Name': i.itemName,
        'Under Stock Group': i.underGroup,
        'Unit': i.unit,
        'GST Rate': i.gstRate || '',
        'Opening Qty': i.openingQty || '',
        'Status': i.excluded ? 'Excluded' : (i.errors.length > 0 ? 'Error' : 'Valid'),
        'Errors': i.errors.join('; '),
        'Warnings': i.warnings.join('; '),
        'Duplicate Status': i.duplicateMessage || 'No'
      }));
      reportName = 'stock-item-validation-report.xlsx';
    } else if (importType === 'StockGroup') {
      reportData = parsedStockGroups.map(g => ({
        'Row Number': g.rowNum,
        'Stock Group Name': g.groupName,
        'Under Parent Group': g.underGroup,
        'Status': g.excluded ? 'Excluded' : (g.errors.length > 0 ? 'Error' : 'Valid'),
        'Errors': g.errors.join('; ')
      }));
      reportName = 'stock-group-report.xlsx';
    } else if (importType === 'Unit') {
      reportData = parsedUnits.map(u => ({
        'Row Number': u.rowNum,
        'Symbol': u.symbol,
        'Formal Name': u.formalName || '',
        'UQC': u.uqc || '',
        'Status': u.excluded ? 'Excluded' : (u.errors.length > 0 ? 'Error' : 'Valid'),
        'Errors': u.errors.join('; ')
      }));
      reportName = 'units-report.xlsx';
    }

    const ws = XLSX.utils.json_to_sheet(reportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Validation Report");
    XLSX.writeFile(wb, reportName);
  };

  const generateMasterXML = async () => {
    if (!user) return;
    setIsProcessing(true);

    try {
      let xmlContent = '';
      let defaultFileName = 'tally-masters.xml';

      if (importType === 'Ledger') {
        const activeLedgers = parsedLedgers.filter(l => !l.excluded && l.isValid);
        if (activeLedgers.length === 0) {
          throw new Error("No valid, active rows to generate XML.");
        }
        xmlContent = generateLedgersXML(activeLedgers);
        defaultFileName = 'tally-ledger-masters.xml';
      } else if (importType === 'StockItem') {
        const activeStockItems = parsedStockItems.filter(i => !i.excluded && i.isValid);
        if (activeStockItems.length === 0) {
          throw new Error("No valid, active rows to generate XML.");
        }
        xmlContent = generateStockItemsXML(activeStockItems);
        defaultFileName = 'tally-stock-masters.xml';
      } else if (importType === 'StockGroup') {
        const activeGroups = parsedStockGroups.filter(g => !g.excluded && g.isValid);
        if (activeGroups.length === 0) {
          throw new Error("No valid, active rows to generate XML.");
        }
        xmlContent = generateStockGroupsXML(activeGroups);
        defaultFileName = 'tally-stock-groups.xml';
      } else if (importType === 'Unit') {
        const activeUnits = parsedUnits.filter(u => !u.excluded && u.isValid);
        if (activeUnits.length === 0) {
          throw new Error("No valid, active rows to generate XML.");
        }
        xmlContent = generateUnitsXML(activeUnits);
        defaultFileName = 'tally-units.xml';
      }

      const conversionPayload: any = {
        fileName: pendingFileName || defaultFileName,
        status: 'completed',
        voucherType: importType,
        xmlContent: xmlContent
      };
      if (getAppMode() === 'web') {
        conversionPayload.timestamp = serverTimestamp();
      }
      await saveConversion(user.uid, conversionPayload);

      downloadXML(xmlContent, defaultFileName);
      setCurrentStep('complete');
      setIsProcessing(false);
    } catch (err) {
      setError(err instanceof Error ? err.message : "XML generation failed");
      setIsProcessing(false);
    }
  };

  const startConversion = async () => {
    if (!user || pendingData.length === 0) return;
    setIsProcessing(true);
    setError(null);

    try {
      const rows = pendingData.map((row: any, rowIndex: number) => {
        let dateVal: any = undefined;
        let partyNameVal = '';
        let voucherNumberVal = '';
        let narrationVal = '';
        let amountVal = 0;
        let referenceVal = '';

        mappings.forEach(m => {
          if (m.tallyField === 'IGNORE') return;
          const value = row[m.excelColumn];
          if (value === undefined || value === null) return;

          switch (m.tallyField) {
            case 'DATE':
              dateVal = value;
              break;
            case 'PARTYNAME':
              partyNameVal = String(value).trim();
              break;
            case 'VOUCHERNUMBER':
              voucherNumberVal = String(value).trim();
              break;
            case 'NARRATION':
              narrationVal = String(value).trim();
              break;
            case 'AMOUNT':
              const parsedAmt = parseFloat(String(value).replace(/,/g, ''));
              amountVal = isNaN(parsedAmt) ? 0 : parsedAmt;
              break;
            case 'REFERENCE':
              referenceVal = String(value).trim();
              break;
          }
        });

        // 1. Validate and Normalize Date
        const dateNorm = normalizeTallyDate(dateVal);
        if (!dateNorm.isValid) {
          throw new Error(`Row ${rowIndex + 1}: Invalid Date value "${dateVal || ''}". Error: ${dateNorm.error}`);
        }

        // If partyNameVal is blank but a mapped suggestion/final ledger exists for that row, use that final ledger.
        if (!partyNameVal || partyNameVal.trim() === '') {
          const mappedTx = aiMappedTransactions[rowIndex];
          if (mappedTx && mappedTx.tallyLedger && mappedTx.tallyLedger.trim() !== '') {
            partyNameVal = mappedTx.tallyLedger.trim();
          }
        }

        // If no ledger exists, block XML generation with row-wise error instead of using "Unknown".
        if (!partyNameVal || partyNameVal.trim() === '') {
          throw new Error(`Row ${rowIndex + 1}: Particulars / Party Ledger is blank or unmapped. Please select or provide a valid ledger.`);
        }

        return {
          rowNo: rowIndex + 1,
          voucherType: selectedVoucherType,
          originalDate: String(dateVal || ''),
          normalizedDate: dateNorm.value,
          finalLedger: partyNameVal,
          bankLedger: selectedBankLedger || 'Bank Account',
          amount: amountVal,
          description: narrationVal,
          reference: referenceVal,
          voucherNumber: voucherNumberVal,
          sourceIdx: rowIndex,
        };
      });

      setVerificationRows(rows);
      setVerificationSourceStep('mapping');
      setCurrentStep('verification');
      setIsProcessing(false);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Verification failed");
      setIsProcessing(false);
    }
  };

  const buildMissingMastersXMLs = () => {
    let groupsXml = '';
    let unitsXml = '';
    let stockGroupsXml = '';
    let ledgersXml = '';
    let stockItemsXml = '';

    const groupsCreated = new Set<string>();
    const unitsCreated = new Set<string>();
    const stockGroupsCreated = new Set<string>();

    // 1. Groups
    missingLedgers.forEach(ml => {
      if (ml.action === 'Create' && ml.proposedGroup) {
        const gp = ml.proposedGroup;
        const gpLower = gp.toLowerCase();
        const existsInTally = tallyContext?.groups?.some(g => g.toLowerCase() === gpLower);
        if (!existsInTally && !groupsCreated.has(gpLower)) {
          groupsCreated.add(gpLower);
          groupsXml += generateGroupMasterXML(gp, 'Suspense A/c') + '\n';
        }
      }
    });

    // 2. Units
    missingStockItems.forEach(ms => {
      if (ms.action === 'Create' && ms.unit) {
        const u = ms.unit;
        const uLower = u.toLowerCase();
        const existsInTally = tallyContext?.units?.some(un => un.toLowerCase() === uLower);
        if (!existsInTally && !unitsCreated.has(uLower)) {
          unitsCreated.add(uLower);
          unitsXml += generateUnitMasterXML(u) + '\n';
        }
      }
    });

    // 3. Stock Groups
    missingStockItems.forEach(ms => {
      if (ms.action === 'Create' && ms.proposedStockGroup) {
        const sg = ms.proposedStockGroup;
        const sgLower = sg.toLowerCase();
        const existsInTally = tallyContext?.stockGroups?.some(s => s.toLowerCase() === sgLower);
        if (!existsInTally && !stockGroupsCreated.has(sgLower)) {
          stockGroupsCreated.add(sgLower);
          stockGroupsXml += generateStockGroupMasterXML(sg) + '\n';
        }
      }
    });

    // 4. Ledgers
    ledgersXml = generateMissingLedgerMastersXML(missingLedgers);

    // 5. Stock Items
    stockItemsXml = generateMissingStockItemMastersXML(missingStockItems);

    return { groupsXml, unitsXml, stockGroupsXml, ledgersXml, stockItemsXml };
  };

  const getReviewedVouchers = () => {
    const nameMap: Record<string, string> = {};
    missingLedgers.forEach(ml => {
      const key = ml.name.toLowerCase();
      if (ml.action === 'Replace' && ml.replacementName) {
        nameMap[key] = ml.replacementName;
      } else if (ml.action === 'Create' && ml.name) {
        nameMap[key] = ml.name;
      }
    });

    return verificationRows.map(row => {
      const updated = { ...row };
      const flKey = (updated.finalLedger || '').toLowerCase();
      const blKey = (updated.bankLedger || '').toLowerCase();
      if (nameMap[flKey]) {
        updated.finalLedger = nameMap[flKey];
      }
      if (nameMap[blKey]) {
        updated.bankLedger = nameMap[blKey];
      }
      return updated;
    });
  };

  const getReviewedInvoices = () => {
    const nameMap: Record<string, string> = {};
    missingLedgers.forEach(ml => {
      const key = ml.name.toLowerCase();
      if (ml.action === 'Replace' && ml.replacementName) {
        nameMap[key] = ml.replacementName;
      } else if (ml.action === 'Create' && ml.name) {
        nameMap[key] = ml.name;
      }
    });

    const stockMap: Record<string, string> = {};
    missingStockItems.forEach(ms => {
      const key = ms.name.toLowerCase();
      if (ms.action === 'Replace' && ms.replacementName) {
        stockMap[key] = ms.replacementName;
      } else if (ms.action === 'Create' && ms.name) {
        stockMap[key] = ms.name;
      }
    });

    return salesPurchaseInvoices.map(inv => {
      const updated = JSON.parse(JSON.stringify(inv)) as SalesPurchaseInvoice;
      const plKey = (updated.partyLedger || '').toLowerCase();
      const splKey = (updated.salesPurchaseLedger || '').toLowerCase();
      if (nameMap[plKey]) updated.partyLedger = nameMap[plKey];
      if (nameMap[splKey]) updated.salesPurchaseLedger = nameMap[splKey];

      if (updated.freightLedger && nameMap[updated.freightLedger.toLowerCase()]) {
        updated.freightLedger = nameMap[updated.freightLedger.toLowerCase()];
      }
      if (updated.packingLedger && nameMap[updated.packingLedger.toLowerCase()]) {
        updated.packingLedger = nameMap[updated.packingLedger.toLowerCase()];
      }
      if (updated.loadingLedger && nameMap[updated.loadingLedger.toLowerCase()]) {
        updated.loadingLedger = nameMap[updated.loadingLedger.toLowerCase()];
      }
      if (updated.insuranceLedger && nameMap[updated.insuranceLedger.toLowerCase()]) {
        updated.insuranceLedger = nameMap[updated.insuranceLedger.toLowerCase()];
      }
      if (updated.otherLedger1 && nameMap[updated.otherLedger1.toLowerCase()]) {
        updated.otherLedger1 = nameMap[updated.otherLedger1.toLowerCase()];
      }
      if (updated.otherLedger2 && nameMap[updated.otherLedger2.toLowerCase()]) {
        updated.otherLedger2 = nameMap[updated.otherLedger2.toLowerCase()];
      }
      if (updated.discountLedger && nameMap[updated.discountLedger.toLowerCase()]) {
        updated.discountLedger = nameMap[updated.discountLedger.toLowerCase()];
      }
      if (updated.roundOffLedger && nameMap[updated.roundOffLedger.toLowerCase()]) {
        updated.roundOffLedger = nameMap[updated.roundOffLedger.toLowerCase()];
      }

      updated.items.forEach(item => {
        const siKey = (item.stockItem || '').toLowerCase();
        if (stockMap[siKey]) item.stockItem = stockMap[siKey];

        if (item.cgstLedger && nameMap[item.cgstLedger.toLowerCase()]) {
          item.cgstLedger = nameMap[item.cgstLedger.toLowerCase()];
        }
        if (item.sgstLedger && nameMap[item.sgstLedger.toLowerCase()]) {
          item.sgstLedger = nameMap[item.sgstLedger.toLowerCase()];
        }
        if (item.igstLedger && nameMap[item.igstLedger.toLowerCase()]) {
          item.igstLedger = nameMap[item.igstLedger.toLowerCase()];
        }
      });

      return updated;
    });
  };

  const validateReview = (): string[] => {
    const errs: string[] = [];

    // Rule 1: Missing ledger action = Create but proposed group is blank
    missingLedgers.forEach(l => {
      if (l.action === 'Create' && !l.proposedGroup.trim()) {
        errs.push(`Ledger "${l.name}": Proposed Group cannot be blank.`);
      }
    });

    // Rule 2 & 3: Missing stock item action = Create but unit/hsn is blank
    missingStockItems.forEach(i => {
      if (i.action === 'Create') {
        if (!i.unit.trim()) {
          errs.push(`Stock Item "${i.name}": Unit is required.`);
        }
        if (!i.hsn.trim()) {
          errs.push(`Stock Item "${i.name}": HSN/SAC is required.`);
        }
        if (i.gstRate === undefined || isNaN(i.gstRate)) {
          errs.push(`Stock Item "${i.name}": GST Rate must be a valid number.`);
        }
      }
    });

    // Rule 4: Replace action selected but replacement master not selected
    missingLedgers.forEach(l => {
      if (l.action === 'Replace' && !l.replacementName) {
        errs.push(`Ledger "${l.name}": Replacement master ledger is not selected.`);
      }
    });
    missingStockItems.forEach(i => {
      if (i.action === 'Replace' && !i.replacementName) {
        errs.push(`Stock Item "${i.name}": Replacement stock item is not selected.`);
      }
    });

    // Rule 5: Missing master is ignored but still used in voucher XML
    missingLedgers.forEach(l => {
      if (l.action === 'Ignore') {
        errs.push(`Ledger "${l.name}": Cannot be ignored because it is used in final vouchers. Please select Create or Replace.`);
      }
    });
    missingStockItems.forEach(i => {
      if (i.action === 'Ignore') {
        errs.push(`Stock Item "${i.name}": Cannot be ignored because it is used in final invoices. Please select Create or Replace.`);
      }
    });

    return errs;
  };

  const checkMissingMastersAndProceed = (type: 'Vouchers' | 'SalesPurchase', rowsOverride?: any[]) => {
    setError(null);
    const ledgersList = tallyContext?.ledgers || [];
    const stockItemsList = tallyContext?.stockItems || [];
    const unitsList = tallyContext?.units || [];

    const usedLedgers: { name: string; source: string; type: string }[] = [];
    const usedStockItems: { name: string; source: string; unit?: string; hsn?: string; gstRate?: number }[] = [];

    if (type === 'Vouchers') {
      const targetRows = rowsOverride || verificationRows;
      targetRows.forEach((row, idx) => {
        const rowId = row.voucherNumber ? `${row.voucherType} No: ${row.voucherNumber}` : `${row.voucherType} Row ${idx + 1}`;
        if (row.finalLedger) {
          usedLedgers.push({ name: row.finalLedger, source: rowId, type: 'Party Ledger' });
        }
        if (row.bankLedger) {
          usedLedgers.push({ name: row.bankLedger, source: rowId, type: 'Bank/Cash Ledger' });
        } else if (selectedBankLedger) {
          usedLedgers.push({ name: selectedBankLedger, source: rowId, type: 'Bank/Cash Ledger' });
        }
      });
    } else {
      salesPurchaseInvoices.forEach((inv, idx) => {
        const invId = inv.invoiceNo ? `${inv.voucherType} No: ${inv.invoiceNo}` : `${inv.voucherType} Row ${idx + 1}`;
        if (inv.partyLedger) {
          usedLedgers.push({ name: inv.partyLedger, source: invId, type: 'Party Ledger' });
        }
        if (inv.salesPurchaseLedger) {
          usedLedgers.push({ name: inv.salesPurchaseLedger, source: invId, type: 'Sales/Purchase Ledger' });
        }
        if (inv.freightAmount && inv.freightLedger) {
          usedLedgers.push({ name: inv.freightLedger, source: invId, type: 'Freight Ledger' });
        }
        if (inv.packingAmount && inv.packingLedger) {
          usedLedgers.push({ name: inv.packingLedger, source: invId, type: 'Packing Ledger' });
        }
        if (inv.loadingAmount && inv.loadingLedger) {
          usedLedgers.push({ name: inv.loadingLedger, source: invId, type: 'Loading Ledger' });
        }
        if (inv.insuranceAmount && inv.insuranceLedger) {
          usedLedgers.push({ name: inv.insuranceLedger, source: invId, type: 'Insurance Ledger' });
        }
        if (inv.otherAmount1 && inv.otherLedger1) {
          usedLedgers.push({ name: inv.otherLedger1, source: invId, type: 'Other Ledger 1' });
        }
        if (inv.otherAmount2 && inv.otherLedger2) {
          usedLedgers.push({ name: inv.otherLedger2, source: invId, type: 'Other Ledger 2' });
        }
        if (inv.billDiscountAmount && inv.discountLedger) {
          usedLedgers.push({ name: inv.discountLedger, source: invId, type: 'Discount Ledger' });
        }
        if (inv.roundOffAmount && inv.roundOffLedger) {
          usedLedgers.push({ name: inv.roundOffLedger, source: invId, type: 'Round Off Ledger' });
        }

        inv.items.forEach((item, itemIdx) => {
          const itemSource = `${invId} Item ${itemIdx + 1}`;
          if (inv.voucherMode !== 'Accounting' && item.stockItem) {
            usedStockItems.push({
              name: item.stockItem,
              source: itemSource,
              unit: item.unit,
              hsn: item.hsn,
              gstRate: item.gstRate
            });
          }
          if (item.cgstLedger && item.cgstAmount) {
            usedLedgers.push({ name: item.cgstLedger, source: itemSource, type: 'CGST Ledger' });
          }
          if (item.sgstLedger && item.sgstAmount) {
            usedLedgers.push({ name: item.sgstLedger, source: itemSource, type: 'SGST Ledger' });
          }
          if (item.igstLedger && item.igstAmount) {
            usedLedgers.push({ name: item.igstLedger, source: itemSource, type: 'IGST Ledger' });
          }
        });
      });
    }

    const detectedLedgers = detectMissingLedgers(usedLedgers, ledgersList);
    const detectedItems = detectMissingStockItems(usedStockItems, stockItemsList, unitsList);

    if (detectedLedgers.length > 0 || detectedItems.length > 0) {
      setMissingLedgers(detectedLedgers);
      setMissingStockItems(detectedItems);
      setPendingExportType(type);
      setShowMissingMastersReview(true);
      setIsProcessing(false);
    } else {
      // Proceed directly without missing masters
      if (type === 'Vouchers') {
        const targetRows = rowsOverride || verificationRows;
        generateFinalXMLFromVerificationDirect(targetRows);
      } else {
        generateSalesPurchaseXMLFromVerificationDirect(salesPurchaseInvoices);
      }
    }
  };

  const generateFinalXMLFromVerificationDirect = async (rowsToUse: any[], includeMasters: boolean = false) => {
    if (!user || rowsToUse.length === 0) return;
    setIsProcessing(true);
    setError(null);

    try {
      const conversionPayload: any = {
        fileName: pendingFileName,
        status: 'processing',
        voucherType: 'Voucher'
      };
      if (getAppMode() === 'web') {
        conversionPayload.timestamp = serverTimestamp();
      }
      const conversionId = await saveConversion(user.uid, conversionPayload);

      const vouchers: TallyVoucher[] = rowsToUse.map(row => {
        const vType = row.voucherType;
        const amt = Math.abs(row.amount);

        const voucher: TallyVoucher = {
          date: row.normalizedDate,
          voucherType: vType,
          partyName: row.finalLedger,
          bankLedger: row.bankLedger,
          voucherNumber: row.voucherNumber || undefined,
          narration: row.description || undefined,
          reference: row.reference || undefined,
          ledgerEntries: []
        };

        if (vType === 'Receipt') {
          voucher.ledgerEntries.push({
            ledgerName: row.finalLedger,
            isDeemedPositive: 'No',
            isLastDeemedPositive: 'No',
            isPartyLedger: 'No',
            amount: amt
          });

          voucher.ledgerEntries.push({
            ledgerName: row.bankLedger || selectedBankLedger || 'Bank Account',
            isDeemedPositive: 'Yes',
            isLastDeemedPositive: 'Yes',
            isPartyLedger: 'Yes',
            amount: -amt
          });
        } else {
          voucher.ledgerEntries.push({
            ledgerName: row.finalLedger,
            isDeemedPositive: 'Yes',
            isLastDeemedPositive: 'Yes',
            isPartyLedger: 'No',
            amount: -amt
          });

          voucher.ledgerEntries.push({
            ledgerName: row.bankLedger || selectedBankLedger || 'Bank Account',
            isDeemedPositive: 'No',
            isLastDeemedPositive: 'No',
            isPartyLedger: 'Yes',
            amount: amt
          });
        }

        return voucher;
      });

      let xml = generateTallyXML(vouchers, useLedgerAsNarration);
      if (includeMasters) {
        xml = generateCombinedImportXML({
          voucherXml: xml,
          ...buildMissingMastersXMLs()
        });
      }

      await updateConversion(user.uid, conversionId, {
        status: 'completed',
        xmlContent: xml
      });

      const newRecord: ConversionRecord = {
        id: conversionId,
        fileName: pendingFileName,
        timestamp: { seconds: Math.floor(Date.now() / 1000) },
        status: 'completed',
        xmlContent: xml,
        voucherType: 'Voucher'
      };

      setConversions(prev => [newRecord, ...prev]);
      setCurrentStep('complete');
      setIsProcessing(false);
      setShowMissingMastersReview(false);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Conversion failed");
      setIsProcessing(false);
    }
  };

  const generateSalesPurchaseXMLFromVerificationDirect = async (invoicesToUse: SalesPurchaseInvoice[], includeMasters: boolean = false) => {
    if (!user || invoicesToUse.length === 0) return;
    setIsProcessing(true);
    setError(null);

    try {
      const errorInvoice = invoicesToUse.find(inv => !inv.isValid);
      if (errorInvoice) {
        throw new Error(`Invoice No "${errorInvoice.invoiceNo}" has errors. Please fix all errors before generating XML.`);
      }

      // Check balancing
      const balancingResults = invoicesToUse.map(inv => verifySalesPurchaseBalancing(inv));
      const unbalanced = balancingResults.filter(r => !r.isBalanced);
      if (unbalanced.length > 0) {
        setSalesPurchaseBalancingErrors(unbalanced);
        throw new Error(`XML generation blocked: ${unbalanced.length} Sales/Purchase invoice(s) are unbalanced.`);
      } else {
        setSalesPurchaseBalancingErrors([]);
      }

      const conversionPayload: any = {
        fileName: pendingFileName,
        status: 'processing',
        voucherType: 'SalesPurchase'
      };
      if (getAppMode() === 'web') {
        conversionPayload.timestamp = serverTimestamp();
      }
      const conversionId = await saveConversion(user.uid, conversionPayload);

      let xml = generateSalesPurchaseXML(invoicesToUse, companyState);
      if (includeMasters) {
        xml = generateCombinedImportXML({
          voucherXml: xml,
          ...buildMissingMastersXMLs()
        });
      }

      await updateConversion(user.uid, conversionId, {
        status: 'completed',
        xmlContent: xml
      });

      const newRecord: ConversionRecord = {
        id: conversionId,
        fileName: pendingFileName,
        timestamp: { seconds: Math.floor(Date.now() / 1000) },
        status: 'completed',
        xmlContent: xml,
        voucherType: 'SalesPurchase'
      };

      setConversions(prev => [newRecord, ...prev]);
      setCurrentStep('complete');
      setIsProcessing(false);
      setShowMissingMastersReview(false);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Conversion failed");
      setIsProcessing(false);
    }
  };

  const downloadReviewedExcel = () => {
    const wb = XLSX.utils.book_new();
    if (pendingExportType === 'Vouchers') {
      const exportRows = getReviewedVouchers().map((row, idx) => ({
        'Date': row.date,
        'Voucher Type': row.voucherType,
        'Voucher No': row.voucherNumber || '',
        'Ledger Name': row.finalLedger,
        'Bank/Cash Ledger': row.bankLedger || selectedBankLedger || 'Bank Account',
        'Amount': Math.abs(row.amount),
        'Narration': row.description || '',
        'Reference': row.reference || ''
      }));
      const ws = XLSX.utils.json_to_sheet(exportRows);
      XLSX.utils.book_append_sheet(wb, ws, "Reviewed_Vouchers");
    } else if (pendingExportType === 'Journal') {
      const exportRows: any[] = [];
      journalGroups.forEach(g => {
        g.lines.forEach(l => {
          exportRows.push({
            'Voucher Date': g.voucherDate,
            'Voucher No': g.voucherNo,
            'Ledger Name': l.ledgerName,
            'Dr/Cr': l.drCr,
            'Amount': l.amount,
            'Narration': l.narration || '',
            'Reference': l.reference || '',
            'Cost Centre': l.costCentre || '',
            'Bill Reference': l.billReference || '',
            'Remarks': l.remarks || ''
          });
        });
      });
      const ws = XLSX.utils.json_to_sheet(exportRows);
      XLSX.utils.book_append_sheet(wb, ws, "Reviewed_Journal");
    } else {
      const exportRows: any[] = [];
      getReviewedInvoices().forEach(inv => {
        inv.items.forEach(item => {
          exportRows.push({
            'Invoice Date': inv.invoiceDate,
            'Invoice No': inv.invoiceNo,
            'Party Ledger': inv.partyLedger,
            'Sales/Purchase Ledger': inv.salesPurchaseLedger,
            'Place of Supply': inv.placeOfSupply || '',
            'Stock Item': item.stockItem,
            'Quantity': item.quantity,
            'Unit': item.unit,
            'Rate': item.rate,
            'Amount': item.itemAmount,
            'GST Rate %': item.gstRate,
            'HSN/SAC': item.hsn || '',
            'Narration': inv.narration || '',
            'Reference': inv.reference || ''
          });
        });
      });
      const ws = XLSX.utils.json_to_sheet(exportRows);
      XLSX.utils.book_append_sheet(wb, ws, "Reviewed_Sales_Purchase");
    }
    XLSX.writeFile(wb, `TallyGen_Reviewed_${pendingExportType}.xlsx`);
  };

  const generateFinalXMLFromVerification = async () => {
    checkMissingMastersAndProceed('Vouchers');
  };

  const generateSalesPurchaseXMLFromVerification = async () => {
    setError(null);
    const balancingResults = salesPurchaseInvoices.map(inv => verifySalesPurchaseBalancing(inv));
    const unbalanced = balancingResults.filter(r => !r.isBalanced);
    if (unbalanced.length > 0) {
      setSalesPurchaseBalancingErrors(unbalanced);
      setError(`XML generation blocked: ${unbalanced.length} Sales/Purchase invoice(s) are unbalanced.`);
      return;
    } else {
      setSalesPurchaseBalancingErrors([]);
    }

    checkMissingMastersAndProceed('SalesPurchase');
  };

  const downloadXML = (content: string, fileName: string) => {
    const blob = new Blob([content], { type: 'text/xml' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName.replace(/\.[^/.]+$/, "") + ".xml";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const getFilteredLedgers = () => {
    return parsedLedgers.filter(l => {
      const matchesSearch = l.ledgerName.toLowerCase().includes(masterReviewSearch.toLowerCase()) || 
                            l.underGroup.toLowerCase().includes(masterReviewSearch.toLowerCase()) ||
                            (l.gstin && l.gstin.toLowerCase().includes(masterReviewSearch.toLowerCase()));
      if (!matchesSearch) return false;
      if (masterReviewFilter === 'all') return true;
      if (masterReviewFilter === 'valid') return l.isValid && !l.excluded;
      if (masterReviewFilter === 'invalid') return !l.isValid && !l.excluded;
      if (masterReviewFilter === 'warning') return l.warnings.length > 0 && !l.excluded;
      if (masterReviewFilter === 'duplicate') return (l.isDuplicate || l.isPossibleDuplicate) && !l.excluded;
      if (masterReviewFilter === 'excluded') return l.excluded;
      return true;
    });
  };

  const getFilteredStockItems = () => {
    return parsedStockItems.filter(i => {
      const matchesSearch = i.itemName.toLowerCase().includes(masterReviewSearch.toLowerCase()) || 
                            i.underGroup.toLowerCase().includes(masterReviewSearch.toLowerCase()) ||
                            i.unit.toLowerCase().includes(masterReviewSearch.toLowerCase());
      if (!matchesSearch) return false;
      if (masterReviewFilter === 'all') return true;
      if (masterReviewFilter === 'valid') return i.isValid && !i.excluded;
      if (masterReviewFilter === 'invalid') return !i.isValid && !i.excluded;
      if (masterReviewFilter === 'warning') return i.warnings.length > 0 && !i.excluded;
      if (masterReviewFilter === 'duplicate') return (i.isDuplicate || i.isPossibleDuplicate) && !i.excluded;
      if (masterReviewFilter === 'excluded') return i.excluded;
      return true;
    });
  };

  const getFilteredStockGroups = () => {
    return parsedStockGroups.filter(g => {
      const matchesSearch = g.groupName.toLowerCase().includes(masterReviewSearch.toLowerCase()) || 
                            g.underGroup.toLowerCase().includes(masterReviewSearch.toLowerCase());
      if (!matchesSearch) return false;
      if (masterReviewFilter === 'all') return true;
      if (masterReviewFilter === 'valid') return g.isValid && !g.excluded;
      if (masterReviewFilter === 'invalid') return !g.isValid && !g.excluded;
      if (masterReviewFilter === 'excluded') return g.excluded;
      return true;
    });
  };

  const getFilteredUnits = () => {
    return parsedUnits.filter(u => {
      const matchesSearch = u.symbol.toLowerCase().includes(masterReviewSearch.toLowerCase()) || 
                            (u.formalName && u.formalName.toLowerCase().includes(masterReviewSearch.toLowerCase()));
      if (!matchesSearch) return false;
      if (masterReviewFilter === 'all') return true;
      if (masterReviewFilter === 'valid') return u.isValid && !u.excluded;
      if (masterReviewFilter === 'invalid') return !u.isValid && !u.excluded;
      if (masterReviewFilter === 'excluded') return u.excluded;
      return true;
    });
  };

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-zinc-50">
        <Loader2 className="w-8 h-8 animate-spin text-zinc-400" />
      </div>
    );
  }

  if (!user) {
    return (
      <div className="min-h-screen flex flex-col items-center justify-center bg-zinc-50 p-4">
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="max-w-md w-full bg-white p-8 rounded-2xl shadow-sm border border-zinc-200 text-center"
        >
          <div className="w-16 h-16 bg-zinc-900 rounded-2xl flex items-center justify-center mx-auto mb-6">
            <FileSpreadsheet className="w-8 h-8 text-white" />
          </div>
          <h1 className="text-2xl font-bold text-zinc-900 mb-2">Tally XML Generator</h1>
          <p className="text-zinc-500 mb-8">Convert your Excel/PDF bank statements into Tally-compatible XML format with local deterministic mapping.</p>
          <button 
            onClick={handleLogin}
            className="w-full flex items-center justify-center gap-2 bg-zinc-900 text-white py-3 px-4 rounded-xl font-medium hover:bg-zinc-800 transition-colors"
          >
            <LogIn className="w-5 h-5" />
            Sign in with Google
          </button>
        </motion.div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-zinc-50 text-zinc-900 font-sans">
      {/* Header */}
      <header className="bg-white border-b border-zinc-200 sticky top-0 z-10">
        <div className="max-w-7xl mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 bg-zinc-900 rounded-lg flex items-center justify-center">
              <FileSpreadsheet className="w-5 h-5 text-white" />
            </div>
            <div className="flex flex-col">
              <span className="font-bold text-lg leading-none">TallyGen Pro</span>
              <span className="text-[10px] text-zinc-400 font-normal">v1.0.0</span>
            </div>
            {getAppMode() === 'desktop-offline' && (
              <span className="bg-amber-100 text-amber-800 text-xs font-semibold px-2.5 py-0.5 rounded border border-amber-200 flex items-center gap-1">
                <span className="w-2 h-2 rounded-full bg-amber-500 animate-pulse"></span>
                Desktop Offline
              </span>
            )}
          </div>

          <div className="flex items-center gap-4">
            <button 
              onClick={() => setCurrentStep('context')}
              className={`flex items-center gap-2 px-3 py-1.5 rounded-full text-sm font-medium transition-colors ${
                tallyContext ? 'bg-green-50 text-green-700 border border-green-100' : 'bg-amber-50 text-amber-700 border border-amber-100'
              }`}
            >
              <Database className="w-4 h-4" />
              {tallyContext ? `${tallyContext.ledgers.length} Ledgers Loaded` : 'Load Tally Masters'}
            </button>

            <button 
              onClick={() => setShowRestartConfirm(true)}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-full text-sm font-medium transition-colors bg-zinc-100 hover:bg-zinc-200 border border-zinc-200 text-zinc-700"
              title="Restart / New Task"
              id="header-restart-btn"
            >
              <RotateCcw className="w-4 h-4" />
              <span className="hidden sm:inline">Restart</span>
            </button>

            <div className="flex items-center gap-2 px-3 py-1.5 bg-zinc-100 rounded-full">
              {user.photoURL ? (
                <img src={user.photoURL} alt="" className="w-6 h-6 rounded-full" referrerPolicy="no-referrer" />
              ) : (
                <div className="w-6 h-6 rounded-full bg-zinc-800 text-white flex items-center justify-center text-[10px] font-bold">
                  {user.displayName?.charAt(0).toUpperCase() || 'U'}
                </div>
              )}
              <span className="text-sm font-medium hidden md:block">{user.displayName}</span>
            </div>
            {getAppMode() !== 'desktop-offline' && (
              <button 
                onClick={handleLogout}
                className="p-2 hover:bg-zinc-100 rounded-full transition-colors text-zinc-500"
                title="Logout"
              >
                <LogOut className="w-5 h-5" />
              </button>
            )}
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-8">
        {showMissingMastersReview ? (
          <motion.section
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className="bg-white p-8 rounded-2xl border border-zinc-200 shadow-sm space-y-8"
          >
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-4 border-b border-zinc-100">
              <div>
                <h2 className="text-2xl font-bold flex items-center gap-3 text-zinc-900">
                  <Database className="w-6 h-6 text-zinc-900" />
                  Missing Masters Review
                </h2>
                <p className="text-sm text-zinc-500 mt-1">
                  These masters are not found in uploaded Tally data. TallyGen can create them before voucher import. Please review group/unit details before downloading XML.
                </p>
              </div>
              <button
                type="button"
                onClick={() => setShowMissingMastersReview(false)}
                className="px-4 py-2 text-xs font-semibold text-zinc-500 hover:text-zinc-900 hover:bg-zinc-100 rounded-lg transition-colors"
              >
                Back to Verification
              </button>
            </div>

            {/* Error alerts if any validation fails */}
            {(() => {
              const validationErrors = validateReview();
              if (validationErrors.length > 0) {
                return (
                  <div className="p-4 bg-red-50 border border-red-200 rounded-xl space-y-2">
                    <div className="flex items-center gap-2 text-red-800 font-semibold text-sm">
                      <AlertCircle className="w-4 h-4 shrink-0" />
                      <span>XML Generation Blocked (Fix errors to export)</span>
                    </div>
                    <ul className="list-disc list-inside text-xs text-red-700 space-y-1 pl-1">
                      {validationErrors.map((err, i) => (
                        <li key={i}>{err}</li>
                      ))}
                    </ul>
                  </div>
                );
              }
              return (
                <div className="p-4 bg-emerald-50 border border-emerald-100 rounded-xl flex items-center gap-2.5 text-emerald-800 text-sm">
                  <CheckCircle2 className="w-4.5 h-4.5 shrink-0 text-emerald-600" />
                  <span className="font-medium">All reviews valid! You can now download the XML files.</span>
                </div>
              );
            })()}

            {/* Missing Ledgers Table */}
            {missingLedgers.length > 0 && (
              <div className="space-y-4">
                <h3 className="font-bold text-zinc-800 text-base flex items-center gap-2">
                  <span className="w-2.5 h-2.5 bg-amber-500 rounded-full"></span>
                  Missing Ledgers ({missingLedgers.length})
                </h3>
                <div className="overflow-x-auto border border-zinc-200 rounded-xl">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="bg-zinc-50 border-b border-zinc-200 text-zinc-600 font-semibold">
                        <th className="p-3">Ledger Name</th>
                        <th className="p-3">Proposed Group</th>
                        <th className="p-3">Source Row / Voucher</th>
                        <th className="p-3">Type</th>
                        <th className="p-3">Action</th>
                        <th className="p-3">Replace With</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-zinc-100">
                      {missingLedgers.map(l => (
                        <tr key={l.id} className="hover:bg-zinc-50/50 transition-colors">
                          <td className="p-3">
                            <input
                              type="text"
                              value={l.name}
                              onChange={e => handleUpdateLedger(l.id, { name: e.target.value })}
                              className="w-full px-2 py-1 border border-zinc-200 rounded-md focus:border-zinc-900 focus:ring-1 focus:ring-zinc-900 outline-none font-medium text-zinc-800"
                            />
                            {l.possibleMatches.length > 0 && (
                              <div className="text-[10px] text-amber-600 mt-1 pl-1">
                                Possible Matches: {l.possibleMatches.join(', ')}
                              </div>
                            )}
                          </td>
                          <td className="p-3">
                            <input
                              type="text"
                              value={l.proposedGroup}
                              placeholder="Suspense"
                              onChange={e => handleUpdateLedger(l.id, { proposedGroup: e.target.value })}
                              className="w-full px-2 py-1 border border-zinc-200 rounded-md focus:border-zinc-900 focus:ring-1 focus:ring-zinc-900 outline-none text-zinc-700"
                            />
                          </td>
                          <td className="p-3 text-zinc-500 max-w-[150px] truncate" title={l.sourceRowOrVoucher}>
                            {l.sourceRowOrVoucher}
                          </td>
                          <td className="p-3">
                            <span className="px-2 py-0.5 bg-zinc-100 text-zinc-600 rounded font-medium text-[10px]">
                              {l.type}
                            </span>
                          </td>
                          <td className="p-3">
                            <select
                              value={l.action}
                              onChange={e => handleUpdateLedger(l.id, { action: e.target.value as any })}
                              className="px-2 py-1 border border-zinc-200 rounded-md bg-white focus:border-zinc-900 focus:ring-1 focus:ring-zinc-900 outline-none font-medium"
                            >
                              <option value="Create">Create</option>
                              <option value="Replace">Replace</option>
                              <option value="Ignore">Ignore</option>
                            </select>
                          </td>
                          <td className="p-3">
                            <select
                              value={l.replacementName}
                              disabled={l.action !== 'Replace'}
                              onChange={e => handleUpdateLedger(l.id, { replacementName: e.target.value })}
                              className="w-full px-2 py-1 border border-zinc-200 rounded-md bg-white disabled:bg-zinc-50 disabled:text-zinc-400 focus:border-zinc-900 focus:ring-1 focus:ring-zinc-900 outline-none"
                            >
                              <option value="">-- Select Master --</option>
                              {(tallyContext?.ledgers || []).map(m => (
                                <option key={m} value={m}>{m}</option>
                              ))}
                            </select>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Missing Stock Items Table */}
            {missingStockItems.length > 0 && (
              <div className="space-y-4">
                <h3 className="font-bold text-zinc-800 text-base flex items-center gap-2">
                  <span className="w-2.5 h-2.5 bg-rose-500 rounded-full"></span>
                  Missing Stock Items ({missingStockItems.length})
                </h3>
                <div className="overflow-x-auto border border-zinc-200 rounded-xl">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="bg-zinc-50 border-b border-zinc-200 text-zinc-600 font-semibold">
                        <th className="p-3">Stock Item Name</th>
                        <th className="p-3">Proposed Stock Group</th>
                        <th className="p-3">Unit</th>
                        <th className="p-3">HSN/SAC</th>
                        <th className="p-3">GST Rate %</th>
                        <th className="p-3">Source Row / Voucher</th>
                        <th className="p-3">Action</th>
                        <th className="p-3">Replace With</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-zinc-100">
                      {missingStockItems.map(i => (
                        <tr key={i.id} className="hover:bg-zinc-50/50 transition-colors">
                          <td className="p-3">
                            <input
                              type="text"
                              value={i.name}
                              onChange={e => handleUpdateStockItem(i.id, { name: e.target.value })}
                              className="w-full px-2 py-1 border border-zinc-200 rounded-md focus:border-zinc-900 focus:ring-1 focus:ring-zinc-900 outline-none font-medium text-zinc-800"
                            />
                            {i.possibleMatches.length > 0 && (
                              <div className="text-[10px] text-amber-600 mt-1 pl-1">
                                Possible Matches: {i.possibleMatches.join(', ')}
                              </div>
                            )}
                          </td>
                          <td className="p-3">
                            <input
                              type="text"
                              value={i.proposedStockGroup}
                              placeholder="Suspense Stock Group"
                              onChange={e => handleUpdateStockItem(i.id, { proposedStockGroup: e.target.value })}
                              className="w-full px-2 py-1 border border-zinc-200 rounded-md focus:border-zinc-900 focus:ring-1 focus:ring-zinc-900 outline-none text-zinc-700"
                            />
                          </td>
                          <td className="p-3">
                            <select
                              value={i.unit}
                              onChange={e => handleUpdateStockItem(i.id, { unit: e.target.value })}
                              className="w-full px-2 py-1 border border-zinc-200 rounded-md bg-white focus:border-zinc-900 focus:ring-1 focus:ring-zinc-900 outline-none text-zinc-700"
                            >
                              <option value="">-- Select --</option>
                              {Array.from(new Set([
                                ...((tallyContext?.units || [])),
                                'NOS', 'PCS', 'BOX', 'KG', 'MTR', 'BAG'
                              ])).map(u => (
                                <option key={u} value={u}>{u}</option>
                              ))}
                            </select>
                          </td>
                          <td className="p-3">
                            <input
                              type="text"
                              value={i.hsn}
                              placeholder="e.g. 8471"
                              onChange={e => handleUpdateStockItem(i.id, { hsn: e.target.value })}
                              className="w-full px-2 py-1 border border-zinc-200 rounded-md focus:border-zinc-900 focus:ring-1 focus:ring-zinc-900 outline-none text-zinc-700"
                            />
                          </td>
                          <td className="p-3">
                            <input
                              type="number"
                              value={i.gstRate}
                              onChange={e => handleUpdateStockItem(i.id, { gstRate: parseFloat(e.target.value) || 0 })}
                              className="w-16 px-2 py-1 border border-zinc-200 rounded-md focus:border-zinc-900 focus:ring-1 focus:ring-zinc-900 outline-none text-zinc-700"
                            />
                          </td>
                          <td className="p-3 text-zinc-500 max-w-[120px] truncate" title={i.sourceRowOrVoucher}>
                            {i.sourceRowOrVoucher}
                          </td>
                          <td className="p-3">
                            <select
                              value={i.action}
                              onChange={e => handleUpdateStockItem(i.id, { action: e.target.value as any })}
                              className="px-2 py-1 border border-zinc-200 rounded-md bg-white focus:border-zinc-900 focus:ring-1 focus:ring-zinc-900 outline-none font-medium"
                            >
                              <option value="Create">Create</option>
                              <option value="Replace">Replace</option>
                              <option value="Ignore">Ignore</option>
                            </select>
                          </td>
                          <td className="p-3">
                            <select
                              value={i.replacementName}
                              disabled={i.action !== 'Replace'}
                              onChange={e => handleUpdateStockItem(i.id, { replacementName: e.target.value })}
                              className="w-full px-2 py-1 border border-zinc-200 rounded-md bg-white disabled:bg-zinc-50 disabled:text-zinc-400 focus:border-zinc-900 focus:ring-1 focus:ring-zinc-900 outline-none"
                            >
                              <option value="">-- Select Master --</option>
                              {(tallyContext?.stockItems || []).map(m => (
                                <option key={m} value={m}>{m}</option>
                              ))}
                            </select>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Final Export Controls */}
            <div className="pt-6 border-t border-zinc-100 flex flex-col md:flex-row md:items-center justify-between gap-4">
              <div className="text-xs text-zinc-500">
                Correct master details will be embedded automatically. Standard Tally structure is enforced.
              </div>
              <div className="flex flex-wrap items-center gap-3">
                <button
                  type="button"
                  onClick={downloadReviewedExcel}
                  className="px-4 py-2.5 border border-zinc-200 hover:bg-zinc-50 text-zinc-700 text-xs font-semibold rounded-xl transition-all"
                >
                  Download Reviewed Excel
                </button>
                <button
                  type="button"
                  onClick={async () => {
                    const validationErrors = validateReview();
                    if (validationErrors.length > 0) return;
                    setIsProcessing(true);
                    if (pendingExportType === 'Vouchers') {
                      const finalVouchers = getReviewedVouchers();
                      await generateFinalXMLFromVerificationDirect(finalVouchers, true);
                    } else if (pendingExportType === 'Journal') {
                      await generateFinalJournalXMLDirect(journalGroups, true);
                    } else {
                      const finalInvoices = getReviewedInvoices();
                      await generateSalesPurchaseXMLFromVerificationDirect(finalInvoices, true);
                    }
                  }}
                  disabled={validateReview().length > 0}
                  className="px-5 py-2.5 bg-zinc-900 hover:bg-zinc-800 disabled:opacity-40 text-white text-xs font-semibold rounded-xl shadow-sm transition-all flex items-center gap-2"
                >
                  <CheckCircle2 className="w-4 h-4" />
                  Download Combined XML
                </button>

                {getAppMode() === 'desktop-offline' && isDirectTallyAvailable() && tallyStatus === 'Connected' && (
                  <button
                    type="button"
                    onClick={handlePushCombinedXmlOnTheFly}
                    disabled={validateReview().length > 0}
                    className="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-40 text-white text-xs font-semibold rounded-xl shadow-sm transition-all flex items-center gap-2"
                  >
                    <RefreshCw className="w-4 h-4" />
                    Push Directly to Tally
                  </button>
                )}
                <button
                  type="button"
                  onClick={async () => {
                    const validationErrors = validateReview();
                    if (validationErrors.length > 0) return;
                    setIsProcessing(true);
                    // Download masters
                    const mastersXml = generateCombinedImportXML({
                      voucherXml: '',
                      ...buildMissingMastersXMLs()
                    });
                    downloadXML(mastersXml, 'TallyGen_Missing_Masters.xml');

                    // Download vouchers
                    if (pendingExportType === 'Vouchers') {
                      const finalVouchers = getReviewedVouchers();
                      await generateFinalXMLFromVerificationDirect(finalVouchers, false);
                    } else if (pendingExportType === 'Journal') {
                      await generateFinalJournalXMLDirect(journalGroups, false);
                    } else {
                      const finalInvoices = getReviewedInvoices();
                      await generateSalesPurchaseXMLFromVerificationDirect(finalInvoices, false);
                    }
                  }}
                  disabled={validateReview().length > 0}
                  className="px-4 py-2.5 border border-zinc-300 hover:bg-zinc-100 disabled:opacity-40 text-zinc-800 text-xs font-semibold rounded-xl transition-all"
                >
                  Download Separate Files
                </button>
              </div>
            </div>
          </motion.section>
        ) : (
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          {/* Left: Main Workflow */}
          <div className="lg:col-span-2 space-y-6">
            <AnimatePresence mode="wait">
              {currentStep === 'context' && (
                <motion.section 
                  key="context"
                  initial={{ opacity: 0, y: 20 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -20 }}
                  className="bg-white p-8 rounded-2xl border border-zinc-200 shadow-sm"
                >
                  <div className="flex items-center justify-between mb-6">
                    <h2 className="text-2xl font-bold flex items-center gap-3">
                      <Database className="w-6 h-6 text-zinc-900" />
                      Tally Context Setup
                    </h2>
                    <button onClick={() => setCurrentStep('upload')} className="text-zinc-500 hover:text-zinc-900">Close</button>
                  </div>
                  <p className="text-zinc-600 mb-8">Upload your Tally Masters and Day Book XML files. You can select both files at once to update your context in one go.</p>
                  
                  <div className="border-2 border-dashed border-zinc-200 rounded-xl p-12 text-center hover:border-zinc-400 transition-colors cursor-pointer relative group">
                    <input 
                      type="file" 
                      accept=".xml" 
                      multiple
                      onChange={handleCombinedTallyUpload}
                      className="absolute inset-0 opacity-0 cursor-pointer"
                      disabled={isContextLoading}
                    />
                    <div className="space-y-4">
                      <div className="w-12 h-12 bg-zinc-100 rounded-full flex items-center justify-center mx-auto group-hover:bg-zinc-200 transition-colors">
                        {isContextLoading ? (
                          <Loader2 className="w-6 h-6 animate-spin text-zinc-600" />
                        ) : (
                          <FileUp className="w-6 h-6 text-zinc-600" />
                        )}
                      </div>
                      <div>
                        <p className="font-medium">Upload Tally XMLs (Masters & Day Book)</p>
                        <p className="text-sm text-zinc-500">Select multiple files to upload both at once</p>
                      </div>
                    </div>
                  </div>

                  {tallyContext && (
                    <div className="mt-8 p-6 bg-zinc-50 rounded-xl border border-zinc-100 grid grid-cols-3 gap-4">
                      <div>
                        <p className="text-xs text-zinc-500 uppercase tracking-wider mb-1">Ledgers</p>
                        <p className="text-2xl font-bold text-zinc-900">{tallyContext.ledgers.length}</p>
                      </div>
                      <div>
                        <p className="text-xs text-zinc-500 uppercase tracking-wider mb-1">Groups</p>
                        <p className="text-2xl font-bold text-zinc-900">{tallyContext.groups.length}</p>
                      </div>
                      <div>
                        <p className="text-xs text-zinc-500 uppercase tracking-wider mb-1">History</p>
                        <p className="text-2xl font-bold text-zinc-900">{tallyContext.historicalMappings?.length || 0}</p>
                      </div>
                    </div>
                  )}
                </motion.section>
              )}

              {currentStep === 'upload' && (
                <motion.section 
                  key="upload"
                  initial={{ opacity: 0, x: -20 }}
                  animate={{ opacity: 1, x: 0 }}
                  exit={{ opacity: 0, x: 20 }}
                  className="bg-white p-6 rounded-2xl border border-zinc-200 shadow-sm"
                >
                  {(tallyContext === null || (!proceededWithContext && !skippedContext)) ? (
                    // Step 1: Load Existing Tally Data SETUP PANEL
                    <div className="space-y-6">
                      {/* Title and Description */}
                      <div className="border-b border-zinc-100 pb-5" id="setup-panel-header">
                        <h2 className="text-2xl font-black text-zinc-900 tracking-tight flex items-center gap-3" id="setup-panel-title">
                          <Database className="w-6 h-6 text-zinc-900" />
                          Step 1: Load Existing Tally Data
                        </h2>
                        <p className="text-zinc-500 text-sm mt-2 leading-relaxed" id="setup-panel-desc">
                          Upload your existing Tally Masters XML / All Items XML and optionally Daybook / Transaction XML so the system can load ledgers, groups, stock groups, units, stock items and old narration history.
                        </p>
                      </div>

                      {/* Upload options grid */}
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4" id="setup-upload-options">
                        {/* Option 1: Masters XML */}
                        <div className="relative group bg-zinc-50 hover:bg-zinc-100/70 border border-zinc-200 rounded-2xl p-6 transition-all flex flex-col justify-between" id="setup-opt-masters">
                          <div>
                            <div className="w-10 h-10 bg-white shadow-sm border border-zinc-100 rounded-xl flex items-center justify-center text-zinc-900 mb-4 group-hover:scale-105 transition-transform">
                              <Database className="w-5 h-5" />
                            </div>
                            <h3 className="font-bold text-zinc-950 text-base">Tally Masters / All Items XML</h3>
                            <p className="text-zinc-500 text-xs mt-1.5 leading-relaxed">
                              Loads your entire ledger list, parent groups, stock items, groups, and units.
                            </p>
                          </div>
                          <div className="mt-5">
                            <label className="inline-flex items-center gap-2 px-4 py-2 bg-white hover:bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-semibold text-zinc-700 cursor-pointer shadow-sm transition-colors">
                              <Upload className="w-3.5 h-3.5" />
                              Upload Masters XML
                              <input 
                                type="file" 
                                accept=".xml" 
                                onChange={handleTallyMastersUpload} 
                                className="hidden" 
                                id="masters-xml-upload-input"
                              />
                            </label>
                          </div>
                        </div>

                        {/* Option 2: Daybook XML */}
                        <div className="relative group bg-zinc-50 hover:bg-zinc-100/70 border border-zinc-200 rounded-2xl p-6 transition-all flex flex-col justify-between" id="setup-opt-daybook">
                          <div>
                            <div className="w-10 h-10 bg-white shadow-sm border border-zinc-100 rounded-xl flex items-center justify-center text-zinc-900 mb-4 group-hover:scale-105 transition-transform">
                              <FileText className="w-5 h-5" />
                            </div>
                            <h3 className="font-bold text-zinc-950 text-base">Daybook / Transaction XML (Optional)</h3>
                            <p className="text-zinc-500 text-xs mt-1.5 leading-relaxed">
                              Extracts old entry ledger names and narration mapping histories for AI learning.
                            </p>
                          </div>
                          <div className="mt-5">
                            <label className="inline-flex items-center gap-2 px-4 py-2 bg-white hover:bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-semibold text-zinc-700 cursor-pointer shadow-sm transition-colors">
                              <Upload className="w-3.5 h-3.5" />
                              Upload Daybook XML
                              <input 
                                type="file" 
                                accept=".xml" 
                                onChange={handleDaybookUpload} 
                                className="hidden" 
                                id="daybook-xml-upload-input"
                              />
                            </label>
                          </div>
                        </div>

                        {/* Option 3: Combined upload */}
                        <div className="md:col-span-2 relative group bg-zinc-900 text-white rounded-2xl p-6 transition-all flex flex-col md:flex-row md:items-center justify-between gap-4" id="setup-opt-combined">
                          <div className="max-w-md">
                            <h3 className="font-bold text-lg text-white">Upload Combined / Multiple Tally XML Files</h3>
                            <p className="text-zinc-300 text-xs mt-1 leading-relaxed">
                              Select multiple Tally XML exports (Masters, Items, and Vouchers) together. The system will merge and process everything automatically.
                            </p>
                          </div>
                          <div className="shrink-0">
                            <label className="inline-flex items-center gap-2 px-5 py-2.5 bg-white hover:bg-zinc-100 text-zinc-900 rounded-xl text-xs font-bold cursor-pointer shadow-md transition-colors">
                              <FileUp className="w-4 h-4" />
                              Select Combined Files
                              <input 
                                type="file" 
                                accept=".xml" 
                                multiple 
                                onChange={handleCombinedTallyUpload} 
                                className="hidden" 
                                id="combined-xml-upload-input"
                              />
                            </label>
                          </div>
                        </div>

                        {/* Option 4: Direct Tally Connection (Offline only / disabled on Web) */}
                        <div className="md:col-span-2 relative bg-zinc-50 border border-zinc-200 rounded-2xl p-6" id="setup-opt-direct-tally">
                          <div className="flex flex-col md:flex-row md:items-start justify-between gap-6">
                            <div className="flex-1 space-y-4">
                              <div className="flex items-center gap-3">
                                <div className="w-10 h-10 bg-emerald-50 border border-emerald-100 rounded-xl flex items-center justify-center text-emerald-600">
                                  <RefreshCw className="w-5 h-5 animate-spin-slow" />
                                </div>
                                <div>
                                  <h3 className="font-bold text-zinc-950 text-base flex items-center gap-2">
                                    Direct Tally Connection
                                    {getAppMode() === 'desktop-offline' && (
                                      <span className="text-[10px] bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded-full font-bold uppercase">
                                        Offline Only
                                      </span>
                                    )}
                                  </h3>
                                  <p className="text-zinc-500 text-xs mt-0.5">
                                    Connect directly to TallyPrime to import masters and push transactions in one click.
                                  </p>
                                </div>
                              </div>

                              {getAppMode() !== 'desktop-offline' ? (
                                <div className="bg-amber-50 border border-amber-200/60 rounded-xl p-4 text-amber-800 text-xs flex items-start gap-2.5">
                                  <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
                                  <div>
                                    <p className="font-bold">Direct Tally connection is available only in Desktop Offline App.</p>
                                    <p className="mt-1 text-amber-700/90">Please download TallyGen Pro Desktop Offline client or run Electron version to enable direct connection.</p>
                                  </div>
                                </div>
                              ) : (
                                <div className="space-y-4">
                                  {/* Inputs */}
                                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                    <div>
                                      <label className="block text-[10px] font-bold text-zinc-600 uppercase tracking-wider mb-1.5">Tally Host / IP</label>
                                      <input 
                                        type="text" 
                                        value={tallyHost} 
                                        onChange={(e) => setTallyHost(e.target.value)}
                                        disabled={tallyStatus === 'Connected' || tallyStatus === 'Connecting'}
                                        placeholder="localhost" 
                                        className="w-full bg-white border border-zinc-200 rounded-xl px-3 py-2 text-sm text-zinc-900 focus:outline-none focus:ring-2 focus:ring-emerald-500 disabled:bg-zinc-100 disabled:text-zinc-500"
                                      />
                                    </div>
                                    <div>
                                      <label className="block text-[10px] font-bold text-zinc-600 uppercase tracking-wider mb-1.5">Tally Port</label>
                                      <input 
                                        type="number" 
                                        value={tallyPort} 
                                        onChange={(e) => setTallyPort(Number(e.target.value))}
                                        disabled={tallyStatus === 'Connected' || tallyStatus === 'Connecting'}
                                        placeholder="9000" 
                                        className="w-full bg-white border border-zinc-200 rounded-xl px-3 py-2 text-sm text-zinc-900 focus:outline-none focus:ring-2 focus:ring-emerald-500 disabled:bg-zinc-100 disabled:text-zinc-500"
                                      />
                                    </div>
                                  </div>

                                  {/* Connection Status Indicator */}
                                  <div className="flex flex-wrap items-center gap-4 py-2 border-t border-b border-zinc-200/60">
                                    <div className="flex items-center gap-2">
                                      <span className="text-xs font-semibold text-zinc-500">Status:</span>
                                      <span className={`inline-flex items-center gap-1.5 text-xs font-bold px-2.5 py-1 rounded-full ${
                                        tallyStatus === 'Connected' ? 'bg-emerald-100 text-emerald-800' :
                                        tallyStatus === 'Connecting' ? 'bg-blue-100 text-blue-800' :
                                        tallyStatus === 'Error' ? 'bg-rose-100 text-rose-800' :
                                        'bg-zinc-200 text-zinc-700'
                                      }`}>
                                        {tallyStatus === 'Connecting' && <Loader2 className="w-3 h-3 animate-spin" />}
                                        {tallyStatus}
                                      </span>
                                    </div>

                                    {tallyCompany && (
                                      <div className="flex items-center gap-2 text-xs">
                                        <Building className="w-3.5 h-3.5 text-zinc-400" />
                                        <span className="font-semibold text-zinc-600">Company:</span>
                                        <span className="font-bold text-emerald-700">{tallyCompany}</span>
                                      </div>
                                    )}

                                    {tallyError && (
                                      <div className="text-xs font-medium text-rose-600 w-full flex items-start gap-1 mt-1">
                                        <AlertCircle className="w-3.5 h-3.5 shrink-0 mt-0.5" />
                                        <span>{tallyError}</span>
                                      </div>
                                    )}
                                  </div>

                                  {/* Actions based on connection */}
                                  <div className="flex flex-wrap gap-2.5">
                                    {tallyStatus !== 'Connected' ? (
                                      <button
                                        type="button"
                                        onClick={handleTestConnection}
                                        disabled={tallyStatus === 'Connecting'}
                                        className="flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 disabled:bg-emerald-400 text-white font-bold text-xs rounded-xl shadow-sm transition-colors"
                                      >
                                        {tallyStatus === 'Connecting' ? 'Connecting...' : 'Connect & Test'}
                                      </button>
                                    ) : (
                                      <>
                                        <button
                                          type="button"
                                          onClick={handleFetchMasters}
                                          className="flex items-center gap-1.5 px-3.5 py-2 bg-zinc-900 hover:bg-zinc-800 text-white font-bold text-xs rounded-xl transition-colors"
                                        >
                                          <Database className="w-3.5 h-3.5" />
                                          Fetch Masters
                                        </button>
                                        
                                        <div className="flex items-center gap-2 border border-zinc-200 bg-white rounded-xl p-1.5">
                                          <div className="flex items-center gap-1 text-[11px] text-zinc-500 font-medium">
                                            <span>From:</span>
                                            <input 
                                              type="date" 
                                              value={tallyFromDate}
                                              onChange={(e) => setTallyFromDate(e.target.value)}
                                              className="bg-zinc-50 border-0 text-zinc-800 text-xs px-1.5 py-0.5 rounded font-semibold focus:ring-0"
                                            />
                                          </div>
                                          <div className="flex items-center gap-1 text-[11px] text-zinc-500 font-medium border-l border-zinc-200 pl-2">
                                            <span>To:</span>
                                            <input 
                                              type="date" 
                                              value={tallyToDate}
                                              onChange={(e) => setTallyToDate(e.target.value)}
                                              className="bg-zinc-50 border-0 text-zinc-800 text-xs px-1.5 py-0.5 rounded font-semibold focus:ring-0"
                                            />
                                          </div>
                                          <button
                                            type="button"
                                            onClick={handleFetchDaybook}
                                            className="flex items-center gap-1 px-3 py-1 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-[11px] rounded-lg transition-colors"
                                          >
                                            <FileText className="w-3 h-3" />
                                            Fetch Daybook
                                          </button>
                                        </div>

                                        <button
                                          type="button"
                                          onClick={handleDisconnect}
                                          className="flex items-center gap-1.5 px-3.5 py-2 bg-zinc-100 hover:bg-zinc-200 text-zinc-700 font-bold text-xs rounded-xl transition-colors"
                                        >
                                          <X className="w-3.5 h-3.5" />
                                          Disconnect
                                        </button>
                                      </>
                                    )}
                                  </div>

                                  {/* Helper note */}
                                  <p className="text-[11px] text-zinc-400 leading-relaxed italic">
                                    Keep TallyPrime open with the correct company loaded. Enable Tally integration/HTTP port in Tally. Default port is usually 9000.
                                  </p>
                                </div>
                              )}
                            </div>
                          </div>
                        </div>

                        {/* Recent Tally Push Logs */}
                        {getAppMode() === 'desktop-offline' && recentPushLogs.length > 0 && (
                          <div className="md:col-span-2 border-t border-zinc-200/60 pt-6 mt-4" id="tally-push-logs-section">
                            <h4 className="text-sm font-bold text-zinc-800 mb-3 flex items-center gap-2">
                              <History className="w-4 h-4 text-zinc-500" />
                              Recent Tally Direct Push Logs
                            </h4>
                            <div className="overflow-x-auto border border-zinc-200 rounded-xl bg-white max-h-60 overflow-y-auto">
                              <table className="min-w-full divide-y divide-zinc-200 text-xs">
                                <thead className="bg-zinc-50 sticky top-0">
                                  <tr>
                                    <th className="px-4 py-2.5 text-left font-bold text-zinc-600">Timestamp</th>
                                    <th className="px-4 py-2.5 text-left font-bold text-zinc-600">Company</th>
                                    <th className="px-4 py-2.5 text-left font-bold text-zinc-600">XML Type</th>
                                    <th className="px-4 py-2.5 text-left font-bold text-zinc-600">Count (Vch/Mst)</th>
                                    <th className="px-4 py-2.5 text-left font-bold text-zinc-600">Status</th>
                                    <th className="px-4 py-2.5 text-left font-bold text-zinc-600">Result / Error</th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-zinc-100 bg-white">
                                  {recentPushLogs.map((log) => (
                                    <tr key={log.id} className="hover:bg-zinc-50/50">
                                      <td className="px-4 py-2 text-zinc-500 whitespace-nowrap">
                                        {new Date(log.timestamp).toLocaleString()}
                                      </td>
                                      <td className="px-4 py-2 font-semibold text-zinc-800">
                                        {log.companyName}
                                      </td>
                                      <td className="px-4 py-2 text-zinc-600">
                                        {log.xmlType}
                                      </td>
                                      <td className="px-4 py-2 text-zinc-600 whitespace-nowrap">
                                        Vch: <span className="font-bold">{log.voucherCount}</span>, Mst: <span className="font-bold">{log.masterCount}</span>
                                      </td>
                                      <td className="px-4 py-2 whitespace-nowrap">
                                        <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full font-bold text-[10px] uppercase ${
                                          log.status === 'success' ? 'bg-green-100 text-green-800' : 'bg-rose-100 text-rose-800'
                                        }`}>
                                          {log.status}
                                        </span>
                                      </td>
                                      <td className="px-4 py-2 font-medium">
                                        {log.status === 'success' ? (
                                          <span className="text-green-600">Import successful</span>
                                        ) : (
                                          <span className="text-rose-600 break-all max-w-xs block" title={log.errorMessage}>
                                            {log.errorMessage || 'Unknown error'}
                                          </span>
                                        )}
                                      </td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        )}
                      </div>

                      {/* Summary / Progress of loading */}
                      {isContextLoading && (
                        <div className="p-4 bg-zinc-50 border border-zinc-200 rounded-xl flex items-center justify-center gap-3" id="setup-loading-indicator">
                          <Loader2 className="w-5 h-5 animate-spin text-zinc-600" />
                          <span className="text-sm text-zinc-600 font-medium">Rebuilding Tally workspace context...</span>
                        </div>
                      )}

                      {tallyContext && (
                        <motion.div 
                          initial={{ opacity: 0, y: 10 }}
                          animate={{ opacity: 1, y: 0 }}
                          className="bg-green-50/50 border border-green-200/60 p-6 rounded-2xl"
                          id="setup-summary-container"
                        >
                          <h3 className="font-bold text-green-900 text-base flex items-center gap-2 mb-4" id="setup-summary-title">
                            <CheckCircle2 className="w-5 h-5 text-green-600" />
                            Tally Context Loaded Successfully!
                          </h3>
                          
                          <div className="grid grid-cols-2 md:grid-cols-3 gap-4" id="setup-summary-metrics">
                            <div className="bg-white p-3 rounded-xl border border-green-100/50 shadow-sm">
                              <span className="block text-[10px] uppercase font-bold tracking-wider text-zinc-400 mb-0.5">Ledgers</span>
                              <span className="text-xl font-extrabold text-zinc-900">{tallyContext.ledgers?.length || 0}</span>
                            </div>
                            <div className="bg-white p-3 rounded-xl border border-green-100/50 shadow-sm">
                              <span className="block text-[10px] uppercase font-bold tracking-wider text-zinc-400 mb-0.5">Groups</span>
                              <span className="text-xl font-extrabold text-zinc-900">{tallyContext.groups?.length || 0}</span>
                            </div>
                            <div className="bg-white p-3 rounded-xl border border-green-100/50 shadow-sm">
                              <span className="block text-[10px] uppercase font-bold tracking-wider text-zinc-400 mb-0.5">Stock Items</span>
                              <span className="text-xl font-extrabold text-zinc-900">{tallyContext.stockItems?.length || 0}</span>
                            </div>
                            <div className="bg-white p-3 rounded-xl border border-green-100/50 shadow-sm">
                              <span className="block text-[10px] uppercase font-bold tracking-wider text-zinc-400 mb-0.5">Stock Groups</span>
                              <span className="text-xl font-extrabold text-zinc-900">{tallyContext.stockGroups?.length || 0}</span>
                            </div>
                            <div className="bg-white p-3 rounded-xl border border-green-100/50 shadow-sm">
                              <span className="block text-[10px] uppercase font-bold tracking-wider text-zinc-400 mb-0.5">Units</span>
                              <span className="text-xl font-extrabold text-zinc-900">{tallyContext.units?.length || 0}</span>
                            </div>
                            <div className="bg-white p-3 rounded-xl border border-green-100/50 shadow-sm">
                              <span className="block text-[10px] uppercase font-bold tracking-wider text-zinc-400 mb-0.5">Historical Mappings</span>
                              <span className="text-xl font-extrabold text-zinc-900">{tallyContext.historicalMappings?.length || 0}</span>
                            </div>
                          </div>

                          <div className="mt-6 flex justify-end" id="setup-summary-proceed">
                            <button
                              type="button"
                              onClick={() => setProceededWithContext(true)}
                              className="flex items-center gap-2 bg-green-600 hover:bg-green-700 text-white px-6 py-2.5 rounded-xl font-bold text-sm shadow-md transition-colors"
                              id="setup-proceed-btn"
                            >
                              Proceed to Voucher / Master Import
                              <ArrowRight className="w-4 h-4" />
                            </button>
                          </div>
                        </motion.div>
                      )}

                      {/* Fallback Option */}
                      <div className="border-t border-zinc-100 pt-5 flex flex-col md:flex-row md:items-center justify-between gap-4" id="setup-fallback-container">
                        <p className="text-zinc-500 text-xs leading-relaxed max-w-md flex items-start gap-2" id="setup-fallback-warning">
                          <AlertCircle className="w-4 h-4 text-amber-500 shrink-0 mt-0.5" />
                          <span>
                            Without Tally context, dropdowns and duplicate checks will use fallback/manual mode only.
                          </span>
                        </p>
                        <button
                          type="button"
                          onClick={() => setSkippedContext(true)}
                          className="px-4 py-2 hover:bg-zinc-100 border border-zinc-200 text-zinc-700 rounded-xl text-xs font-bold transition-all shrink-0"
                          id="setup-skip-btn"
                        >
                          Skip context and continue manually
                        </button>
                      </div>
                    </div>
                  ) : (
                    // Normal Configure & Upload Import Panel
                    <>
                      <h2 className="text-xl font-bold mb-6 flex items-center gap-2">
                        <Download className="w-5 h-5" />
                        Step 1: Configure & Upload
                      </h2>

                      {/* Mode Selector Tabs */}
                      <div className="flex bg-zinc-100 p-1 rounded-xl mb-6">
                        <button
                          type="button"
                          onClick={() => setImportType('Voucher')}
                          className={`flex-1 py-2 px-3 rounded-lg text-sm font-medium transition-colors ${
                            importType === 'Voucher'
                              ? 'bg-white text-zinc-900 shadow-sm'
                              : 'text-zinc-600 hover:text-zinc-900'
                          }`}
                        >
                          Voucher Import
                        </button>
                        <button
                          type="button"
                          onClick={() => setImportType('Ledger')}
                          className={`flex-1 py-2 px-3 rounded-lg text-sm font-medium transition-colors ${
                            importType !== 'Voucher'
                              ? 'bg-white text-zinc-900 shadow-sm'
                              : 'text-zinc-600 hover:text-zinc-900'
                          }`}
                        >
                          Master Import
                        </button>
                      </div>

                      {importType === 'Voucher' ? (
                        <div className="space-y-6 mb-8">
                          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                            <div className="space-y-2">
                              <label className="text-sm font-medium text-zinc-700">Voucher Type</label>
                              <select 
                                value={selectedVoucherType}
                                onChange={(e) => setSelectedVoucherType(e.target.value)}
                                className="w-full p-3 bg-zinc-50 border border-zinc-200 rounded-xl focus:ring-2 focus:ring-zinc-900 outline-none transition-all"
                              >
                                <option value="Payment">Payment</option>
                                <option value="Receipt">Receipt</option>
                                <option value="Contra">Contra</option>
                                <option value="Journal">Journal</option>
                                <option value="Sales">Sales</option>
                                <option value="Purchase">Purchase</option>
                              </select>
                            </div>

                            {!['Sales', 'Purchase'].includes(selectedVoucherType) ? (
                              <div className="space-y-2">
                                <div className="flex items-center justify-between">
                                  <label className="text-sm font-medium text-zinc-700">Bank / Cash Ledger</label>
                                  {tallyContext && (
                                    <label className="flex items-center gap-1.5 text-xs text-zinc-500 cursor-pointer hover:text-zinc-900 transition-colors select-none">
                                      <input 
                                        type="checkbox"
                                        checked={showAllLedgers}
                                        onChange={(e) => setShowAllLedgers(e.target.checked)}
                                        className="rounded border-zinc-300 text-zinc-900 focus:ring-zinc-950 w-3.5 h-3.5"
                                      />
                                      <span>Show all ledgers</span>
                                    </label>
                                  )}
                                </div>
                                <select 
                                  value={selectedBankLedger}
                                  onChange={(e) => setSelectedBankLedger(e.target.value)}
                                  className="w-full p-3 bg-zinc-50 border border-zinc-200 rounded-xl focus:ring-2 focus:ring-zinc-900 outline-none transition-all text-sm"
                                >
                                  <option value="">Select Bank / Cash Ledger...</option>
                                  {(() => {
                                    if (!tallyContext) return null;
                                    
                                    if (showAllLedgers) {
                                      return tallyContext.ledgers.map(ledger => {
                                        const groupName = tallyContext.ledgerGroupMap?.[ledger];
                                        return (
                                          <option key={ledger} value={ledger}>
                                            {ledger}{groupName ? ` — ${groupName}` : ''}
                                          </option>
                                        );
                                      });
                                    }

                                    const strictMatches: string[] = [];
                                    const possibleMatches: string[] = [];

                                    tallyContext.ledgers.forEach(ledger => {
                                      if (isStrictBankCashLedger(ledger, tallyContext)) {
                                        strictMatches.push(ledger);
                                      } else {
                                        // Check if it's a potential bank/cash ledger by name when group classification is missing or not strict
                                        const nameLower = ledger.toLowerCase();
                                        const hasBankWord = ['bank', 'cash', 'hdfc', 'icici', 'sbi', 'axis', 'kotak', 'yes bank', 'canara', 'pnb', 'union bank', 'bob', 'idfc', 'rbl'].some(kw => nameLower.includes(kw));
                                        const hasExcludeWord = ['charges', 'interest', 'loan', 'rebate', 'discount', 'gst', 'vehicle', 'fee', 'charge', 'chg', 'tax'].some(kw => nameLower.includes(kw));
                                        
                                        if (hasBankWord && !hasExcludeWord) {
                                          possibleMatches.push(ledger);
                                        }
                                      }
                                    });

                                    return (
                                      <>
                                        {strictMatches.length > 0 && (
                                          <optgroup label="Strict Matches (Recommended)">
                                            {strictMatches.map(ledger => {
                                              const groupName = tallyContext.ledgerGroupMap?.[ledger];
                                              return (
                                                <option key={ledger} value={ledger}>
                                                  {ledger}{groupName ? ` — ${groupName}` : ''}
                                                </option>
                                              );
                                            })}
                                          </optgroup>
                                        )}
                                        {possibleMatches.length > 0 && (
                                          <optgroup label="Possible Matches (Review Carefully)">
                                            {possibleMatches.map(ledger => {
                                              const groupName = tallyContext.ledgerGroupMap?.[ledger];
                                              return (
                                                <option key={ledger} value={ledger}>
                                                  {ledger}{groupName ? ` — ${groupName}` : ''}
                                                </option>
                                              );
                                            })}
                                          </optgroup>
                                        )}
                                        {strictMatches.length === 0 && possibleMatches.length === 0 && (
                                          <option value="" disabled>No bank/cash ledgers found. Try "Show all ledgers".</option>
                                        )}
                                      </>
                                    );
                                  })()}
                                </select>
                                {tallyContext ? (
                                  <p className="text-[11px] text-zinc-400 mt-1 leading-relaxed">
                                    {showAllLedgers 
                                      ? "Showing all ledgers. Please select only actual Bank / Cash ledger for voucher generation."
                                      : "Default list shows only ledgers classified under Bank Accounts, Cash-in-Hand, Bank OD/CC groups in uploaded Tally Masters. Use 'Show all ledgers' only if your bank ledger is not classified correctly in Tally."
                                    }
                                  </p>
                                ) : (
                                  <p className="text-xs text-amber-600 flex items-center gap-1 mt-1 leading-relaxed">
                                    <AlertCircle className="w-3.5 h-3.5 shrink-0" />
                                    Load Tally Masters first for Bank/Cash ledger list, or continue manually if available.
                                  </p>
                                )}
                              </div>
                            ) : (
                              <div className="space-y-2">
                                <label className="text-sm font-medium text-zinc-700">Sales/Purchase Ledgers</label>
                                <div className="p-4 bg-zinc-50 border border-zinc-200 rounded-xl">
                                  <p className="text-xs text-zinc-600 leading-relaxed font-medium">
                                    Party Ledgers and Sales/Purchase Accounts are parsed directly from your Excel spreadsheet. No Bank/Cash ledger selection is needed.
                                  </p>
                                </div>
                              </div>
                            )}
                          </div>

                          {/* Two clear import method cards */}
                          <div className="space-y-4 pt-4 border-t border-zinc-100">
                            <label className="text-sm font-medium text-zinc-700 block">Select Import Workflow</label>
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                              {/* Card 1: Template */}
                              <button
                                type="button"
                                onClick={() => setVoucherImportMethod('template')}
                                className={`p-5 rounded-xl border-2 text-left transition-all flex flex-col justify-between ${
                                  voucherImportMethod === 'template'
                                    ? 'border-zinc-900 bg-zinc-50'
                                    : 'border-zinc-200/60 hover:border-zinc-400 bg-white'
                                }`}
                              >
                                <div>
                                  <h3 className="font-bold text-zinc-900 flex items-center gap-2">
                                    <FileSpreadsheet className="w-5 h-5 text-zinc-600" />
                                    Use Tally Voucher Template
                                  </h3>
                                  <p className="text-xs text-zinc-500 mt-2 leading-relaxed">
                                    Use this when you already know voucher type, ledger name, amount, narration and reference. Requires downloading our standard format.
                                  </p>
                                </div>
                                <span className="text-xs font-bold text-zinc-900 mt-4 inline-flex items-center gap-1">
                                  Download templates below
                                  <ArrowRight className="w-3.5 h-3.5" />
                                </span>
                              </button>

                              {/* Card 2: Bank Statement */}
                              <button
                                type="button"
                                onClick={() => setVoucherImportMethod('bankStatement')}
                                className={`p-5 rounded-xl border-2 text-left transition-all flex flex-col justify-between ${
                                  voucherImportMethod === 'bankStatement'
                                    ? 'border-zinc-900 bg-zinc-50'
                                    : 'border-zinc-200/60 hover:border-zinc-400 bg-white'
                                }`}
                              >
                                <div>
                                  <h3 className="font-bold text-zinc-900 flex items-center gap-2">
                                    <Building className="w-5 h-5 text-zinc-600" />
                                    Upload Bank Statement
                                  </h3>
                                  <p className="text-xs text-zinc-500 mt-2 leading-relaxed">
                                    Use this when you want to convert bank debit/credit transactions into Payment and Receipt vouchers automatically.
                                  </p>
                                </div>
                                <span className="text-xs font-bold text-zinc-900 mt-4 inline-flex items-center gap-1">
                                  Local deterministic parsing
                                  <ArrowRight className="w-3.5 h-3.5" />
                                </span>
                              </button>
                            </div>

                            {/* Render Template Download Quick Buttons */}
                            {voucherImportMethod === 'template' && (
                              <motion.div
                                initial={{ opacity: 0, y: -10 }}
                                animate={{ opacity: 1, y: 0 }}
                                className="p-4 bg-zinc-50 rounded-xl border border-zinc-200/60 mt-3 flex flex-wrap gap-2 items-center"
                              >
                                <span className="text-xs font-bold text-zinc-600 block mr-2">Download Template:</span>
                                {['Payment', 'Receipt', 'Contra', 'Journal', 'Sales Itemwise', 'Sales Voucherwise', 'Purchase Itemwise', 'Purchase Voucherwise'].map(type => (
                                  <button
                                    key={type}
                                    type="button"
                                    onClick={() => downloadTemplate(type)}
                                    className="py-1.5 px-3 bg-white border border-zinc-200 text-zinc-700 hover:bg-zinc-50 rounded-lg text-xs font-semibold transition-colors flex items-center gap-1"
                                  >
                                    <Download className="w-3.5 h-3.5 text-zinc-500" />
                                    {type}
                                  </button>
                                ))}
                              </motion.div>
                            )}

                            {/* Render Voucher Mode Selector ONLY if method is bankStatement */}
                            {voucherImportMethod === 'bankStatement' && (
                              <motion.div
                                initial={{ opacity: 0, y: -10 }}
                                animate={{ opacity: 1, y: 0 }}
                                className="bg-zinc-50 p-4 rounded-xl border border-zinc-200/60 mt-4 grid grid-cols-1 md:grid-cols-3 gap-4"
                              >
                                <div className="md:col-span-3">
                                  <label className="text-xs font-bold text-zinc-700 block mb-1">Bank Statement Mapping Mode</label>
                                  <p className="text-[11px] text-zinc-500 mb-3">
                                    Control how statement rows are classified into Payment and Receipt vouchers.
                                  </p>
                                </div>
                                <button
                                  type="button"
                                  onClick={() => setVoucherMode('auto')}
                                  className={`py-2 px-3 rounded-lg text-xs font-semibold border transition-all ${
                                    voucherMode === 'auto'
                                      ? 'bg-zinc-900 border-zinc-950 text-white'
                                      : 'bg-white border-zinc-200 text-zinc-700 hover:bg-zinc-100'
                                  }`}
                                >
                                  Auto from Debit/Credit
                                </button>
                                <button
                                  type="button"
                                  onClick={() => setVoucherMode('payment')}
                                  className={`py-2 px-3 rounded-lg text-xs font-semibold border transition-all ${
                                    voucherMode === 'payment'
                                      ? 'bg-zinc-900 border-zinc-950 text-white'
                                      : 'bg-white border-zinc-200 text-zinc-700 hover:bg-zinc-100'
                                  }`}
                                >
                                  Payment Only (Debits)
                                </button>
                                <button
                                  type="button"
                                  onClick={() => setVoucherMode('receipt')}
                                  className={`py-2 px-3 rounded-lg text-xs font-semibold border transition-all ${
                                    voucherMode === 'receipt'
                                      ? 'bg-zinc-900 border-zinc-950 text-white'
                                      : 'bg-white border-zinc-200 text-zinc-700 hover:bg-zinc-100'
                                  }`}
                                >
                                  Receipt Only (Credits)
                                </button>
                              </motion.div>
                            )}
                          </div>
                        </div>
                      ) : (
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
                          <div className="space-y-2">
                            <label className="text-sm font-medium text-zinc-700">Master Type</label>
                            <select 
                              value={importType}
                              onChange={(e) => setImportType(e.target.value as any)}
                              className="w-full p-3 bg-zinc-50 border border-zinc-200 rounded-xl focus:ring-2 focus:ring-zinc-900 outline-none transition-all"
                            >
                              <option value="Ledger">Ledger Master</option>
                              <option value="StockItem">Stock Item Master</option>
                              <option value="StockGroup">Stock Group Master</option>
                              <option value="Unit">Unit Master</option>
                            </select>
                          </div>

                          <div className="space-y-2 flex flex-col justify-end">
                            <button
                              type="button"
                              onClick={() => downloadMasterTemplate(importType)}
                              className="w-full py-3 px-4 rounded-xl font-medium border border-zinc-200 hover:bg-zinc-50 transition-colors flex items-center justify-center gap-2 bg-zinc-50"
                            >
                              <Download className="w-4 h-4" />
                              Download Excel Template
                            </button>
                          </div>
                        </div>
                      )}
                      
                      <div className={`border-2 border-dashed rounded-xl p-12 text-center transition-colors relative group ${
                        importType === 'Voucher' && !['Sales', 'Purchase'].includes(selectedVoucherType) && !selectedBankLedger 
                          ? 'border-zinc-100 bg-zinc-50/50 cursor-not-allowed' 
                          : 'border-zinc-200 hover:border-zinc-400 cursor-pointer'
                      }`}>
                        <input 
                           type="file" 
                           accept=".xlsx, .xls, .csv, .pdf" 
                           onChange={handleFileUpload}
                           className="absolute inset-0 opacity-0 cursor-pointer disabled:cursor-not-allowed"
                           disabled={isProcessing || (importType === 'Voucher' && !['Sales', 'Purchase'].includes(selectedVoucherType) && !selectedBankLedger)}
                        />
                        <div className="space-y-4">
                          <div className={`w-12 h-12 rounded-full flex items-center justify-center mx-auto transition-colors ${
                            importType === 'Voucher' && !['Sales', 'Purchase'].includes(selectedVoucherType) && !selectedBankLedger 
                              ? 'bg-zinc-100 text-zinc-300' 
                              : 'bg-zinc-100 text-zinc-600 group-hover:bg-zinc-200'
                          }`}>
                            {isProcessing ? (
                              <Loader2 className="w-6 h-6 animate-spin" />
                            ) : (
                              <FileSpreadsheet className="w-6 h-6" />
                            )}
                          </div>
                          <div>
                            {importType !== 'Voucher' || selectedBankLedger ? (
                              <>
                                <p className="font-medium">Click or drag Excel/CSV/PDF to upload</p>
                                <p className="text-sm text-zinc-500">
                                  {importType === 'Voucher' 
                                    ? 'Upload bank statement to convert to vouchers'
                                    : `Upload Master data file to convert to Tally ${importType} Masters`
                                  }
                                </p>
                              </>
                            ) : (
                              <p className="font-medium text-zinc-400">Please select a Bank Ledger first</p>
                            )}
                          </div>
                        </div>
                      </div>

                      {skippedContext && (
                        <div className="mt-4 p-4 bg-amber-50 border border-amber-200 rounded-xl flex items-start gap-3 text-amber-800 text-xs">
                          <AlertCircle className="w-4 h-4 shrink-0 mt-0.5 text-amber-600" />
                          <div>
                            <span className="font-semibold">Fallback mode active</span>: You are continuing without Tally context. Dropdowns, ledger auto-completions, and duplicate checks will operate in manual fallback mode.
                          </div>
                        </div>
                      )}
                    </>
                  )}

                  {error && (
                    <div className={error.includes("reset completed") || error.includes("fallback reset")
                      ? "mt-4 p-4 bg-blue-50 border border-blue-200 rounded-xl flex items-center gap-3 text-blue-700"
                      : "mt-4 p-4 bg-red-50 border border-red-100 rounded-xl flex items-center gap-3 text-red-700"}>
                      <AlertCircle className="w-5 h-5 shrink-0" />
                      <p className="text-sm">{error}</p>
                    </div>
                  )}
                </motion.section>
              )}

              {currentStep === 'bank-statement-detection-review' && (
                <motion.section 
                  key="bank-statement-detection-review"
                  initial={{ opacity: 0, x: -20 }}
                  animate={{ opacity: 1, x: 0 }}
                  exit={{ opacity: 0, x: 20 }}
                  className="bg-white p-6 rounded-2xl border border-zinc-200 shadow-sm space-y-6"
                >
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-zinc-100 pb-5">
                    <div>
                      <h2 className="text-xl font-bold flex items-center gap-2 text-zinc-900">
                        <Sparkles className="w-5 h-5 text-zinc-800 animate-pulse" />
                        Bank Statement Detection Review
                      </h2>
                      <p className="text-xs text-zinc-500 mt-1 leading-relaxed">
                        Verify and customize the automatically detected sheet structure, row boundaries, and column mappings.
                      </p>
                    </div>
                    <div className="text-right">
                      <span className="text-xs bg-zinc-100 px-3 py-1.5 rounded-lg text-zinc-700 font-semibold block md:inline-block">
                        {pendingFileName}
                      </span>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                    {/* Left: Row Boundaries & Modes (1/3 width) */}
                    <div className="space-y-4 lg:col-span-1">
                      <div className="bg-zinc-50/50 p-4 rounded-xl border border-zinc-200/50 space-y-4">
                        <h3 className="text-sm font-bold text-zinc-800 flex items-center gap-2 pb-2 border-b border-zinc-200/60">
                          <FileText className="w-4 h-4 text-zinc-500" />
                          Row Settings
                        </h3>

                        <div>
                          <label className="block text-xs font-semibold text-zinc-600 mb-1">
                            Header Row (1-Indexed)
                          </label>
                          <input 
                            type="number"
                            min="1"
                            max={rawGrid.length}
                            value={headerRowIdx + 1}
                            onChange={(e) => {
                              const val = parseInt(e.target.value) - 1;
                              if (!isNaN(val) && val >= 0 && val < rawGrid.length) {
                                setHeaderRowIdx(val);
                              }
                            }}
                            className="w-full p-2.5 bg-white border border-zinc-200 rounded-lg text-xs outline-none focus:ring-1 focus:ring-zinc-950 font-medium"
                          />
                          <span className="text-[10px] text-zinc-400 mt-1 block">
                            Used to find column names and auto-detect mapping.
                          </span>
                        </div>

                        <div>
                          <label className="block text-xs font-semibold text-zinc-600 mb-1">
                            Data Start Row (1-Indexed)
                          </label>
                          <input 
                            type="number"
                            min="1"
                            max={rawGrid.length}
                            value={dataStartRowIdx + 1}
                            onChange={(e) => {
                              const val = parseInt(e.target.value) - 1;
                              if (!isNaN(val) && val >= 0 && val < rawGrid.length) {
                                setDataStartRowIdx(val);
                              }
                            }}
                            className="w-full p-2.5 bg-white border border-zinc-200 rounded-lg text-xs outline-none focus:ring-1 focus:ring-zinc-950 font-medium"
                          />
                          <span className="text-[10px] text-zinc-400 mt-1 block">
                            Where actual transactions begin (skips opening balance).
                          </span>
                        </div>

                        <div>
                          <label className="block text-xs font-semibold text-zinc-600 mb-1">
                            Data End Row (1-Indexed)
                          </label>
                          <input 
                            type="number"
                            min="1"
                            max={rawGrid.length}
                            value={dataEndRowIdx + 1}
                            onChange={(e) => {
                              const val = parseInt(e.target.value) - 1;
                              if (!isNaN(val) && val >= 0 && val < rawGrid.length) {
                                setDataEndRowIdx(val);
                              }
                            }}
                            className="w-full p-2.5 bg-white border border-zinc-200 rounded-lg text-xs outline-none focus:ring-1 focus:ring-zinc-950 font-medium"
                          />
                          <span className="text-[10px] text-zinc-400 mt-1 block">
                            Where transactions end (ignores summaries or total rows).
                          </span>
                        </div>

                        <div>
                          <label className="block text-xs font-semibold text-zinc-600 mb-1">
                            Single-Amount Mode Default
                          </label>
                          <select
                            value={detectedVoucherMode}
                            onChange={(e) => setDetectedVoucherMode(e.target.value as any)}
                            className="w-full p-2.5 bg-white border border-zinc-200 rounded-lg text-xs outline-none focus:ring-1 focus:ring-zinc-950 font-medium"
                          >
                            <option value="auto">Auto-detect from Debit/Credit</option>
                            <option value="payment">All Rows as Payment (Debit)</option>
                            <option value="receipt">All Rows as Receipt (Credit)</option>
                          </select>
                          <span className="text-[10px] text-zinc-400 mt-1 block">
                            Fallback for statement tables with a single amount column.
                          </span>
                        </div>
                      </div>
                    </div>

                    {/* Right: Column Mapping Dropdowns (2/3 width) */}
                    <div className="space-y-4 lg:col-span-2">
                      <div className="bg-zinc-50/50 p-4 rounded-xl border border-zinc-200/50">
                        <h3 className="text-sm font-bold text-zinc-800 flex items-center gap-2 pb-2 border-b border-zinc-200/60 mb-4">
                          <Database className="w-4 h-4 text-zinc-500" />
                          Column Mappings
                        </h3>

                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                          {/* Date Column Mapping */}
                          <div>
                            <label className="block text-xs font-semibold text-zinc-700 mb-1">
                              Date Column <span className="text-red-500">*</span>
                            </label>
                            <select
                              value={columnMappings.date ?? ''}
                              onChange={(e) => {
                                const val = e.target.value === '' ? null : parseInt(e.target.value);
                                setColumnMappings(prev => ({ ...prev, date: val }));
                              }}
                              className="w-full p-2.5 bg-white border border-zinc-200 rounded-lg text-xs outline-none focus:ring-1 focus:ring-zinc-950 font-medium"
                            >
                              <option value="">-- Unmapped --</option>
                              {Array.from({ length: rawGrid.reduce((max, r) => Math.max(max, r.length), 0) }, (_, c) => {
                                const colLetter = String.fromCharCode(65 + (c % 26)) + (c >= 26 ? Math.floor(c / 26) : '');
                                const val = rawGrid[headerRowIdx]?.[c];
                                const label = val !== undefined && val !== null ? String(val).trim() : '';
                                return <option key={c} value={c}>{colLetter} - {label || '(Empty cell)'}</option>;
                              })}
                            </select>
                          </div>

                          {/* Narration Column Mapping */}
                          <div>
                            <label className="block text-xs font-semibold text-zinc-700 mb-1">
                              Narration Column <span className="text-red-500">*</span>
                            </label>
                            <select
                              value={columnMappings.narration ?? ''}
                              onChange={(e) => {
                                const val = e.target.value === '' ? null : parseInt(e.target.value);
                                setColumnMappings(prev => ({ ...prev, narration: val }));
                              }}
                              className="w-full p-2.5 bg-white border border-zinc-200 rounded-lg text-xs outline-none focus:ring-1 focus:ring-zinc-950 font-medium"
                            >
                              <option value="">-- Unmapped --</option>
                              {Array.from({ length: rawGrid.reduce((max, r) => Math.max(max, r.length), 0) }, (_, c) => {
                                const colLetter = String.fromCharCode(65 + (c % 26)) + (c >= 26 ? Math.floor(c / 26) : '');
                                const val = rawGrid[headerRowIdx]?.[c];
                                const label = val !== undefined && val !== null ? String(val).trim() : '';
                                return <option key={c} value={c}>{colLetter} - {label || '(Empty cell)'}</option>;
                              })}
                            </select>
                          </div>

                          {/* Debit Column Mapping */}
                          <div>
                            <label className="block text-xs font-semibold text-zinc-700 mb-1">
                              Debit / Withdrawal Column (Optional)
                            </label>
                            <select
                              value={columnMappings.debit ?? ''}
                              onChange={(e) => {
                                const val = e.target.value === '' ? null : parseInt(e.target.value);
                                setColumnMappings(prev => ({ ...prev, debit: val }));
                              }}
                              className="w-full p-2.5 bg-white border border-zinc-200 rounded-lg text-xs outline-none focus:ring-1 focus:ring-zinc-950 font-medium"
                            >
                              <option value="">-- Unmapped --</option>
                              {Array.from({ length: rawGrid.reduce((max, r) => Math.max(max, r.length), 0) }, (_, c) => {
                                const colLetter = String.fromCharCode(65 + (c % 26)) + (c >= 26 ? Math.floor(c / 26) : '');
                                const val = rawGrid[headerRowIdx]?.[c];
                                const label = val !== undefined && val !== null ? String(val).trim() : '';
                                return <option key={c} value={c}>{colLetter} - {label || '(Empty cell)'}</option>;
                              })}
                            </select>
                          </div>

                          {/* Credit Column Mapping */}
                          <div>
                            <label className="block text-xs font-semibold text-zinc-700 mb-1">
                              Credit / Deposit Column (Optional)
                            </label>
                            <select
                              value={columnMappings.credit ?? ''}
                              onChange={(e) => {
                                const val = e.target.value === '' ? null : parseInt(e.target.value);
                                setColumnMappings(prev => ({ ...prev, credit: val }));
                              }}
                              className="w-full p-2.5 bg-white border border-zinc-200 rounded-lg text-xs outline-none focus:ring-1 focus:ring-zinc-950 font-medium"
                            >
                              <option value="">-- Unmapped --</option>
                              {Array.from({ length: rawGrid.reduce((max, r) => Math.max(max, r.length), 0) }, (_, c) => {
                                const colLetter = String.fromCharCode(65 + (c % 26)) + (c >= 26 ? Math.floor(c / 26) : '');
                                const val = rawGrid[headerRowIdx]?.[c];
                                const label = val !== undefined && val !== null ? String(val).trim() : '';
                                return <option key={c} value={c}>{colLetter} - {label || '(Empty cell)'}</option>;
                              })}
                            </select>
                          </div>

                          {/* Amount Column Mapping */}
                          <div>
                            <label className="block text-xs font-semibold text-zinc-700 mb-1">
                              Single Amount Column (Optional)
                            </label>
                            <select
                              value={columnMappings.amount ?? ''}
                              onChange={(e) => {
                                const val = e.target.value === '' ? null : parseInt(e.target.value);
                                setColumnMappings(prev => ({ ...prev, amount: val }));
                              }}
                              className="w-full p-2.5 bg-white border border-zinc-200 rounded-lg text-xs outline-none focus:ring-1 focus:ring-zinc-950 font-medium"
                            >
                              <option value="">-- Unmapped --</option>
                              {Array.from({ length: rawGrid.reduce((max, r) => Math.max(max, r.length), 0) }, (_, c) => {
                                const colLetter = String.fromCharCode(65 + (c % 26)) + (c >= 26 ? Math.floor(c / 26) : '');
                                const val = rawGrid[headerRowIdx]?.[c];
                                const label = val !== undefined && val !== null ? String(val).trim() : '';
                                return <option key={c} value={c}>{colLetter} - {label || '(Empty cell)'}</option>;
                              })}
                            </select>
                          </div>

                          {/* DrCr Indicator Mapping */}
                          <div>
                            <label className="block text-xs font-semibold text-zinc-700 mb-1">
                              Dr / Cr Indicator (Optional)
                            </label>
                            <select
                              value={columnMappings.drCr ?? ''}
                              onChange={(e) => {
                                const val = e.target.value === '' ? null : parseInt(e.target.value);
                                setColumnMappings(prev => ({ ...prev, drCr: val }));
                              }}
                              className="w-full p-2.5 bg-white border border-zinc-200 rounded-lg text-xs outline-none focus:ring-1 focus:ring-zinc-950 font-medium"
                            >
                              <option value="">-- Unmapped --</option>
                              {Array.from({ length: rawGrid.reduce((max, r) => Math.max(max, r.length), 0) }, (_, c) => {
                                const colLetter = String.fromCharCode(65 + (c % 26)) + (c >= 26 ? Math.floor(c / 26) : '');
                                const val = rawGrid[headerRowIdx]?.[c];
                                const label = val !== undefined && val !== null ? String(val).trim() : '';
                                return <option key={c} value={c}>{colLetter} - {label || '(Empty cell)'}</option>;
                              })}
                            </select>
                          </div>

                          {/* Reference Mapping */}
                          <div>
                            <label className="block text-xs font-semibold text-zinc-700 mb-1">
                              Reference / Cheque No (Optional)
                            </label>
                            <select
                              value={columnMappings.reference ?? ''}
                              onChange={(e) => {
                                const val = e.target.value === '' ? null : parseInt(e.target.value);
                                setColumnMappings(prev => ({ ...prev, reference: val }));
                              }}
                              className="w-full p-2.5 bg-white border border-zinc-200 rounded-lg text-xs outline-none focus:ring-1 focus:ring-zinc-950 font-medium"
                            >
                              <option value="">-- Unmapped --</option>
                              {Array.from({ length: rawGrid.reduce((max, r) => Math.max(max, r.length), 0) }, (_, c) => {
                                const colLetter = String.fromCharCode(65 + (c % 26)) + (c >= 26 ? Math.floor(c / 26) : '');
                                const val = rawGrid[headerRowIdx]?.[c];
                                const label = val !== undefined && val !== null ? String(val).trim() : '';
                                return <option key={c} value={c}>{colLetter} - {label || '(Empty cell)'}</option>;
                              })}
                            </select>
                          </div>

                          {/* Balance Mapping */}
                          <div>
                            <label className="block text-xs font-semibold text-zinc-700 mb-1">
                              Balance Column (Optional, Ignored)
                            </label>
                            <select
                              value={columnMappings.balance ?? ''}
                              onChange={(e) => {
                                const val = e.target.value === '' ? null : parseInt(e.target.value);
                                setColumnMappings(prev => ({ ...prev, balance: val }));
                              }}
                              className="w-full p-2.5 bg-white border border-zinc-200 rounded-lg text-xs outline-none focus:ring-1 focus:ring-zinc-950 font-medium"
                            >
                              <option value="">-- Unmapped --</option>
                              {Array.from({ length: rawGrid.reduce((max, r) => Math.max(max, r.length), 0) }, (_, c) => {
                                const colLetter = String.fromCharCode(65 + (c % 26)) + (c >= 26 ? Math.floor(c / 26) : '');
                                const val = rawGrid[headerRowIdx]?.[c];
                                const label = val !== undefined && val !== null ? String(val).trim() : '';
                                return <option key={c} value={c}>{colLetter} - {label || '(Empty cell)'}</option>;
                              })}
                            </select>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Raw Data Preview */}
                  <div className="space-y-2">
                    <h3 className="text-xs font-bold text-zinc-700 uppercase tracking-wider">
                      Detected Transactions Preview (First 10 Rows)
                    </h3>
                    <div className="border border-zinc-200 rounded-xl overflow-hidden bg-white shadow-sm max-h-[300px] overflow-y-auto">
                      <table className="w-full text-xs text-left border-collapse">
                        <thead className="bg-zinc-50 text-zinc-600 font-bold border-b border-zinc-200 sticky top-0">
                          <tr>
                            <th className="p-2 w-12 text-center">Row</th>
                            <th className="p-2">Date</th>
                            <th className="p-2 max-w-xs">Narration</th>
                            <th className="p-2 text-right">Debit</th>
                            <th className="p-2 text-right">Credit</th>
                            <th className="p-2 text-right">Amount</th>
                            <th className="p-2 text-center">Dr/Cr</th>
                            <th className="p-2">Reference</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-zinc-100 font-medium">
                          {(() => {
                            const pStart = Math.max(0, dataStartRowIdx);
                            const pEnd = Math.min(rawGrid.length - 1, dataStartRowIdx + 9);
                            const previewRowsList = [];
                            
                            const dCol = columnMappings.date;
                            const nCol = columnMappings.narration;
                            const dbCol = columnMappings.debit;
                            const crCol = columnMappings.credit;
                            const aCol = columnMappings.amount;
                            const dcCol = columnMappings.drCr;
                            const rCol = columnMappings.reference;

                            for (let r = pStart; r <= pEnd; r++) {
                              const row = rawGrid[r];
                              if (!row) continue;
                              
                              let dateStr = '';
                              let dVal = dCol !== null ? row[dCol] : '';
                              if (dVal instanceof Date) {
                                dateStr = `${dVal.getDate()}/${dVal.getMonth()+1}/${dVal.getFullYear()}`;
                              } else if (dVal !== undefined && dVal !== null) {
                                dateStr = String(dVal);
                              }

                              previewRowsList.push(
                                <tr key={r} className="hover:bg-zinc-50/50">
                                  <td className="p-2 text-center font-mono text-zinc-500">{r + 1}</td>
                                  <td className="p-2 font-semibold text-zinc-800">{dateStr}</td>
                                  <td className="p-2 max-w-xs truncate" title={nCol !== null ? String(row[nCol] || '') : ''}>
                                    {nCol !== null ? String(row[nCol] || '') : <span className="text-zinc-400">Not mapped</span>}
                                  </td>
                                  <td className="p-2 text-right font-mono text-rose-600">
                                    {dbCol !== null && row[dbCol] !== undefined && row[dbCol] !== null ? String(row[dbCol]) : '-'}
                                  </td>
                                  <td className="p-2 text-right font-mono text-emerald-600">
                                    {crCol !== null && row[crCol] !== undefined && row[crCol] !== null ? String(row[crCol]) : '-'}
                                  </td>
                                  <td className="p-2 text-right font-mono text-zinc-900">
                                    {aCol !== null && row[aCol] !== undefined && row[aCol] !== null ? String(row[aCol]) : '-'}
                                  </td>
                                  <td className="p-2 text-center font-mono text-zinc-500">
                                    {dcCol !== null && row[dcCol] !== undefined && row[dcCol] !== null ? String(row[dcCol]) : '-'}
                                  </td>
                                  <td className="p-2 truncate max-w-[120px]">
                                    {rCol !== null && row[rCol] !== undefined && row[rCol] !== null ? String(row[rCol]) : '-'}
                                  </td>
                                </tr>
                              );
                            }
                            
                            if (previewRowsList.length === 0) {
                              return (
                                <tr>
                                  <td colSpan={8} className="p-4 text-center text-zinc-400">
                                    No data rows in range. Please adjust Start and End row values.
                                  </td>
                                </tr>
                              );
                            }
                            return previewRowsList;
                          })()}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {/* Actions Row */}
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 pt-4 border-t border-zinc-100">
                    <button
                      type="button"
                      onClick={() => {
                        setRawGrid([]);
                        setCurrentStep('upload');
                      }}
                      className="px-5 py-3 hover:bg-zinc-100 border border-zinc-200 text-zinc-700 rounded-xl text-xs font-bold transition-all flex items-center gap-2 self-start md:self-auto"
                    >
                      <RotateCcw className="w-4 h-4" />
                      Back to Upload
                    </button>
                    <div className="flex flex-wrap items-center gap-3 justify-end shrink-0">
                      <button
                        type="button"
                        onClick={() => {
                          const detected = autoDetectBankStatement(rawGrid);
                          setHeaderRowIdx(detected.headerRowIdx);
                          setDataStartRowIdx(detected.dataStartRowIdx);
                          setDataEndRowIdx(detected.dataEndRowIdx);
                          setColumnMappings(detected.mappings);
                        }}
                        className="px-4 py-3 bg-zinc-100 hover:bg-zinc-200 text-zinc-800 rounded-xl text-xs font-bold transition-all"
                      >
                        Reset to Autodetect
                      </button>
                      <button
                        type="button"
                        onClick={applyDetectionAndMapLedgers}
                        disabled={isProcessing}
                        className="px-5 py-3 bg-zinc-900 hover:bg-zinc-850 text-white rounded-xl text-xs font-bold transition-all flex items-center gap-2 shadow-sm"
                      >
                        {isProcessing ? (
                          <>
                            <Loader2 className="w-4 h-4 animate-spin" />
                            Mapping Ledgers...
                          </>
                        ) : (
                          <>
                            <Sparkles className="w-4 h-4 text-amber-300" />
                            Use Detected Mapping / Continue
                          </>
                        )}
                      </button>
                    </div>
                  </div>
                </motion.section>
              )}

              {currentStep === 'bank-statement-review' && (
                <motion.section 
                  key="bank-statement-review"
                  initial={{ opacity: 0, x: -20 }}
                  animate={{ opacity: 1, x: 0 }}
                  exit={{ opacity: 0, x: 20 }}
                  className="bg-white p-6 rounded-2xl border border-zinc-200 shadow-sm space-y-6"
                >
                  {/* Top Header */}
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-zinc-100 pb-5">
                    <div>
                      <h2 className="text-xl font-bold flex items-center gap-2 text-zinc-900">
                        <Building className="w-5 h-5 text-zinc-800" />
                        Step 2: Bank Statement Ledger Review
                      </h2>
                      <p className="text-xs text-zinc-500 mt-1 leading-relaxed">
                        Verify transactions, toggle payment/receipt, select final ledger accounts, and exclude rows before exporting to Tally XML.
                      </p>
                    </div>
                    <div className="text-right flex flex-col items-end gap-1">
                      <span className="text-xs bg-zinc-100 px-3 py-1.5 rounded-lg text-zinc-700 font-semibold">
                        {pendingFileName}
                      </span>
                    </div>
                  </div>

                  {/* Summary Indicators */}
                  <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                    <div className="bg-zinc-50 p-4 rounded-xl border border-zinc-200/50">
                      <p className="text-[10px] uppercase font-bold tracking-wider text-zinc-400">Total Rows</p>
                      <p className="text-xl font-bold text-zinc-800 mt-1">{bankStatementRows.length}</p>
                    </div>
                    <div className="bg-emerald-50/50 p-4 rounded-xl border border-emerald-100">
                      <p className="text-[10px] uppercase font-bold tracking-wider text-emerald-600">Valid & Included</p>
                      <p className="text-xl font-bold text-emerald-700 mt-1">
                        {bankStatementRows.filter(r => r.status === 'valid' && !r.excluded).length}
                      </p>
                    </div>
                    <div className="bg-amber-50/50 p-4 rounded-xl border border-amber-100">
                      <p className="text-[10px] uppercase font-bold tracking-wider text-amber-600">Warnings / Excluded</p>
                      <p className="text-xl font-bold text-amber-700 mt-1">
                        {bankStatementRows.filter(r => r.status === 'warning' || r.excluded).length}
                      </p>
                    </div>
                    <div className="bg-rose-50/50 p-4 rounded-xl border border-rose-100">
                      <p className="text-[10px] uppercase font-bold tracking-wider text-rose-600">Invalid / Needs Review</p>
                      <p className="text-xl font-bold text-rose-700 mt-1">
                        {bankStatementRows.filter(r => r.status === 'invalid' || (!r.userLedger || !r.userLedger.trim())).length}
                      </p>
                    </div>
                  </div>

                  {/* Active Bank Ledger Selection & Quick Config */}
                  <div className="bg-zinc-50 border border-zinc-200/60 p-4 rounded-xl flex flex-col md:flex-row md:items-center justify-between gap-4">
                    <div className="flex items-center gap-3">
                      <div className="bg-zinc-900 text-white p-2 rounded-lg shrink-0">
                        <Building className="w-5 h-5" />
                      </div>
                      <div>
                        <p className="text-xs font-semibold text-zinc-500 uppercase tracking-wider">Active Bank / Cash Ledger</p>
                        {tallyContext ? (
                          <select
                            value={selectedBankLedger}
                            onChange={(e) => setSelectedBankLedger(e.target.value)}
                            className="bg-transparent text-sm font-bold text-zinc-800 focus:outline-none border-b border-dashed border-zinc-400 pb-0.5 cursor-pointer mt-0.5"
                          >
                            <option value="">-- Select Bank/Cash Ledger --</option>
                            {tallyContext.ledgers.map(l => (
                              <option key={l} value={l}>{l}</option>
                            ))}
                          </select>
                        ) : (
                          <input
                            type="text"
                            value={selectedBankLedger}
                            onChange={(e) => setSelectedBankLedger(e.target.value)}
                            placeholder="Type Bank/Cash Ledger name"
                            className="bg-transparent text-sm font-bold text-zinc-800 focus:outline-none border-b border-dashed border-zinc-400 pb-0.5 placeholder-zinc-400 w-64 mt-0.5"
                          />
                        )}
                      </div>
                    </div>
                    
                    <div className="flex flex-wrap items-center gap-4 text-xs font-medium text-zinc-600">
                      <label className="flex items-center gap-2 cursor-pointer select-none">
                        <input
                          type="checkbox"
                          checked={bankShowAttentionRequiredFirst}
                          onChange={(e) => setBankShowAttentionRequiredFirst(e.target.checked)}
                          className="rounded border-zinc-300 text-zinc-900 focus:ring-zinc-950 w-4 h-4 cursor-pointer"
                        />
                        Show attention required first
                      </label>
                      <label className="flex items-center gap-2 cursor-pointer select-none">
                        <input
                          type="checkbox"
                          checked={bankCompactRowHeight}
                          onChange={(e) => setBankCompactRowHeight(e.target.checked)}
                          className="rounded border-zinc-300 text-zinc-900 focus:ring-zinc-950 w-4 h-4 cursor-pointer"
                        />
                        Compact Row Height
                      </label>
                    </div>
                  </div>

                  {/* Search and Filters Toolbar */}
                  <div className="space-y-3">
                    <div className="flex flex-col md:flex-row gap-3">
                      {/* Search Input */}
                      <div className="relative flex-1">
                        <Search className="absolute left-3 top-2.5 h-4 w-4 text-zinc-400" />
                        <input
                          type="text"
                          placeholder="Search rows by date, narration, reference, ledger, voucher type, or amount..."
                          value={bankSearchTerm}
                          onChange={(e) => setBankSearchTerm(e.target.value)}
                          className="w-full pl-10 pr-4 py-2 border border-zinc-200 rounded-xl text-xs outline-none focus:border-zinc-900 focus:ring-1 focus:ring-zinc-900"
                        />
                        {bankSearchTerm && (
                          <button
                            onClick={() => setBankSearchTerm('')}
                            className="absolute right-3 top-2 text-zinc-400 hover:text-zinc-600 text-xs font-semibold"
                          >
                            Clear
                          </button>
                        )}
                      </div>
                    </div>

                    {/* Filter Chips list */}
                    {(() => {
                      const totalCount = bankStatementRows.length;
                      const validCount = bankStatementRows.filter(r => r.status === 'valid' && !r.excluded).length;
                      const invalidCount = bankStatementRows.filter(r => r.status === 'invalid').length;
                      const warningCount = bankStatementRows.filter(r => r.status === 'warning').length;
                      const blankCount = bankStatementRows.filter(r => !r.userLedger || !r.userLedger.trim()).length;
                      const suspenseCount = bankStatementRows.filter(r => r.userLedger && /suspense/i.test(r.userLedger)).length;
                      const missingCount = tallyContext ? bankStatementRows.filter(r => r.userLedger && !tallyContext.ledgers.includes(r.userLedger)).length : 0;
                      const lowConfCount = bankStatementRows.filter(r => r.confidence < 90).length;
                      const paymentCount = bankStatementRows.filter(r => r.detectedVoucherType === 'Payment').length;
                      const receiptCount = bankStatementRows.filter(r => r.detectedVoucherType === 'Receipt').length;

                      const chips = [
                        { label: 'All', value: 'All', count: totalCount },
                        { label: 'Valid', value: 'Valid', count: validCount },
                        { label: 'Invalid', value: 'Invalid', count: invalidCount },
                        { label: 'Warnings', value: 'Warnings', count: warningCount },
                        { label: 'Blank Ledger', value: 'Blank Ledger', count: blankCount },
                        { label: 'Suspense Account', value: 'Suspense', count: suspenseCount },
                        ...(tallyContext ? [{ label: 'Missing Master', value: 'Missing Master', count: missingCount }] : []),
                        { label: 'Low Confidence (<90%)', value: 'Low Confidence', count: lowConfCount },
                        { label: 'Payment (Dr)', value: 'Payment', count: paymentCount },
                        { label: 'Receipt (Cr)', value: 'Receipt', count: receiptCount },
                      ];

                      return (
                        <div className="flex flex-wrap gap-2 pt-1">
                          {chips.map(chip => (
                            <button
                              key={chip.value}
                              type="button"
                              onClick={() => setBankFilterType(chip.value)}
                              className={`px-3 py-1.5 rounded-full text-xs font-medium border transition-all flex items-center gap-1.5 ${
                                bankFilterType === chip.value
                                  ? 'bg-zinc-950 text-white border-zinc-950'
                                  : 'bg-zinc-50 text-zinc-600 border-zinc-200 hover:bg-zinc-100'
                              }`}
                            >
                              <span>{chip.label}</span>
                              <span className={`text-[10px] px-1.5 py-0.5 rounded-full ${
                                bankFilterType === chip.value ? 'bg-zinc-800 text-zinc-200' : 'bg-zinc-200 text-zinc-600'
                              }`}>{chip.count}</span>
                            </button>
                          ))}
                        </div>
                      );
                    })()}
                  </div>

                  {/* Master checklist warnings if not uploaded */}
                  {!tallyContext && (
                    <div className="p-3 bg-amber-50 border border-amber-200 rounded-xl flex items-center gap-2 text-amber-800 text-xs">
                      <AlertCircle className="w-4 h-4 shrink-0 text-amber-600" />
                      <span><strong>Warning:</strong> Tally Masters context is not loaded. Missing masters cannot be verified before XML generation. Vouchers will map using standard typing.</span>
                    </div>
                  )}

                  {/* Notifications */}
                  {bankSuccessMessage && (
                    <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-xl flex items-center gap-2 text-emerald-800 text-xs font-semibold animate-pulse">
                      <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
                      <span>{bankSuccessMessage}</span>
                    </div>
                  )}

                  {/* Full Ledger Grid Table */}
                  <div className="border border-zinc-200 rounded-xl overflow-hidden bg-white shadow-sm max-h-[600px] overflow-y-auto">
                    {(() => {
                      // Compute matching rows
                      let items = [...bankStatementRows];
                      
                      if (bankSearchTerm.trim()) {
                        const term = bankSearchTerm.toLowerCase();
                        items = items.filter(row => {
                          const dateMatch = (row.date || '').toLowerCase().includes(term);
                          const descMatch = (row.description || '').toLowerCase().includes(term);
                          const refMatch = (row.reference || '').toLowerCase().includes(term);
                          const suggMatch = (row.suggestedLedger || '').toLowerCase().includes(term);
                          const finalMatch = (row.userLedger || '').toLowerCase().includes(term);
                          const typeMatch = (row.detectedVoucherType || '').toLowerCase().includes(term);
                          const amountMatch = (row.amount || 0).toString().includes(term);
                          return dateMatch || descMatch || refMatch || suggMatch || finalMatch || typeMatch || amountMatch;
                        });
                      }

                      if (bankFilterType !== 'All') {
                        items = items.filter(row => {
                          const isBlank = !row.userLedger || !row.userLedger.trim();
                          const isSuspense = row.userLedger && /suspense/i.test(row.userLedger);
                          const isMissingFromMaster = tallyContext && row.userLedger && !tallyContext.ledgers.includes(row.userLedger);
                          const isLowConfidence = row.confidence < 90;
                          
                          switch (bankFilterType) {
                            case 'Valid':
                              return row.status === 'valid' && !row.excluded;
                            case 'Invalid':
                              return row.status === 'invalid';
                            case 'Warnings':
                              return row.status === 'warning';
                            case 'Blank Ledger':
                              return isBlank;
                            case 'Suspense':
                              return isSuspense;
                            case 'Missing Master':
                              return isMissingFromMaster;
                            case 'Low Confidence':
                              return isLowConfidence;
                            case 'Payment':
                              return row.detectedVoucherType === 'Payment';
                            case 'Receipt':
                              return row.detectedVoucherType === 'Receipt';
                            default:
                              return true;
                          }
                        });
                      }

                      if (bankShowAttentionRequiredFirst) {
                        items.sort((a, b) => {
                          const getPriority = (row: any) => {
                            const isBlank = !row.userLedger || !row.userLedger.trim();
                            const isSuspense = row.userLedger && /suspense/i.test(row.userLedger);
                            const isMissingFromMaster = tallyContext && row.userLedger && !tallyContext.ledgers.includes(row.userLedger);
                            const isLowConf = row.confidence < 90;
                            const isIncluded = !row.excluded;

                            if (row.status === 'invalid') return 1;
                            if (isIncluded && isBlank) return 2;
                            if (isIncluded && isSuspense) return 3;
                            if (isIncluded && isMissingFromMaster) return 4;
                            if (isIncluded && isLowConf) return 5;
                            if (row.status === 'warning') return 6;
                            return 7;
                          };

                          return getPriority(a) - getPriority(b);
                        });
                      }

                      if (items.length === 0) {
                        return (
                          <div className="p-12 text-center text-zinc-400">
                            No transactions match the selected search or filter options.
                          </div>
                        );
                      }

                      return (
                        <table className="w-full text-[11px] text-left border-collapse table-fixed min-w-[1400px]">
                          <thead className="bg-zinc-50 text-zinc-600 uppercase font-bold border-b border-zinc-200 sticky top-0 z-10">
                            <tr>
                              <th className="p-3 w-16 text-center">Include</th>
                              <th className="p-3 w-14 text-center">Row</th>
                              <th className="p-3 w-24">Date</th>
                              <th className="p-3 w-40 text-center">Voucher Type</th>
                              <th className="p-3 w-80">Narration / Description</th>
                              <th className="p-3 w-28 text-right">Debit</th>
                              <th className="p-3 w-28 text-right">Credit</th>
                              <th className="p-3 w-28 text-right">Amount</th>
                              <th className="p-3 w-28">Reference</th>
                              <th className="p-3 w-40">Suggested Ledger</th>
                              <th className="p-3 w-20 text-center">Conf %</th>
                              <th className="p-3 w-56">Final Ledger Account</th>
                              <th className="p-3 w-32">Master Status</th>
                              <th className="p-3 w-48">Warning / Error</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-zinc-100">
                            {items.map((row) => {
                              const isInvalid = row.status === 'invalid';
                              const isWarning = row.status === 'warning';
                              const isExcluded = row.excluded;

                              const isBlank = !row.userLedger || !row.userLedger.trim();
                              const isSuspense = row.userLedger && /suspense/i.test(row.userLedger);
                              const isMissingFromMaster = tallyContext && row.userLedger && !tallyContext.ledgers.includes(row.userLedger);
                              const isManuallyChanged = row.userLedger !== row.suggestedLedger;
                              const isValidMaster = tallyContext && row.userLedger && tallyContext.ledgers.includes(row.userLedger) && !isSuspense;

                              const paddingClass = bankCompactRowHeight ? 'p-1.5' : 'p-3';

                              return (
                                <tr 
                                  key={row.rowNo} 
                                  className={`transition-colors ${
                                    isExcluded 
                                      ? 'bg-zinc-50/60 text-zinc-400' 
                                      : isInvalid 
                                        ? 'bg-rose-50/30' 
                                        : isWarning || isBlank || isSuspense
                                          ? 'bg-amber-50/20' 
                                          : 'hover:bg-zinc-50/50'
                                  }`}
                                >
                                  {/* Checkbox */}
                                  <td className={`${paddingClass} text-center`}>
                                    <input
                                      type="checkbox"
                                      checked={!row.excluded}
                                      onChange={(e) => {
                                        const updated = [...bankStatementRows];
                                        const origIdx = bankStatementRows.findIndex(r => r.rowNo === row.rowNo);
                                        if (origIdx !== -1) {
                                          updated[origIdx].excluded = !e.target.checked;
                                          setBankStatementRows(updated);
                                        }
                                      }}
                                      className="rounded border-zinc-300 text-zinc-900 focus:ring-zinc-950 w-4 h-4 cursor-pointer"
                                    />
                                  </td>

                                  {/* Row No */}
                                  <td className={`${paddingClass} text-center font-mono font-medium`}>{row.rowNo}</td>

                                  {/* Date */}
                                  <td className={`${paddingClass} font-semibold whitespace-nowrap`}>{row.date}</td>

                                  {/* Voucher Type Dropdown Toggle */}
                                  <td className={`${paddingClass} text-center`}>
                                    <select
                                      value={row.detectedVoucherType}
                                      onChange={(e) => {
                                        const val = e.target.value as 'Payment' | 'Receipt' | 'Unknown';
                                        const updated = [...bankStatementRows];
                                        const origIdx = bankStatementRows.findIndex(r => r.rowNo === row.rowNo);
                                        if (origIdx !== -1) {
                                          updated[origIdx].detectedVoucherType = val;
                                          setBankStatementRows(updated);
                                        }
                                      }}
                                      disabled={row.excluded}
                                      className="p-1.5 bg-white border border-zinc-200 rounded-lg text-xs font-bold outline-none focus:ring-1 focus:ring-zinc-950 disabled:opacity-50 text-center w-full shadow-xs cursor-pointer"
                                    >
                                      <option value="Payment">Payment</option>
                                      <option value="Receipt">Receipt</option>
                                      <option value="Unknown">Unknown</option>
                                    </select>
                                  </td>

                                  {/* Narration */}
                                  <td className={`${paddingClass} truncate font-medium`} title={row.description}>
                                    {row.description}
                                  </td>

                                  {/* Debit */}
                                  <td className={`${paddingClass} text-right font-mono font-medium text-rose-600`}>
                                    {row.debit !== null && row.debit !== undefined ? row.debit.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'}
                                  </td>

                                  {/* Credit */}
                                  <td className={`${paddingClass} text-right font-mono font-medium text-emerald-600`}>
                                    {row.credit !== null && row.credit !== undefined ? row.credit.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'}
                                  </td>

                                  {/* Amount */}
                                  <td className={`${paddingClass} text-right font-mono font-bold text-zinc-900`}>
                                    {row.amount ? row.amount.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '0.00'}
                                  </td>

                                  {/* Reference */}
                                  <td className={`${paddingClass}`}>
                                    <input
                                      type="text"
                                      value={row.reference || ''}
                                      onChange={(e) => {
                                        const updated = [...bankStatementRows];
                                        const origIdx = bankStatementRows.findIndex(r => r.rowNo === row.rowNo);
                                        if (origIdx !== -1) {
                                          updated[origIdx].reference = e.target.value;
                                          setBankStatementRows(updated);
                                        }
                                      }}
                                      disabled={isExcluded}
                                      placeholder="No Ref"
                                      className="w-full bg-zinc-50 border border-zinc-200 rounded px-2 py-1 text-xs text-zinc-800 focus:border-zinc-400 focus:outline-none font-mono"
                                    />
                                  </td>

                                  {/* Suggested Ledger */}
                                  <td className={`${paddingClass} font-medium text-zinc-700 truncate`} title={row.reasoning || row.suggestedLedger}>
                                    {row.suggestedLedger || <span className="text-zinc-400 italic">No suggestion</span>}
                                  </td>

                                  {/* Confidence */}
                                  <td className={`${paddingClass} text-center`}>
                                    <span className={`inline-block px-1.5 py-0.5 rounded text-[10px] font-bold ${
                                      row.confidence >= 90
                                        ? 'bg-emerald-50 text-emerald-700 border border-emerald-200'
                                        : row.confidence >= 50
                                          ? 'bg-amber-50 text-amber-700 border border-amber-200'
                                          : 'bg-zinc-100 text-zinc-600'
                                    }`}>
                                      {row.confidence ? `${row.confidence}%` : '0%'}
                                    </span>
                                  </td>

                                  {/* Final Ledger Account Search/Editable Input */}
                                  <td className={`${paddingClass}`}>
                                    <input
                                      type="text"
                                      list="tally-ledgers-list"
                                      value={row.userLedger || ''}
                                      onChange={(e) => {
                                        const val = e.target.value;
                                        const updated = [...bankStatementRows];
                                        const origIdx = bankStatementRows.findIndex(r => r.rowNo === row.rowNo);
                                        if (origIdx !== -1) {
                                          updated[origIdx].userLedger = val;
                                          setBankStatementRows(updated);
                                        }
                                      }}
                                      disabled={isExcluded}
                                      placeholder="Type or select ledger"
                                      className="w-full p-2 bg-white border border-zinc-200 rounded-lg text-xs outline-none focus:ring-1 focus:ring-zinc-950 disabled:opacity-50 font-semibold text-zinc-800"
                                    />
                                  </td>

                                  {/* Master Status Badge */}
                                  <td className={`${paddingClass}`}>
                                    {isBlank && (
                                      <span className="px-2 py-0.5 rounded bg-zinc-100 text-zinc-600 text-[9px] font-bold">Blank</span>
                                    )}
                                    {isSuspense && (
                                      <span className="px-2 py-0.5 rounded bg-amber-100 text-amber-800 text-[9px] font-bold">Suspense</span>
                                    )}
                                    {isMissingFromMaster && (
                                      <span className="px-2 py-0.5 rounded bg-rose-100 text-rose-800 text-[9px] font-bold">Missing Master</span>
                                    )}
                                    {isManuallyChanged && !isMissingFromMaster && !isSuspense && !isBlank && (
                                      <span className="px-2 py-0.5 rounded bg-blue-100 text-blue-800 text-[9px] font-bold" title={`Original suggested: ${row.suggestedLedger || 'none'}`}>Manual edit</span>
                                    )}
                                    {isValidMaster && !isManuallyChanged && (
                                      <span className="px-2 py-0.5 rounded bg-emerald-100 text-emerald-800 text-[9px] font-bold">Valid Master</span>
                                    )}
                                  </td>

                                  {/* Error / Warning info */}
                                  <td className={`${paddingClass} truncate text-[10px]`} title={row.errorMsg || row.reasoning}>
                                    {isInvalid && (
                                      <span className="text-rose-600 font-medium flex items-center gap-1">
                                        <AlertCircle className="w-3 h-3 shrink-0" />
                                        <span>{row.errorMsg || "Invalid State"}</span>
                                      </span>
                                    )}
                                    {isBlank && (
                                      <span className="text-rose-600 font-medium flex items-center gap-1">
                                        <AlertCircle className="w-3 h-3 shrink-0" />
                                        <span>Needs Review: Ledger empty</span>
                                      </span>
                                    )}
                                    {isSuspense && (
                                      <span className="text-amber-600 font-medium flex items-center gap-1">
                                        <AlertCircle className="w-3 h-3 shrink-0" />
                                        <span>Suspense selected</span>
                                      </span>
                                    )}
                                    {isWarning && !isInvalid && !isBlank && !isSuspense && (
                                      <span className="text-amber-600 font-medium flex items-center gap-1">
                                        <AlertCircle className="w-3 h-3 shrink-0" />
                                        <span>{row.errorMsg || "Review Warning"}</span>
                                      </span>
                                    )}
                                    {!isInvalid && !isBlank && !isSuspense && !isWarning && (
                                      <span className="text-emerald-600 font-semibold flex items-center gap-1">
                                        <CheckCircle2 className="w-3 h-3 shrink-0" />
                                        <span>Ready</span>
                                      </span>
                                    )}
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      );
                    })()}
                  </div>

                  {/* Shared datalist for Tally Master ledgers list (loaded once for performance) */}
                  <datalist id="tally-ledgers-list">
                    {tallyContext?.ledgers.map(l => (
                      <option key={l} value={l} />
                    ))}
                  </datalist>

                  {/* Row-wise Error and Validation summary panel */}
                  {bankStatementValidationErrors.length > 0 && (
                    <div className="p-4 bg-rose-50 border border-rose-200 rounded-xl space-y-3">
                      <div className="flex items-center gap-2 text-rose-800 font-bold text-sm">
                        <AlertCircle className="w-4 h-4 shrink-0 text-rose-600" />
                        <span>XML Export Blocked: {bankStatementValidationErrors.length} Issue(s) Need Attention</span>
                      </div>
                      <div className="overflow-x-auto">
                        <table className="w-full text-left border-collapse text-xs">
                          <thead>
                            <tr className="border-b border-rose-200 text-rose-800 font-bold">
                              <th className="pb-2 w-20 text-center">Row No</th>
                              <th className="pb-2 w-1/3">Issue / Validation Failure</th>
                              <th className="pb-2">Suggested Resolution / Fix</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-rose-100 text-rose-700">
                            {bankStatementValidationErrors.map((err, i) => (
                              <tr key={i}>
                                <td className="py-2 text-center font-mono font-bold">{err.rowNo === 0 ? "General" : `#${err.rowNo}`}</td>
                                <td className="py-2 font-medium">{err.issue}</td>
                                <td className="py-2 text-zinc-600 font-medium">{err.suggestedFix}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}

                  {/* Bottom Panel Controls (Suspense conformation & actions) */}
                  <div className="pt-4 border-t border-zinc-100 flex flex-col gap-4">
                    {/* Suspense Confirmation Checkbox */}
                    <div className="bg-zinc-50 border border-zinc-200/60 p-3 rounded-xl flex items-center justify-between gap-4">
                      <div className="text-xs text-zinc-600 font-medium">
                        <span className="font-bold text-zinc-800">Suspense Fallback Confirmation:</span> If you are intentionally parking transactions in Suspense / Suspense Account, check the confirmation to allow exporting those rows.
                      </div>
                      <label className="flex items-center gap-2 cursor-pointer select-none font-bold text-xs text-zinc-900 bg-white border border-zinc-200 rounded-lg px-3 py-1.5 shadow-xs shrink-0 shrink-0">
                        <input
                          type="checkbox"
                          checked={confirmProceedWithSuspense}
                          onChange={(e) => {
                            setConfirmProceedWithSuspense(e.target.checked);
                            // Auto reset errors list if checked
                            if (e.target.checked) {
                              setBankStatementValidationErrors(prev => prev.filter(err => !/suspense/i.test(err.issue)));
                            }
                          }}
                          className="rounded border-zinc-300 text-zinc-900 focus:ring-zinc-950 w-4 h-4 cursor-pointer"
                        />
                        Proceed with Suspense
                      </label>
                    </div>

                    <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                      <div className="text-xs text-zinc-400 font-medium leading-relaxed max-w-xl">
                        Valid, non-excluded rows will be converted. standard Tally XML format will preserve the Excel narration and append reference codes.
                      </div>
                      
                      <div className="flex flex-wrap items-center gap-3 justify-end shrink-0">
                        {/* Cancel / Restart */}
                        <button
                          type="button"
                          onClick={() => {
                            setBankStatementRows([]);
                            setRawGrid([]);
                            setCurrentStep('upload');
                          }}
                          className="px-4 py-2.5 hover:bg-zinc-100 border border-zinc-200 text-zinc-700 rounded-xl text-xs font-semibold transition-all flex items-center gap-1.5"
                        >
                          <RotateCcw className="w-3.5 h-3.5" />
                          Restart over
                        </button>

                        {/* Back to Mapping */}
                        <button
                          type="button"
                          onClick={() => {
                            setCurrentStep('mapping');
                          }}
                          className="px-4 py-2.5 hover:bg-zinc-100 border border-zinc-200 text-zinc-700 rounded-xl text-xs font-semibold transition-all flex items-center gap-1.5"
                        >
                          <ArrowLeft className="w-3.5 h-3.5" />
                          Back to Mapping
                        </button>

                        {/* Download Reviewed Excel */}
                        <button
                          type="button"
                          onClick={downloadReviewedTemplateExcel}
                          className="px-4 py-2.5 bg-zinc-100 hover:bg-zinc-200 border border-zinc-200 text-zinc-800 rounded-xl text-xs font-semibold transition-all flex items-center gap-1.5 shadow-xs"
                        >
                          <Download className="w-3.5 h-3.5" />
                          Download Reviewed Excel
                        </button>

                        {/* Generate XML */}
                        <button
                          type="button"
                          onClick={handleBankStatementGenerateXML}
                          disabled={isProcessing}
                          className="px-5 py-2.5 bg-zinc-950 hover:bg-zinc-850 disabled:bg-zinc-300 text-white rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 disabled:opacity-50 shadow-sm"
                        >
                          {isProcessing ? (
                            <>
                              <Loader2 className="w-3.5 h-3.5 animate-spin" />
                              Generating...
                            </>
                          ) : (
                            <>
                              <FileCode className="w-3.5 h-3.5" />
                              Generate Tally XML
                            </>
                          )}
                        </button>
                      </div>
                    </div>
                  </div>
                </motion.section>
              )}

              {currentStep === 'mapping' && (
                <motion.section 
                  key="mapping"
                  initial={{ opacity: 0, x: -20 }}
                  animate={{ opacity: 1, x: 0 }}
                  exit={{ opacity: 0, x: 20 }}
                  className="bg-white p-6 rounded-2xl border border-zinc-200 shadow-sm"
                >
                  <div className="flex items-center justify-between mb-6">
                    <h2 className="text-xl font-bold flex items-center gap-2">
                      <Sparkles className="w-5 h-5 text-amber-500" />
                      {isGeminiAvailable() ? 'Step 2: AI Column Mapping' : 'Step 2: Local Deterministic Mapping'}
                    </h2>
                    <span className="text-sm text-zinc-500">{pendingFileName}</span>
                  </div>

                  <div className="space-y-3 mb-8">
                    {mappings.map((m, idx) => (
                      <div key={idx} className="flex items-center gap-4 p-4 bg-zinc-50 rounded-xl border border-zinc-100">
                        <div className="flex-1">
                          <p className="text-xs text-zinc-500 uppercase tracking-wider mb-1">Source Column</p>
                          <p className="font-medium">{m.excelColumn}</p>
                        </div>
                        <ArrowRight className="w-4 h-4 text-zinc-300" />
                        <div className="flex-1">
                          <p className="text-xs text-zinc-500 uppercase tracking-wider mb-1">Tally Field</p>
                          <select 
                            value={m.tallyField}
                            onChange={(e) => {
                              const newMappings = [...mappings];
                              newMappings[idx].tallyField = e.target.value;
                              setMappings(newMappings);
                            }}
                            className="w-full bg-transparent font-bold text-zinc-900 outline-none cursor-pointer hover:text-zinc-600 transition-colors"
                          >
                            <option value="DATE">DATE</option>
                            <option value="PARTYNAME">PARTYNAME</option>
                            <option value="VOUCHERNUMBER">VOUCHERNUMBER</option>
                            <option value="NARRATION">NARRATION</option>
                            <option value="AMOUNT">AMOUNT</option>
                            <option value="REFERENCE">REFERENCE</option>
                            <option value="IGNORE">IGNORE</option>
                          </select>
                        </div>
                        <div className="hidden md:block w-24">
                          <p className="text-xs text-zinc-500 uppercase tracking-wider mb-1">Confidence</p>
                          <div className="h-1.5 w-full bg-zinc-200 rounded-full overflow-hidden">
                            <div className="h-full bg-green-500" style={{ width: `${m.confidence * 100}%` }} />
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>

                  {error && (
                    <div className="mb-4 p-4 bg-red-50 border border-red-100 rounded-xl flex items-center gap-3 text-red-700">
                      <AlertCircle className="w-5 h-5 shrink-0" />
                      <p className="text-sm">{error}</p>
                    </div>
                  )}

                  <div className="flex gap-3">
                    <button 
                      onClick={() => setCurrentStep('upload')}
                      className="flex-1 py-3 px-4 rounded-xl font-medium border border-zinc-200 hover:bg-zinc-50 transition-colors"
                    >
                      Back
                    </button>
                    <button 
                      onClick={startConversion}
                      disabled={isProcessing}
                      className="flex-[2] flex items-center justify-center gap-2 bg-zinc-900 text-white py-3 px-4 rounded-xl font-medium hover:bg-zinc-800 transition-colors disabled:opacity-50"
                    >
                      {isProcessing ? <Loader2 className="w-5 h-5 animate-spin" /> : <CheckCircle2 className="w-5 h-5" />}
                      Generate Tally XML
                    </button>
                  </div>

                  {aiMappedTransactions.length > 0 && (
                    <div className="mt-12 space-y-4">
                      <h3 className="text-lg font-bold flex items-center gap-2">
                        <Sparkles className="w-5 h-5 text-amber-500" />
                        {isGeminiAvailable() ? 'Transaction-Level Ledger Suggestions' : 'Local Deterministic Ledger Suggestions'}
                      </h3>
                      <div className="max-h-[400px] overflow-y-auto space-y-2 pr-2">
                        {aiMappedTransactions.map((tx, idx) => (
                          <div key={idx} className="p-4 bg-zinc-50 rounded-xl border border-zinc-100 text-sm">
                            <div className="flex justify-between items-start mb-2">
                              <span className="text-zinc-500 font-mono">{tx.date}</span>
                              <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold uppercase ${
                                tx.confidence > 0.8 ? 'bg-green-100 text-green-700' : 'bg-amber-100 text-amber-700'
                              }`}>
                                {Math.round(tx.confidence * 100)}% Confidence
                              </span>
                            </div>
                            <p className="text-zinc-900 font-medium mb-2">{tx.description}</p>
                            <div className="flex items-center gap-2 p-2 bg-white rounded-lg border border-zinc-200">
                              <ArrowRight className="w-3 h-3 text-zinc-400 shrink-0" />
                              <span className="text-xs text-zinc-500 font-semibold shrink-0">Ledger:</span>
                              {tallyContext ? (
                                <select
                                  value={tx.tallyLedger || ''}
                                  onChange={(e) => {
                                    const updated = [...aiMappedTransactions];
                                    updated[idx].tallyLedger = e.target.value;
                                    setAiMappedTransactions(updated);
                                  }}
                                  className="flex-1 bg-transparent font-bold text-zinc-900 outline-none border border-zinc-200 py-1 text-xs cursor-pointer focus:ring-1 focus:ring-zinc-950 rounded px-1"
                                >
                                  <option value="">-- Select Ledger --</option>
                                  {tallyContext.ledgers.map((l, lIdx) => (
                                    <option key={lIdx} value={l}>{l}</option>
                                  ))}
                                </select>
                              ) : (
                                <input
                                  type="text"
                                  value={tx.tallyLedger || ''}
                                  onChange={(e) => {
                                    const updated = [...aiMappedTransactions];
                                    updated[idx].tallyLedger = e.target.value;
                                    setAiMappedTransactions(updated);
                                  }}
                                  className="flex-1 bg-transparent font-bold text-zinc-900 outline-none border border-zinc-200 py-1 text-xs px-2 focus:ring-1 focus:ring-zinc-950 rounded"
                                />
                              )}
                            </div>
                            <p className="text-[11px] text-zinc-500 mt-2 italic">Reason: {tx.reasoning}</p>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </motion.section>
              )}

              {currentStep === 'master-review' && (
                <motion.section 
                  key="master-review"
                  initial={{ opacity: 0, x: -20 }}
                  animate={{ opacity: 1, x: 0 }}
                  exit={{ opacity: 0, x: 20 }}
                  className="bg-white p-6 rounded-2xl border border-zinc-200 shadow-sm max-w-full overflow-hidden"
                >
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-6">
                    <div>
                      <h2 className="text-xl font-bold flex items-center gap-2">
                        <Sparkles className="w-5 h-5 text-amber-500" />
                        Review Tally {importType} Masters
                      </h2>
                      <p className="text-sm text-zinc-500">
                        Edit records inline, resolve validation issues, or exclude rows before exporting to Tally XML.
                      </p>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      <button
                        type="button"
                        onClick={downloadValidationReport}
                        className="py-2 px-3 rounded-lg border border-zinc-200 hover:bg-zinc-50 text-xs font-medium flex items-center gap-1.5 transition-colors"
                      >
                        <Download className="w-4 h-4" />
                        Download Report
                      </button>
                      <button
                        type="button"
                        onClick={() => setCurrentStep('upload')}
                        className="py-2 px-3 rounded-lg border border-zinc-200 hover:bg-zinc-50 text-xs font-medium transition-colors"
                      >
                        Cancel
                      </button>
                      <button
                        type="button"
                        onClick={generateMasterXML}
                        disabled={isProcessing}
                        className="py-2 px-4 rounded-lg bg-zinc-900 text-white hover:bg-zinc-800 text-xs font-semibold flex items-center gap-1.5 transition-colors shadow-sm disabled:opacity-50"
                      >
                        {isProcessing ? <Loader2 className="w-4 h-4 animate-spin" /> : <CheckCircle2 className="w-4 h-4" />}
                        Generate XML ({
                          importType === 'Ledger' ? parsedLedgers.filter(l => !l.excluded && l.isValid).length :
                          importType === 'StockItem' ? parsedStockItems.filter(i => !i.excluded && i.isValid).length :
                          importType === 'StockGroup' ? parsedStockGroups.filter(g => !g.excluded && g.isValid).length :
                          parsedUnits.filter(u => !u.excluded && u.isValid).length
                        } Active)
                      </button>
                    </div>
                  </div>

                  {/* Filter Toolbar */}
                  <div className="flex flex-col sm:flex-row items-center gap-3 mb-6 bg-zinc-50 p-4 rounded-xl border border-zinc-100">
                    <div className="w-full sm:w-72 relative">
                      <Search className="w-4 h-4 text-zinc-400 absolute left-3 top-1/2 -translate-y-1/2" />
                      <input
                        type="text"
                        placeholder="Search rows..."
                        value={masterReviewSearch}
                        onChange={(e) => setMasterReviewSearch(e.target.value)}
                        className="w-full pl-9 pr-4 py-2 bg-white border border-zinc-200 rounded-lg text-sm outline-none focus:ring-2 focus:ring-zinc-900 transition-all"
                      />
                    </div>

                    <div className="flex items-center gap-2 w-full sm:w-auto">
                      <span className="text-xs font-medium text-zinc-500 whitespace-nowrap">Filter:</span>
                      <select
                        value={masterReviewFilter}
                        onChange={(e) => setMasterReviewFilter(e.target.value as any)}
                        className="bg-white border border-zinc-200 rounded-lg text-xs py-2 px-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all cursor-pointer"
                      >
                        <option value="all">All Rows</option>
                        <option value="valid">Valid Rows</option>
                        <option value="invalid">Invalid/Errors</option>
                        {importType === 'Ledger' || importType === 'StockItem' ? (
                          <>
                            <option value="warning">With Warnings</option>
                            <option value="duplicate">Duplicates/Possible Duplicates</option>
                          </>
                        ) : null}
                        <option value="excluded">Excluded Rows</option>
                      </select>
                    </div>

                    {/* Stats */}
                    <div className="flex flex-wrap gap-2 sm:ml-auto">
                      <span className="text-xs bg-zinc-100 text-zinc-700 font-medium px-2.5 py-1 rounded-full">
                        Total: {
                          importType === 'Ledger' ? parsedLedgers.length :
                          importType === 'StockItem' ? parsedStockItems.length :
                          importType === 'StockGroup' ? parsedStockGroups.length :
                          parsedUnits.length
                        }
                      </span>
                      <span className="text-xs bg-green-50 text-green-700 font-medium px-2.5 py-1 rounded-full border border-green-100">
                        Valid: {
                          importType === 'Ledger' ? parsedLedgers.filter(l => l.isValid && !l.excluded).length :
                          importType === 'StockItem' ? parsedStockItems.filter(i => i.isValid && !i.excluded).length :
                          importType === 'StockGroup' ? parsedStockGroups.filter(g => g.isValid && !g.excluded).length :
                          parsedUnits.filter(u => u.isValid && !u.excluded).length
                        }
                      </span>
                      {(importType === 'Ledger' || importType === 'StockItem') && (
                        <>
                          <span className="text-xs bg-red-50 text-red-700 font-medium px-2.5 py-1 rounded-full border border-red-100">
                            Errors: {
                              importType === 'Ledger' ? parsedLedgers.filter(l => !l.isValid).length :
                              parsedStockItems.filter(i => !i.isValid).length
                            }
                          </span>
                          <span className="text-xs bg-amber-50 text-amber-700 font-medium px-2.5 py-1 rounded-full border border-amber-100">
                            Warnings: {
                              importType === 'Ledger' ? parsedLedgers.filter(l => l.warnings.length > 0).length :
                              parsedStockItems.filter(i => i.warnings.length > 0).length
                            }
                          </span>
                        </>
                      )}
                    </div>
                  </div>

                  {/* Spreadsheet Grid Container */}
                  <div className="overflow-x-auto border border-zinc-200 rounded-xl max-h-[500px]">
                    {importType === 'Ledger' && (
                      <table className="min-w-full divide-y divide-zinc-200 text-sm">
                        <thead className="bg-zinc-50 text-xs font-semibold text-zinc-700 sticky top-0 z-10">
                          <tr>
                            <th className="px-3 py-3 text-left w-12 bg-zinc-50">Exclude</th>
                            <th className="px-3 py-3 text-left w-12 bg-zinc-50">Row</th>
                            <th className="px-3 py-3 text-left w-24 bg-zinc-50">Status</th>
                            <th className="px-3 py-3 text-left min-w-[200px] bg-zinc-50">Ledger Name *</th>
                            <th className="px-3 py-3 text-left min-w-[180px] bg-zinc-50">Under Group *</th>
                            <th className="px-3 py-3 text-left min-w-[120px] bg-zinc-50">Opening Bal</th>
                            <th className="px-3 py-3 text-left w-20 bg-zinc-50">Dr/Cr</th>
                            <th className="px-3 py-3 text-left min-w-[180px] bg-zinc-50">Mailing Name</th>
                            <th className="px-3 py-3 text-left min-w-[150px] bg-zinc-50">GSTIN</th>
                            <th className="px-3 py-3 text-left min-w-[150px] bg-zinc-50">State</th>
                            <th className="px-3 py-3 text-left min-w-[130px] bg-zinc-50">Reg. Type</th>
                            <th className="px-3 py-3 text-left min-w-[150px] bg-zinc-50">Email</th>
                          </tr>
                        </thead>
                        <tbody className="bg-white divide-y divide-zinc-200 font-mono text-xs">
                          {getFilteredLedgers().map((l, idx) => {
                            const originalIdx = parsedLedgers.findIndex(item => item.rowNum === l.rowNum);
                            return (
                              <tr key={l.rowNum} className={`hover:bg-zinc-50 transition-colors ${l.excluded ? 'bg-zinc-100/50 opacity-60' : ''}`}>
                                <td className="px-3 py-2 text-center">
                                  <input
                                    type="checkbox"
                                    checked={l.excluded}
                                    onChange={(e) => {
                                      const updated = [...parsedLedgers];
                                      updated[originalIdx].excluded = e.target.checked;
                                      setParsedLedgers(updated);
                                    }}
                                    className="rounded border-zinc-300 text-zinc-900 focus:ring-zinc-900 w-4 h-4 cursor-pointer"
                                  />
                                </td>
                                <td className="px-3 py-2 text-zinc-500 font-medium">{l.rowNum}</td>
                                <td className="px-3 py-2">
                                  <div className="flex flex-col gap-1">
                                    {l.excluded ? (
                                      <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-zinc-100 text-zinc-600">
                                        Excluded
                                      </span>
                                    ) : (
                                      <>
                                        {l.isValid ? (
                                          <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-green-50 text-green-700 border border-green-100">
                                            Valid
                                          </span>
                                        ) : (
                                          <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-red-50 text-red-700 border border-red-100" title={l.errors.join(', ')}>
                                            Error ({l.errors.length})
                                          </span>
                                        )}
                                        {l.warnings.length > 0 && (
                                          <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-amber-50 text-amber-700 border border-amber-100" title={l.warnings.join(', ')}>
                                            Warning
                                          </span>
                                        )}
                                        {(l.isDuplicate || l.isPossibleDuplicate) && (
                                          <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-indigo-50 text-indigo-700 border border-indigo-100" title={l.duplicateMessage}>
                                            Duplicate
                                          </span>
                                        )}
                                      </>
                                    )}
                                  </div>
                                </td>
                                <td className="px-2 py-1.5">
                                  <input
                                    type="text"
                                    value={l.ledgerName}
                                    onChange={(e) => handleLedgerCellChange(originalIdx, 'ledgerName', e.target.value)}
                                    className={`w-full px-2 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none ${!l.ledgerName.trim() ? 'border-red-300 bg-red-50/20' : ''}`}
                                  />
                                </td>
                                <td className="px-2 py-1.5">
                                  <select
                                    value={l.underGroup}
                                    onChange={(e) => handleLedgerCellChange(originalIdx, 'underGroup', e.target.value)}
                                    className={`w-full px-1.5 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none ${!l.underGroup.trim() ? 'border-red-300 bg-red-50/20' : ''}`}
                                  >
                                    <option value="">Select Group...</option>
                                    {getGroupOptions().map(grp => (
                                      <option key={grp} value={grp}>{grp}</option>
                                    ))}
                                  </select>
                                </td>
                                <td className="px-2 py-1.5">
                                  <input
                                    type="text"
                                    value={l.openingBalance}
                                    onChange={(e) => handleLedgerCellChange(originalIdx, 'openingBalance', e.target.value)}
                                    className="w-full px-2 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none text-right"
                                  />
                                </td>
                                <td className="px-2 py-1.5">
                                  <select
                                    value={l.drCr}
                                    onChange={(e) => handleLedgerCellChange(originalIdx, 'drCr', e.target.value)}
                                    className="w-full px-1.5 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none"
                                  >
                                    <option value="Dr">Dr</option>
                                    <option value="Cr">Cr</option>
                                  </select>
                                </td>
                                <td className="px-2 py-1.5">
                                  <input
                                    type="text"
                                    value={l.mailingName || ''}
                                    onChange={(e) => handleLedgerCellChange(originalIdx, 'mailingName', e.target.value)}
                                    className="w-full px-2 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none"
                                  />
                                </td>
                                <td className="px-2 py-1.5">
                                  <input
                                    type="text"
                                    value={l.gstin || ''}
                                    placeholder="27AAAAA0000A1Z0"
                                    onChange={(e) => handleLedgerCellChange(originalIdx, 'gstin', e.target.value)}
                                    className={`w-full px-2 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none uppercase ${l.gstin && !validateGSTIN(l.gstin).isValid ? 'border-red-300 bg-red-50/20' : ''}`}
                                  />
                                </td>
                                <td className="px-2 py-1.5">
                                  <input
                                    type="text"
                                    value={l.state || ''}
                                    onChange={(e) => handleLedgerCellChange(originalIdx, 'state', e.target.value)}
                                    className="w-full px-2 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none"
                                  />
                                </td>
                                <td className="px-2 py-1.5">
                                  <select
                                    value={l.registrationType}
                                    onChange={(e) => handleLedgerCellChange(originalIdx, 'registrationType', e.target.value)}
                                    className="w-full px-1.5 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none"
                                  >
                                    <option value="Regular">Regular</option>
                                    <option value="Composition">Composition</option>
                                    <option value="Consumer">Consumer</option>
                                    <option value="Unregistered">Unregistered</option>
                                  </select>
                                </td>
                                <td className="px-2 py-1.5">
                                  <input
                                    type="text"
                                    value={l.email || ''}
                                    onChange={(e) => handleLedgerCellChange(originalIdx, 'email', e.target.value)}
                                    className="w-full px-2 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none text-zinc-600"
                                  />
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    )}

                    {importType === 'StockItem' && (
                      <table className="min-w-full divide-y divide-zinc-200 text-sm">
                        <thead className="bg-zinc-50 text-xs font-semibold text-zinc-700 sticky top-0 z-10">
                          <tr>
                            <th className="px-3 py-3 text-left w-12 bg-zinc-50">Exclude</th>
                            <th className="px-3 py-3 text-left w-12 bg-zinc-50">Row</th>
                            <th className="px-3 py-3 text-left w-24 bg-zinc-50">Status</th>
                            <th className="px-3 py-3 text-left min-w-[200px] bg-zinc-50">Stock Item Name *</th>
                            <th className="px-3 py-3 text-left min-w-[180px] bg-zinc-50">Stock Group *</th>
                            <th className="px-3 py-3 text-left min-w-[120px] bg-zinc-50">Unit *</th>
                            <th className="px-3 py-3 text-left min-w-[100px] bg-zinc-50">GST Rate (%)</th>
                            <th className="px-3 py-3 text-left min-w-[120px] bg-zinc-50">Opening Qty</th>
                            <th className="px-3 py-3 text-left min-w-[120px] bg-zinc-50">Opening Rate</th>
                            <th className="px-3 py-3 text-left min-w-[120px] bg-zinc-50">Opening Value</th>
                            <th className="px-3 py-3 text-left min-w-[150px] bg-zinc-50">HSN/SAC</th>
                            <th className="px-3 py-3 text-left min-w-[150px] bg-zinc-50">Description</th>
                          </tr>
                        </thead>
                        <tbody className="bg-white divide-y divide-zinc-200 font-mono text-xs">
                          {getFilteredStockItems().map((i, idx) => {
                            const originalIdx = parsedStockItems.findIndex(item => item.rowNum === i.rowNum);
                            return (
                              <tr key={i.rowNum} className={`hover:bg-zinc-50 transition-colors ${i.excluded ? 'bg-zinc-100/50 opacity-60' : ''}`}>
                                <td className="px-3 py-2 text-center">
                                  <input
                                    type="checkbox"
                                    checked={i.excluded}
                                    onChange={(e) => {
                                      const updated = [...parsedStockItems];
                                      updated[originalIdx].excluded = e.target.checked;
                                      setParsedStockItems(updated);
                                    }}
                                    className="rounded border-zinc-300 text-zinc-900 focus:ring-zinc-900 w-4 h-4 cursor-pointer"
                                  />
                                </td>
                                <td className="px-3 py-2 text-zinc-500 font-medium">{i.rowNum}</td>
                                <td className="px-3 py-2">
                                  <div className="flex flex-col gap-1">
                                    {i.excluded ? (
                                      <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-zinc-100 text-zinc-600">
                                        Excluded
                                      </span>
                                    ) : (
                                      <>
                                        {i.isValid ? (
                                          <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-green-50 text-green-700 border border-green-100">
                                            Valid
                                          </span>
                                        ) : (
                                          <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-red-50 text-red-700 border border-red-100" title={i.errors.join(', ')}>
                                            Error ({i.errors.length})
                                          </span>
                                        )}
                                        {i.warnings.length > 0 && (
                                          <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-amber-50 text-amber-700 border border-amber-100" title={i.warnings.join(', ')}>
                                            Warning
                                          </span>
                                        )}
                                        {(i.isDuplicate || i.isPossibleDuplicate) && (
                                          <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-indigo-50 text-indigo-700 border border-indigo-100" title={i.duplicateMessage}>
                                            Duplicate
                                          </span>
                                        )}
                                      </>
                                    )}
                                  </div>
                                </td>
                                <td className="px-2 py-1.5">
                                  <input
                                    type="text"
                                    value={i.itemName}
                                    onChange={(e) => handleStockItemCellChange(originalIdx, 'itemName', e.target.value)}
                                    className={`w-full px-2 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none ${!i.itemName.trim() ? 'border-red-300 bg-red-50/20' : ''}`}
                                  />
                                </td>
                                <td className="px-2 py-1.5">
                                  <select
                                    value={i.underGroup}
                                    onChange={(e) => handleStockItemCellChange(originalIdx, 'underGroup', e.target.value)}
                                    className={`w-full px-1.5 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none ${!i.underGroup.trim() ? 'border-red-300 bg-red-50/20' : ''}`}
                                  >
                                    <option value="">Select Stock Group...</option>
                                    {getStockGroupOptions().map(grp => (
                                      <option key={grp} value={grp}>{grp}</option>
                                    ))}
                                  </select>
                                </td>
                                <td className="px-2 py-1.5">
                                  <select
                                    value={i.unit}
                                    onChange={(e) => handleStockItemCellChange(originalIdx, 'unit', e.target.value)}
                                    className={`w-full px-1.5 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none ${!i.unit.trim() ? 'border-red-300 bg-red-50/20' : ''}`}
                                  >
                                    <option value="">Select Unit...</option>
                                    {getUnitOptions().map(ut => (
                                      <option key={ut} value={ut}>{ut}</option>
                                    ))}
                                  </select>
                                </td>
                                <td className="px-2 py-1.5">
                                  <input
                                    type="text"
                                    value={i.gstRate || ''}
                                    placeholder="18"
                                    onChange={(e) => handleStockItemCellChange(originalIdx, 'gstRate', e.target.value)}
                                    className={`w-full px-2 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none text-right ${i.gstRate && isNaN(Number(i.gstRate)) ? 'border-red-300 bg-red-50/20' : ''}`}
                                  />
                                </td>
                                <td className="px-2 py-1.5">
                                  <input
                                    type="text"
                                    value={i.openingQty || ''}
                                    onChange={(e) => handleStockItemCellChange(originalIdx, 'openingQty', e.target.value)}
                                    className={`w-full px-2 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none text-right ${i.openingQty && isNaN(Number(i.openingQty)) ? 'border-red-300 bg-red-50/20' : ''}`}
                                  />
                                </td>
                                <td className="px-2 py-1.5">
                                  <input
                                    type="text"
                                    value={i.openingRate || ''}
                                    onChange={(e) => handleStockItemCellChange(originalIdx, 'openingRate', e.target.value)}
                                    className={`w-full px-2 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none text-right ${i.openingRate && isNaN(Number(i.openingRate)) ? 'border-red-300 bg-red-50/20' : ''}`}
                                  />
                                </td>
                                <td className="px-2 py-1.5">
                                  <input
                                    type="text"
                                    value={i.openingValue || ''}
                                    onChange={(e) => handleStockItemCellChange(originalIdx, 'openingValue', e.target.value)}
                                    className={`w-full px-2 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none text-right ${i.openingValue && isNaN(Number(i.openingValue)) ? 'border-red-300 bg-red-50/20' : ''}`}
                                  />
                                </td>
                                <td className="px-2 py-1.5">
                                  <input
                                    type="text"
                                    value={i.hsn || ''}
                                    onChange={(e) => handleStockItemCellChange(originalIdx, 'hsn', e.target.value)}
                                    className="w-full px-2 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none"
                                  />
                                </td>
                                <td className="px-2 py-1.5">
                                  <input
                                    type="text"
                                    value={i.description || ''}
                                    onChange={(e) => handleStockItemCellChange(originalIdx, 'description', e.target.value)}
                                    className="w-full px-2 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none text-zinc-600"
                                  />
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    )}

                    {importType === 'StockGroup' && (
                      <table className="min-w-full divide-y divide-zinc-200 text-sm">
                        <thead className="bg-zinc-50 text-xs font-semibold text-zinc-700 sticky top-0 z-10">
                          <tr>
                            <th className="px-3 py-3 text-left w-12 bg-zinc-50">Exclude</th>
                            <th className="px-3 py-3 text-left w-12 bg-zinc-50">Row</th>
                            <th className="px-3 py-3 text-left w-24 bg-zinc-50">Status</th>
                            <th className="px-3 py-3 text-left min-w-[250px] bg-zinc-50">Stock Group Name *</th>
                            <th className="px-3 py-3 text-left min-w-[250px] bg-zinc-50">Under Parent Group</th>
                          </tr>
                        </thead>
                        <tbody className="bg-white divide-y divide-zinc-200 font-mono text-xs">
                          {getFilteredStockGroups().map((g, idx) => {
                            const originalIdx = parsedStockGroups.findIndex(item => item.rowNum === g.rowNum);
                            return (
                              <tr key={g.rowNum} className={`hover:bg-zinc-50 transition-colors ${g.excluded ? 'bg-zinc-100/50 opacity-60' : ''}`}>
                                <td className="px-3 py-2 text-center">
                                  <input
                                    type="checkbox"
                                    checked={g.excluded}
                                    onChange={(e) => {
                                      const updated = [...parsedStockGroups];
                                      updated[originalIdx].excluded = e.target.checked;
                                      setParsedStockGroups(updated);
                                    }}
                                    className="rounded border-zinc-300 text-zinc-900 focus:ring-zinc-900 w-4 h-4 cursor-pointer"
                                  />
                                </td>
                                <td className="px-3 py-2 text-zinc-500 font-medium">{g.rowNum}</td>
                                <td className="px-3 py-2">
                                  {g.excluded ? (
                                    <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-zinc-100 text-zinc-600">
                                      Excluded
                                    </span>
                                  ) : g.isValid ? (
                                    <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-green-50 text-green-700 border border-green-100">
                                      Valid
                                    </span>
                                  ) : (
                                    <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-red-50 text-red-700 border border-red-100" title={g.errors.join(', ')}>
                                      Error
                                    </span>
                                  )}
                                </td>
                                <td className="px-2 py-1.5">
                                  <input
                                    type="text"
                                    value={g.groupName}
                                    onChange={(e) => {
                                      const updated = [...parsedStockGroups];
                                      updated[originalIdx].groupName = e.target.value;
                                      updated[originalIdx].isValid = !!e.target.value.trim();
                                      updated[originalIdx].errors = e.target.value.trim() ? [] : ["Stock Group Name cannot be blank."];
                                      setParsedStockGroups(updated);
                                    }}
                                    className={`w-full px-2 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none ${!g.groupName.trim() ? 'border-red-300 bg-red-50/20' : ''}`}
                                  />
                                </td>
                                <td className="px-2 py-1.5">
                                  <input
                                    type="text"
                                    value={g.underGroup}
                                    onChange={(e) => {
                                      const updated = [...parsedStockGroups];
                                      updated[originalIdx].underGroup = e.target.value;
                                      setParsedStockGroups(updated);
                                    }}
                                    className="w-full px-2 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none"
                                  />
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    )}

                    {importType === 'Unit' && (
                      <table className="min-w-full divide-y divide-zinc-200 text-sm">
                        <thead className="bg-zinc-50 text-xs font-semibold text-zinc-700 sticky top-0 z-10">
                          <tr>
                            <th className="px-3 py-3 text-left w-12 bg-zinc-50">Exclude</th>
                            <th className="px-3 py-3 text-left w-12 bg-zinc-50">Row</th>
                            <th className="px-3 py-3 text-left w-24 bg-zinc-50">Status</th>
                            <th className="px-3 py-3 text-left min-w-[150px] bg-zinc-50">Symbol *</th>
                            <th className="px-3 py-3 text-left min-w-[200px] bg-zinc-50">Formal Name</th>
                            <th className="px-3 py-3 text-left min-w-[180px] bg-zinc-50">UQC</th>
                            <th className="px-3 py-3 text-left min-w-[120px] bg-zinc-50">Decimals</th>
                          </tr>
                        </thead>
                        <tbody className="bg-white divide-y divide-zinc-200 font-mono text-xs">
                          {getFilteredUnits().map((u, idx) => {
                            const originalIdx = parsedUnits.findIndex(item => item.rowNum === u.rowNum);
                            return (
                              <tr key={u.rowNum} className={`hover:bg-zinc-50 transition-colors ${u.excluded ? 'bg-zinc-100/50 opacity-60' : ''}`}>
                                <td className="px-3 py-2 text-center">
                                  <input
                                    type="checkbox"
                                    checked={u.excluded}
                                    onChange={(e) => {
                                      const updated = [...parsedUnits];
                                      updated[originalIdx].excluded = e.target.checked;
                                      setParsedUnits(updated);
                                    }}
                                    className="rounded border-zinc-300 text-zinc-900 focus:ring-zinc-900 w-4 h-4 cursor-pointer"
                                  />
                                </td>
                                <td className="px-3 py-2 text-zinc-500 font-medium">{u.rowNum}</td>
                                <td className="px-3 py-2">
                                  {u.excluded ? (
                                    <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-zinc-100 text-zinc-600">
                                      Excluded
                                    </span>
                                  ) : u.isValid ? (
                                    <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-green-50 text-green-700 border border-green-100">
                                      Valid
                                    </span>
                                  ) : (
                                    <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-medium bg-red-50 text-red-700 border border-red-100" title={u.errors.join(', ')}>
                                      Error
                                    </span>
                                  )}
                                </td>
                                <td className="px-2 py-1.5">
                                  <input
                                    type="text"
                                    value={u.symbol}
                                    onChange={(e) => {
                                      const updated = [...parsedUnits];
                                      updated[originalIdx].symbol = e.target.value;
                                      updated[originalIdx].isValid = !!e.target.value.trim();
                                      updated[originalIdx].errors = e.target.value.trim() ? [] : ["Symbol cannot be blank."];
                                      setParsedUnits(updated);
                                    }}
                                    className={`w-full px-2 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none ${!u.symbol.trim() ? 'border-red-300 bg-red-50/20' : ''}`}
                                  />
                                </td>
                                <td className="px-2 py-1.5">
                                  <input
                                    type="text"
                                    value={u.formalName || ''}
                                    onChange={(e) => {
                                      const updated = [...parsedUnits];
                                      updated[originalIdx].formalName = e.target.value;
                                      setParsedUnits(updated);
                                    }}
                                    className="w-full px-2 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none"
                                  />
                                </td>
                                <td className="px-2 py-1.5">
                                  <input
                                    type="text"
                                    value={u.uqc || ''}
                                    placeholder="NOS-NUMBERS"
                                    onChange={(e) => {
                                      const updated = [...parsedUnits];
                                      updated[originalIdx].uqc = e.target.value;
                                      setParsedUnits(updated);
                                    }}
                                    className="w-full px-2 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none uppercase"
                                  />
                                </td>
                                <td className="px-2 py-1.5">
                                  <input
                                    type="text"
                                    value={u.decimalPlaces}
                                    onChange={(e) => {
                                      const updated = [...parsedUnits];
                                      updated[originalIdx].decimalPlaces = e.target.value;
                                      setParsedUnits(updated);
                                    }}
                                    className="w-full px-2 py-1 bg-transparent border border-transparent rounded hover:border-zinc-200 focus:border-zinc-900 focus:bg-white focus:ring-1 focus:ring-zinc-900 outline-none text-right"
                                  />
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    )}
                  </div>
                </motion.section>
              )}

              {currentStep === 'verification' && (
                <motion.section
                  key="verification"
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -10 }}
                  className="bg-white p-6 rounded-2xl border border-zinc-200 shadow-sm space-y-6"
                >
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-zinc-100 pb-5">
                    <div>
                      <h2 className="text-xl font-bold flex items-center gap-2 text-zinc-900">
                        <CheckCircle2 className="w-5 h-5 text-emerald-600" />
                        Step 3: Verification Screen Before XML
                      </h2>
                      <p className="text-xs text-zinc-500 mt-1 leading-relaxed">
                        Review the precise debit/credit sign logic, amounts, and ledger mappings. You can correct the ledgers here before generating the final XML.
                      </p>
                    </div>
                    <div className="text-right">
                      <span className="text-xs bg-emerald-50 text-emerald-700 border border-emerald-100 px-3 py-1.5 rounded-lg font-bold">
                        {verificationRows.length} Transactions Ready
                      </span>
                    </div>
                  </div>

                  {/* Narration Settings Toggle */}
                  <div className="bg-zinc-50 border border-zinc-200 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                    <div className="space-y-1">
                      <h4 className="text-sm font-bold text-zinc-900">Default Blank Narration Fallback</h4>
                      <p className="text-xs text-zinc-500">
                        When the Excel Narration column is empty, should the system fallback to the selected Ledger Name?
                      </p>
                    </div>
                    <label className="relative inline-flex items-center cursor-pointer select-none">
                      <input 
                        type="checkbox" 
                        checked={useLedgerAsNarration} 
                        onChange={(e) => setUseLedgerAsNarration(e.target.checked)}
                        className="sr-only peer" 
                      />
                      <div className="w-11 h-6 bg-zinc-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-zinc-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-zinc-850"></div>
                      <span className="ml-3 text-xs font-bold text-zinc-700">Use ledger as narration</span>
                    </label>
                  </div>

                  {/* Table Container */}
                  <div className="border border-zinc-200 rounded-xl overflow-hidden bg-white">
                    <div className="max-h-[500px] overflow-y-auto">
                      <table className="w-full border-collapse text-left text-xs">
                        <thead>
                          <tr className="bg-zinc-50 border-b border-zinc-200 font-bold text-zinc-700">
                            <th className="p-3 w-16 text-center">Row</th>
                            <th className="p-3 w-28">Voucher Type</th>
                            <th className="p-3 w-40">Date</th>
                            <th className="p-3">Ledger / Particulars</th>
                            <th className="p-3 w-48">Narration</th>
                            <th className="p-3 w-36">Reference</th>
                            <th className="p-3 w-64">Final XML Narration Preview</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-zinc-100">
                          {verificationRows.map((row, idx) => {
                            const isPayment = row.voucherType === 'Payment';
                            const amtAbs = Math.abs(row.amount);
                            const isSuspense = row.finalLedger && /suspense/i.test(row.finalLedger);

                            return (
                              <tr key={idx} className="hover:bg-zinc-50/50 transition-colors">
                                <td className="p-3 text-center font-mono text-zinc-400">{row.rowNo}</td>
                                <td className="p-3">
                                  <span className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-[10px] font-bold uppercase tracking-wide border ${
                                    isPayment 
                                      ? 'bg-rose-50 text-rose-700 border-rose-100' 
                                      : 'bg-emerald-50 text-emerald-700 border-emerald-100'
                                  }`}>
                                    {row.voucherType}
                                  </span>
                                </td>
                                <td className="p-3">
                                  <div className="font-medium text-zinc-900">{row.originalDate}</div>
                                  <div className="text-[10px] font-mono text-zinc-400 mt-0.5">Tally: {row.normalizedDate}</div>
                                </td>
                                <td className="p-3 space-y-2">
                                  {/* Debit Leg */}
                                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 p-2 bg-zinc-50/50 rounded-lg border border-zinc-100">
                                    <div className="flex items-center gap-2">
                                      <span className="text-[10px] font-bold bg-zinc-200 text-zinc-700 px-1.5 py-0.5 rounded uppercase">Dr</span>
                                      {tallyContext ? (
                                        <select
                                          value={isPayment ? row.finalLedger : row.bankLedger}
                                          onChange={(e) => {
                                            if (isPayment) {
                                              updateVerificationRowLedger(idx, e.target.value);
                                            } else {
                                              const updated = [...verificationRows];
                                              updated[idx].bankLedger = e.target.value;
                                              setVerificationRows(updated);
                                            }
                                          }}
                                          className="bg-transparent font-bold text-zinc-800 outline-none text-xs border border-zinc-200 rounded px-1.5 py-0.5 max-w-[200px]"
                                        >
                                          <option value="">-- Select Ledger --</option>
                                          {tallyContext.ledgers.map(l => (
                                            <option key={l} value={l}>{l}</option>
                                          ))}
                                        </select>
                                      ) : (
                                        <input
                                          type="text"
                                          value={isPayment ? row.finalLedger : row.bankLedger}
                                          onChange={(e) => {
                                            if (isPayment) {
                                              updateVerificationRowLedger(idx, e.target.value);
                                            } else {
                                              const updated = [...verificationRows];
                                              updated[idx].bankLedger = e.target.value;
                                              setVerificationRows(updated);
                                            }
                                          }}
                                          className="bg-transparent font-bold text-zinc-800 outline-none text-xs border border-zinc-200 rounded px-1.5 py-0.5"
                                        />
                                      )}
                                      {isPayment && isSuspense && (
                                        <span className="inline-flex items-center gap-1 bg-amber-50 text-amber-700 border border-amber-200 px-2 py-0.5 rounded text-[10px] font-bold animate-pulse">
                                          ⚠️ Suspense
                                        </span>
                                      )}
                                    </div>
                                    <span className="font-mono text-xs font-semibold text-rose-600">
                                      -{amtAbs.toFixed(2)}
                                    </span>
                                  </div>

                                  {/* Credit Leg */}
                                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 p-2 bg-zinc-50/50 rounded-lg border border-zinc-100">
                                    <div className="flex items-center gap-2">
                                      <span className="text-[10px] font-bold bg-zinc-200 text-zinc-700 px-1.5 py-0.5 rounded uppercase">Cr</span>
                                      {tallyContext ? (
                                        <select
                                          value={isPayment ? row.bankLedger : row.finalLedger}
                                          onChange={(e) => {
                                            if (!isPayment) {
                                              updateVerificationRowLedger(idx, e.target.value);
                                            } else {
                                              const updated = [...verificationRows];
                                              updated[idx].bankLedger = e.target.value;
                                              setVerificationRows(updated);
                                            }
                                          }}
                                          className="bg-transparent font-bold text-zinc-800 outline-none text-xs border border-zinc-200 rounded px-1.5 py-0.5 max-w-[200px]"
                                        >
                                          <option value="">-- Select Ledger --</option>
                                          {tallyContext.ledgers.map(l => (
                                            <option key={l} value={l}>{l}</option>
                                          ))}
                                        </select>
                                      ) : (
                                        <input
                                          type="text"
                                          value={isPayment ? row.bankLedger : row.finalLedger}
                                          onChange={(e) => {
                                            if (!isPayment) {
                                              updateVerificationRowLedger(idx, e.target.value);
                                            } else {
                                              const updated = [...verificationRows];
                                              updated[idx].bankLedger = e.target.value;
                                              setVerificationRows(updated);
                                            }
                                          }}
                                          className="bg-transparent font-bold text-zinc-800 outline-none text-xs border border-zinc-200 rounded px-1.5 py-0.5"
                                        />
                                      )}
                                      {!isPayment && isSuspense && (
                                        <span className="inline-flex items-center gap-1 bg-amber-50 text-amber-700 border border-amber-200 px-2 py-0.5 rounded text-[10px] font-bold animate-pulse">
                                          ⚠️ Suspense
                                        </span>
                                      )}
                                    </div>
                                    <span className="font-mono text-xs font-semibold text-emerald-600">
                                      +{amtAbs.toFixed(2)}
                                    </span>
                                  </div>
                                </td>
                                <td className="p-3">
                                  <input
                                    type="text"
                                    value={row.description && row.description !== 'No Narration Found' ? row.description : ''}
                                    onChange={(e) => {
                                      const updated = [...verificationRows];
                                      updated[idx].description = e.target.value;
                                      setVerificationRows(updated);
                                    }}
                                    placeholder="No Narration"
                                    className="w-full bg-transparent border border-zinc-200 rounded px-2 py-1 text-xs text-zinc-800 focus:border-zinc-400 focus:outline-none"
                                  />
                                </td>
                                <td className="p-3">
                                  <input
                                    type="text"
                                    value={row.reference || ''}
                                    onChange={(e) => {
                                      const updated = [...verificationRows];
                                      updated[idx].reference = e.target.value;
                                      setVerificationRows(updated);
                                    }}
                                    placeholder="No Ref"
                                    className="w-full bg-transparent border border-zinc-200 rounded px-2 py-1 text-xs text-zinc-800 focus:border-zinc-400 focus:outline-none font-mono"
                                  />
                                </td>
                                <td className="p-3 font-mono text-[11px] text-zinc-600 bg-zinc-50/30 truncate max-w-[200px]" title={getFinalXMLNarration(row, useLedgerAsNarration)}>
                                  &lt;NARRATION&gt;{getFinalXMLNarration(row, useLedgerAsNarration)}&lt;/NARRATION&gt;
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {/* Footer actions */}
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 pt-4 border-t border-zinc-100">
                    <div className="text-xs text-zinc-500">
                      Review the voucher structure. Both credit and debit legs match perfectly in value. 
                    </div>
                    <div className="flex items-center gap-3 justify-end shrink-0">
                      <button
                        type="button"
                        onClick={() => setCurrentStep(verificationSourceStep)}
                        className="px-5 py-3 hover:bg-zinc-100 border border-zinc-200 text-zinc-700 rounded-xl text-xs font-bold transition-all flex items-center gap-2"
                      >
                        <RotateCcw className="w-4 h-4" />
                        Back to Edit
                      </button>
                      <button
                        type="button"
                        onClick={generateFinalXMLFromVerification}
                        disabled={isProcessing}
                        className="px-6 py-3 bg-zinc-900 hover:bg-zinc-850 text-white rounded-xl text-xs font-bold transition-all flex items-center gap-2 disabled:opacity-50 shadow-sm"
                      >
                        {isProcessing ? (
                          <>
                            <Loader2 className="w-4 h-4 animate-spin" />
                            Generating XML...
                          </>
                        ) : (
                          <>
                            <CheckCircle2 className="w-4 h-4" />
                            Confirm & Generate Tally XML
                          </>
                        )}
                      </button>
                    </div>
                  </div>
                </motion.section>
              )}

              {currentStep === 'sales-purchase-verification' && (() => {
                const activeInv = salesPurchaseInvoices[selectedInvoiceIdx];
                const ledgersList = tallyContext?.ledgers || DEFAULT_LEDGERS;
                const stockItemsList = tallyContext?.stockItems || [];
                const unitsList = tallyContext?.units || DEFAULT_UNITS;

                return (
                  <motion.section
                    key="sales-purchase-verification"
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, y: -10 }}
                    className="bg-white p-6 rounded-2xl border border-zinc-200 shadow-sm space-y-6"
                  >
                    <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-zinc-100 pb-5">
                      <div>
                        <h2 className="text-xl font-bold flex items-center gap-2 text-zinc-900">
                          <CheckCircle2 className="w-5 h-5 text-emerald-600" />
                          Step 3: Verification of Sales / Purchase Invoices
                        </h2>
                        <p className="text-xs text-zinc-500 mt-1 leading-relaxed">
                          Review item lines, verify auto GST calculations, modify ledgers/amounts, and fill transport details before exporting to Tally.
                        </p>
                      </div>
                      <div className="flex items-center gap-3">
                        <span className="text-xs bg-emerald-50 text-emerald-700 border border-emerald-100 px-3 py-1.5 rounded-lg font-bold">
                          {salesPurchaseInvoices.length} Invoices Found
                        </span>
                        <span className="text-xs bg-amber-50 text-amber-700 border border-amber-100 px-3 py-1.5 rounded-lg font-bold">
                          {salesPurchaseInvoices.filter(i => !i.isValid).length} Require Review
                        </span>
                      </div>
                    </div>

                    {salesPurchaseBalancingErrors.length > 0 && (
                      <div className="bg-red-50 border border-red-200 rounded-2xl p-5 space-y-3 shadow-sm">
                        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                          <div className="flex items-center gap-2">
                            <AlertCircle className="w-5 h-5 text-red-600 shrink-0" />
                            <h3 className="font-bold text-red-800 text-sm">
                              Sales/Purchase Voucher-Level Balancing Failed
                            </h3>
                          </div>
                          <button
                            type="button"
                            onClick={() => generateSalesPurchaseBalancingErrorExcel(salesPurchaseBalancingErrors)}
                            className="px-3 py-1.5 bg-red-100 hover:bg-red-200 text-red-700 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 self-start sm:self-center"
                          >
                            <Download className="w-3.5 h-3.5" />
                            Download Balancing Error Log
                          </button>
                        </div>
                        <p className="text-xs text-red-700 leading-relaxed">
                          The following {salesPurchaseBalancingErrors.length} invoice(s) are unbalanced. For each Sales/Purchase voucher, the signed XML ledger amounts must sum to zero before XML generation is allowed.
                        </p>
                        <div className="overflow-x-auto border border-red-200 rounded-xl bg-white">
                          <table className="w-full text-left border-collapse text-xs">
                            <thead>
                              <tr className="bg-red-50/50 border-b border-red-200 font-bold text-red-800">
                                <th className="p-2.5">Invoice No</th>
                                <th className="p-2.5 text-right">Debit Total</th>
                                <th className="p-2.5 text-right">Credit Total</th>
                                <th className="p-2.5 text-right">Expected Party Amount</th>
                                <th className="p-2.5 text-right">Actual Party Amount</th>
                                <th className="p-2.5 text-right">Difference</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-red-100 font-mono">
                              {salesPurchaseBalancingErrors.map((err, errIdx) => (
                                <tr key={errIdx} className="text-red-700 hover:bg-red-50/30">
                                  <td className="p-2.5 font-bold font-sans">{err.invoiceNo || 'Blank'}</td>
                                  <td className="p-2.5 text-right">₹{err.debitTotal.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                                  <td className="p-2.5 text-right">₹{err.creditTotal.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                                  <td className="p-2.5 text-right">₹{err.expectedPartyAmount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                                  <td className="p-2.5 text-right">₹{err.actualPartyAmount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                                  <td className="p-2.5 text-right font-bold text-red-800">₹{err.difference.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}

                    <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">
                      {/* Left Sidebar: Invoice list */}
                      <div className="lg:col-span-1 space-y-3">
                        <div className="p-3 bg-zinc-50 border border-zinc-200/60 rounded-xl mb-4 space-y-1.5">
                          <label className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider block">Company State (GST Base)</label>
                          <select
                            value={companyState}
                            onChange={(e) => {
                              const newState = e.target.value;
                              setCompanyState(newState);
                              const updated = salesPurchaseInvoices.map(inv => {
                                const copy = JSON.parse(JSON.stringify(inv));
                                recalculateInvoice(copy, newState);
                                return copy;
                              });
                              setSalesPurchaseInvoices(updated);
                            }}
                            className="w-full text-xs p-2 bg-white border border-zinc-200 rounded-lg focus:outline-none focus:ring-1 focus:ring-zinc-900 font-medium"
                          >
                            {['Andaman and Nicobar Islands', 'Andhra Pradesh', 'Arunachal Pradesh', 'Assam', 'Bihar', 'Chandigarh', 'Chhattisgarh', 'Dadra and Nagar Haveli and Daman and Diu', 'Delhi', 'Goa', 'Gujarat', 'Haryana', 'Himachal Pradesh', 'Jammu and Kashmir', 'Jharkhand', 'Karnataka', 'Kerala', 'Ladakh', 'Lakshadweep', 'Madhya Pradesh', 'Maharashtra', 'Manipur', 'Meghalaya', 'Mizoram', 'Nagaland', 'Odisha', 'Puducherry', 'Punjab', 'Rajasthan', 'Sikkim', 'Tamil Nadu', 'Telangana', 'Tripura', 'Uttar Pradesh', 'Uttarakhand', 'West Bengal'].map(st => (
                              <option key={st} value={st}>{st}</option>
                            ))}
                          </select>
                        </div>

                        <div className="max-h-[600px] overflow-y-auto space-y-2 pr-1">
                          {salesPurchaseInvoices.map((inv, idx) => {
                            const isSelected = selectedInvoiceIdx === idx;
                            return (
                              <button
                                key={idx}
                                type="button"
                                onClick={() => setSelectedInvoiceIdx(idx)}
                                className={`w-full text-left p-3.5 rounded-xl border-2 transition-all flex flex-col gap-1.5 ${
                                  isSelected
                                    ? 'border-zinc-900 bg-zinc-50'
                                    : 'border-zinc-100 hover:border-zinc-300 bg-white'
                                }`}
                              >
                                <div className="flex items-center justify-between gap-1">
                                  <span className="font-mono text-xs font-bold truncate text-zinc-900 max-w-[110px]">
                                    #{inv.invoiceNo || 'No No'}
                                  </span>
                                  <span className={`text-[10px] font-bold px-2 py-0.5 rounded-full ${
                                    inv.voucherType === 'Sales' ? 'bg-indigo-50 text-indigo-700' : 'bg-amber-50 text-amber-700'
                                  }`}>
                                    {inv.voucherType}
                                  </span>
                                </div>
                                <div className="flex justify-between items-center text-[11px] text-zinc-500">
                                  <span>{inv.invoiceDate}</span>
                                  <span className="font-semibold text-zinc-800">
                                    ₹{inv.invoiceTotal.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                  </span>
                                </div>
                                <div className="text-[10px] text-zinc-400 truncate max-w-full">
                                  {inv.partyLedger || 'No Party'}
                                </div>
                                {inv.errors.length > 0 ? (
                                  <span className="text-[9px] text-red-600 font-bold bg-red-50/50 px-1.5 py-0.5 rounded flex items-center gap-1 mt-1">
                                    ⚠️ {inv.errors.length} errors
                                  </span>
                                ) : inv.warnings.length > 0 ? (
                                  <span className="text-[9px] text-amber-600 font-bold bg-amber-50/50 px-1.5 py-0.5 rounded flex items-center gap-1 mt-1">
                                    ⚠️ {inv.warnings.length} warnings
                                  </span>
                                ) : (
                                  <span className="text-[9px] text-emerald-700 font-bold bg-emerald-50/50 px-1.5 py-0.5 rounded flex items-center gap-1 mt-1">
                                    ✓ Ready
                                  </span>
                                )}
                              </button>
                            );
                          })}
                        </div>
                      </div>

                      {/* Right Detail Panel */}
                      <div className="lg:col-span-3 space-y-6">
                        {activeInv ? (
                          <div className="space-y-6">
                            {/* Card Header & Basic fields */}
                            <div className="bg-zinc-50 p-5 rounded-2xl border border-zinc-200/60 space-y-4">
                              <div className="flex items-center justify-between">
                                <h3 className="font-bold text-zinc-800 text-sm">Invoice Configuration</h3>
                                <span className={`text-[11px] font-bold px-2.5 py-1 rounded-full ${
                                  activeInv.isValid ? 'bg-emerald-50 text-emerald-700 border border-emerald-100' : 'bg-red-50 text-red-700 border border-red-100'
                                }`}>
                                  {activeInv.isValid ? '✓ Valid' : '⚠️ Has Errors'}
                                </span>
                              </div>

                              <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
                                <div className="space-y-1">
                                  <label className="font-bold text-zinc-600">Invoice Date</label>
                                  <input
                                    type="text"
                                    value={activeInv.invoiceDate}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.invoiceDate = e.target.value; })}
                                    placeholder="DD/MM/YYYY"
                                    className="w-full p-2 bg-white border border-zinc-200 rounded-lg outline-none focus:ring-1 focus:ring-zinc-900"
                                  />
                                </div>
                                <div className="space-y-1">
                                  <label className="font-bold text-zinc-600">Invoice No</label>
                                  <input
                                    type="text"
                                    value={activeInv.invoiceNo}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.invoiceNo = e.target.value; })}
                                    className="w-full p-2 bg-white border border-zinc-200 rounded-lg outline-none focus:ring-1 focus:ring-zinc-900 font-mono"
                                  />
                                </div>
                                <div className="space-y-1">
                                  <label className="font-bold text-zinc-600">Place of Supply (State)</label>
                                  <select
                                    value={activeInv.placeOfSupply}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.placeOfSupply = e.target.value; })}
                                    className="w-full p-2 bg-white border border-zinc-200 rounded-lg outline-none focus:ring-1 focus:ring-zinc-900 font-medium text-xs"
                                  >
                                    {['Andaman and Nicobar Islands', 'Andhra Pradesh', 'Arunachal Pradesh', 'Assam', 'Bihar', 'Chandigarh', 'Chhattisgarh', 'Dadra and Nagar Haveli and Daman and Diu', 'Delhi', 'Goa', 'Gujarat', 'Haryana', 'Himachal Pradesh', 'Jammu and Kashmir', 'Jharkhand', 'Karnataka', 'Kerala', 'Ladakh', 'Lakshadweep', 'Madhya Pradesh', 'Maharashtra', 'Manipur', 'Meghalaya', 'Mizoram', 'Nagaland', 'Odisha', 'Puducherry', 'Punjab', 'Rajasthan', 'Sikkim', 'Tamil Nadu', 'Telangana', 'Tripura', 'Uttar Pradesh', 'Uttarakhand', 'West Bengal'].map(st => (
                                      <option key={st} value={st}>{st}</option>
                                    ))}
                                  </select>
                                </div>

                                <div className="space-y-1 md:col-span-1">
                                  <label className="font-bold text-zinc-600">Party Ledger</label>
                                  <select
                                    value={activeInv.partyLedger}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.partyLedger = e.target.value; })}
                                    className="w-full p-2 bg-white border border-zinc-200 rounded-lg outline-none focus:ring-1 focus:ring-zinc-900 font-medium"
                                  >
                                    <option value="">-- Select Party Ledger --</option>
                                    {ledgersList.map(l => (
                                      <option key={l} value={l}>{l}</option>
                                    ))}
                                  </select>
                                </div>

                                <div className="space-y-1 md:col-span-1">
                                  <label className="font-bold text-zinc-600">{activeInv.voucherType} Account Ledger</label>
                                  <select
                                    value={activeInv.salesPurchaseLedger}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.salesPurchaseLedger = e.target.value; })}
                                    className="w-full p-2 bg-white border border-zinc-200 rounded-lg outline-none focus:ring-1 focus:ring-zinc-900 font-medium"
                                  >
                                    <option value="">-- Select Ledger --</option>
                                    {ledgersList.map(l => (
                                      <option key={l} value={l}>{l}</option>
                                    ))}
                                  </select>
                                </div>

                                <div className="space-y-1 md:col-span-1">
                                  <label className="font-bold text-zinc-600">GST Mode</label>
                                  <select
                                    value={activeInv.gstMode}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.gstMode = e.target.value as 'Auto' | 'Manual'; })}
                                    className="w-full p-2 bg-white border border-zinc-200 rounded-lg outline-none focus:ring-1 focus:ring-zinc-900 font-medium"
                                  >
                                    <option value="Auto">Auto (Auto-Calculated from Pos)</option>
                                    <option value="Manual">Manual (Custom Ledger/Amounts)</option>
                                  </select>
                                </div>
                              </div>
                            </div>

                            {/* Dispatch, Transport & Party details in grid */}
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                              <div className="p-4 border border-zinc-100 rounded-xl bg-white space-y-3">
                                <h4 className="font-bold text-xs text-zinc-800 flex items-center gap-1.5 border-b border-zinc-50 pb-2">
                                  <Building className="w-4 h-4 text-zinc-400" />
                                  Dispatch & Transport Details
                                </h4>
                                <div className="grid grid-cols-2 gap-3.5 text-xs">
                                  <div className="space-y-1">
                                    <span className="font-semibold text-zinc-500">Dispatch Date</span>
                                    <input
                                      type="text"
                                      value={activeInv.dispatchDate}
                                      onChange={(e) => updateActiveInvoice(inv => { inv.dispatchDate = e.target.value; })}
                                      className="w-full p-1.5 border border-zinc-200 rounded outline-none text-xs"
                                      placeholder="DD/MM/YYYY"
                                    />
                                  </div>
                                  <div className="space-y-1">
                                    <span className="font-semibold text-zinc-500">Vehicle No</span>
                                    <input
                                      type="text"
                                      value={activeInv.vehicleNo}
                                      onChange={(e) => updateActiveInvoice(inv => { inv.vehicleNo = e.target.value; })}
                                      className="w-full p-1.5 border border-zinc-200 rounded outline-none text-xs font-mono"
                                      placeholder="MH-12-XX-1234"
                                    />
                                  </div>
                                  <div className="space-y-1">
                                    <span className="font-semibold text-zinc-500">Transporter Name</span>
                                    <input
                                      type="text"
                                      value={activeInv.transporterName}
                                      onChange={(e) => updateActiveInvoice(inv => { inv.transporterName = e.target.value; })}
                                      className="w-full p-1.5 border border-zinc-200 rounded outline-none text-xs"
                                    />
                                  </div>
                                  <div className="space-y-1">
                                    <span className="font-semibold text-zinc-500">Transporter GSTIN</span>
                                    <input
                                      type="text"
                                      value={activeInv.transporterGSTIN}
                                      onChange={(e) => updateActiveInvoice(inv => { inv.transporterGSTIN = e.target.value; })}
                                      className="w-full p-1.5 border border-zinc-200 rounded outline-none text-xs font-mono"
                                    />
                                  </div>
                                  <div className="space-y-1">
                                    <span className="font-semibold text-zinc-500">Delivery Note No</span>
                                    <input
                                      type="text"
                                      value={activeInv.deliveryNoteNo}
                                      onChange={(e) => updateActiveInvoice(inv => { inv.deliveryNoteNo = e.target.value; })}
                                      className="w-full p-1.5 border border-zinc-200 rounded outline-none text-xs"
                                    />
                                  </div>
                                  <div className="space-y-1">
                                    <span className="font-semibold text-zinc-500">Bilty LR No</span>
                                    <input
                                      type="text"
                                      value={activeInv.biltyLRNo}
                                      onChange={(e) => updateActiveInvoice(inv => { inv.biltyLRNo = e.target.value; })}
                                      className="w-full p-1.5 border border-zinc-200 rounded outline-none text-xs"
                                    />
                                  </div>
                                  <div className="space-y-1">
                                    <span className="font-semibold text-zinc-500">Mode of Transport</span>
                                    <input
                                      type="text"
                                      value={activeInv.modeOfTransport}
                                      onChange={(e) => updateActiveInvoice(inv => { inv.modeOfTransport = e.target.value; })}
                                      className="w-full p-1.5 border border-zinc-200 rounded outline-none text-xs"
                                      placeholder="Road/Air/Rail"
                                    />
                                  </div>
                                  <div className="space-y-1">
                                    <span className="font-semibold text-zinc-500">Eway Bill No</span>
                                    <input
                                      type="text"
                                      value={activeInv.ewayBillNo}
                                      onChange={(e) => updateActiveInvoice(inv => { inv.ewayBillNo = e.target.value; })}
                                      className="w-full p-1.5 border border-zinc-200 rounded outline-none text-xs font-mono"
                                    />
                                  </div>
                                </div>
                              </div>

                              <div className="p-4 border border-zinc-100 rounded-xl bg-white space-y-3">
                                <h4 className="font-bold text-xs text-zinc-800 flex items-center gap-1.5 border-b border-zinc-50 pb-2">
                                  <Building className="w-4 h-4 text-zinc-400" />
                                  Party Address & Billing Info
                                </h4>
                                <div className="grid grid-cols-2 gap-3.5 text-xs">
                                  <div className="space-y-1 col-span-2">
                                    <span className="font-semibold text-zinc-500">Party GSTIN</span>
                                    <input
                                      type="text"
                                      value={activeInv.partyGSTIN}
                                      onChange={(e) => updateActiveInvoice(inv => { inv.partyGSTIN = e.target.value; })}
                                      className="w-full p-1.5 border border-zinc-200 rounded outline-none text-xs font-mono"
                                      placeholder="GSTIN"
                                    />
                                  </div>
                                  <div className="space-y-1 col-span-2">
                                    <span className="font-semibold text-zinc-500">Address Line 1</span>
                                    <input
                                      type="text"
                                      value={activeInv.partyAddress1}
                                      onChange={(e) => updateActiveInvoice(inv => { inv.partyAddress1 = e.target.value; })}
                                      className="w-full p-1.5 border border-zinc-200 rounded outline-none text-xs"
                                    />
                                  </div>
                                  <div className="space-y-1 col-span-2">
                                    <span className="font-semibold text-zinc-500">Address Line 2</span>
                                    <input
                                      type="text"
                                      value={activeInv.partyAddress2}
                                      onChange={(e) => updateActiveInvoice(inv => { inv.partyAddress2 = e.target.value; })}
                                      className="w-full p-1.5 border border-zinc-200 rounded outline-none text-xs"
                                    />
                                  </div>
                                  <div className="space-y-1 col-span-2">
                                    <span className="font-semibold text-zinc-500">Party State</span>
                                    <input
                                      type="text"
                                      value={activeInv.partyState}
                                      onChange={(e) => updateActiveInvoice(inv => { inv.partyState = e.target.value; })}
                                      className="w-full p-1.5 border border-zinc-200 rounded outline-none text-xs"
                                    />
                                  </div>
                                  <div className="space-y-1 col-span-2">
                                    <span className="font-semibold text-zinc-500">Party Registration Type</span>
                                    <select
                                      value={activeInv.partyRegistrationType || 'Regular'}
                                      onChange={(e) => updateActiveInvoice(inv => { inv.partyRegistrationType = e.target.value; })}
                                      className="w-full p-1.5 border border-zinc-200 rounded outline-none text-xs bg-white"
                                    >
                                      <option value="Regular">Regular</option>
                                      <option value="Unregistered">Unregistered</option>
                                      <option value="Composition">Composition</option>
                                      <option value="Consumer">Consumer</option>
                                    </select>
                                  </div>
                                </div>
                              </div>
                            </div>

                            {/* Inventory Items table */}
                            <div className="border border-zinc-200 rounded-xl overflow-hidden">
                              <div className="p-3 bg-zinc-50 border-b border-zinc-200 flex items-center justify-between">
                                <h4 className="font-bold text-xs text-zinc-700">Inventory Items ({activeInv.items.length})</h4>
                                <button
                                  type="button"
                                  onClick={() => updateActiveInvoice(inv => {
                                    inv.items.push({
                                      stockItem: '',
                                      description: '',
                                      quantity: 1,
                                      unit: 'NOS',
                                      rate: 0,
                                      itemAmount: 0,
                                      discountPercent: 0,
                                      discountAmount: 0,
                                      taxableValue: 0,
                                      hsn: '',
                                      gstRate: 18,
                                      cgstLedger: '',
                                      cgstAmount: 0,
                                      sgstLedger: '',
                                      sgstAmount: 0,
                                      igstLedger: '',
                                      igstAmount: 0,
                                      gstRateSource: 'User Entered'
                                    });
                                  })}
                                  className="text-[11px] font-bold text-zinc-900 bg-white border border-zinc-200 hover:bg-zinc-50 px-2.5 py-1.5 rounded-lg transition-colors"
                                >
                                  + Add Item Line
                                </button>
                              </div>

                              <div className="overflow-x-auto">
                                <table className="w-full text-left text-xs border-collapse">
                                  <thead>
                                    <tr className="bg-zinc-50/50 border-b border-zinc-200 text-zinc-600 font-bold">
                                      <th className="p-2.5">Stock Item Name</th>
                                      <th className="p-2.5 w-16 text-center">Qty</th>
                                      <th className="p-2.5 w-16 text-center">Unit</th>
                                      <th className="p-2.5 w-20 text-right">Rate</th>
                                      <th className="p-2.5 w-24 text-right">Amount</th>
                                      <th className="p-2.5 w-16 text-right">Disc %</th>
                                      <th className="p-2.5 w-24 text-right">Taxable</th>
                                      <th className="p-2.5 w-16 text-center">GST %</th>
                                      <th className="p-2.5 w-12 text-center"></th>
                                    </tr>
                                  </thead>
                                  <tbody className="divide-y divide-zinc-100">
                                    {activeInv.items.map((item, itemIdx) => {
                                      const isStockInMaster = stockItemsList.some(s => s.toLowerCase() === item.stockItem.toLowerCase());
                                      return (
                                        <React.Fragment key={itemIdx}>
                                          <tr>
                                            <td className="p-2.5 min-w-[150px]">
                                              <select
                                                value={item.stockItem}
                                                onChange={(e) => updateActiveInvoice(inv => {
                                                  const it = inv.items[itemIdx];
                                                  it.stockItem = e.target.value;
                                                  // Auto autofill master values
                                                  const sDetails = tallyContext?.stockItemDetails?.find(sm => sm.itemName === e.target.value);
                                                  if (sDetails) {
                                                    it.unit = sDetails.unit || 'NOS';
                                                    it.hsn = sDetails.hsn || '';
                                                    if (sDetails.gstRate) {
                                                      it.gstRate = parseFloat(sDetails.gstRate);
                                                      it.gstRateSource = 'Stock Master';
                                                    }
                                                  }
                                                })}
                                                className="w-full p-1.5 bg-white border border-zinc-200 rounded font-medium text-xs focus:ring-1 focus:ring-zinc-900"
                                              >
                                                <option value="">-- Select Item --</option>
                                                {stockItemsList.map(st => (
                                                  <option key={st} value={st}>{st}</option>
                                                ))}
                                                {item.stockItem && !isStockInMaster && (
                                                  <option value={item.stockItem}>{item.stockItem} (not in masters)</option>
                                                )}
                                              </select>
                                              <input
                                                type="text"
                                                value={item.description}
                                                onChange={(e) => updateActiveInvoice(inv => { inv.items[itemIdx].description = e.target.value; })}
                                                placeholder="Description / Batch info..."
                                                className="w-full p-1 mt-1 text-[11px] border border-zinc-100 rounded text-zinc-500 placeholder-zinc-300"
                                              />
                                              {!isStockInMaster && item.stockItem && (
                                                <span className="text-[10px] text-amber-600 block mt-0.5">⚠️ Stock item not found in Masters</span>
                                              )}
                                            </td>
                                            <td className="p-2.5">
                                              <input
                                                type="number"
                                                value={item.quantity}
                                                onChange={(e) => updateActiveInvoice(inv => {
                                                  const it = inv.items[itemIdx];
                                                  it.quantity = parseFloat(e.target.value) || 0;
                                                  it.itemAmount = it.quantity * it.rate;
                                                  it.discountAmount = it.itemAmount * (it.discountPercent / 100);
                                                  it.taxableValue = it.itemAmount - it.discountAmount;
                                                })}
                                                className="w-full p-1.5 border border-zinc-200 rounded text-center text-xs"
                                              />
                                            </td>
                                            <td className="p-2.5">
                                              <select
                                                value={item.unit}
                                                onChange={(e) => updateActiveInvoice(inv => { inv.items[itemIdx].unit = e.target.value; })}
                                                className="w-full p-1.5 border border-zinc-200 rounded text-center text-xs font-medium"
                                              >
                                                {unitsList.map(u => (
                                                  <option key={u} value={u}>{u}</option>
                                                ))}
                                              </select>
                                            </td>
                                            <td className="p-2.5">
                                              <input
                                                type="number"
                                                value={item.rate}
                                                onChange={(e) => updateActiveInvoice(inv => {
                                                  const it = inv.items[itemIdx];
                                                  it.rate = parseFloat(e.target.value) || 0;
                                                  it.itemAmount = it.quantity * it.rate;
                                                  it.discountAmount = it.itemAmount * (it.discountPercent / 100);
                                                  it.taxableValue = it.itemAmount - it.discountAmount;
                                                })}
                                                className="w-full p-1.5 border border-zinc-200 rounded text-right text-xs"
                                              />
                                            </td>
                                            <td className="p-2.5 text-right font-semibold text-zinc-800">
                                              ₹{item.itemAmount.toFixed(2)}
                                            </td>
                                            <td className="p-2.5">
                                              <input
                                                type="number"
                                                value={item.discountPercent}
                                                onChange={(e) => updateActiveInvoice(inv => {
                                                  const it = inv.items[itemIdx];
                                                  it.discountPercent = parseFloat(e.target.value) || 0;
                                                  it.discountAmount = it.itemAmount * (it.discountPercent / 100);
                                                  it.taxableValue = it.itemAmount - it.discountAmount;
                                                })}
                                                className="w-full p-1.5 border border-zinc-200 rounded text-right text-xs"
                                              />
                                            </td>
                                            <td className="p-2.5">
                                              <input
                                                type="number"
                                                value={item.taxableValue}
                                                onChange={(e) => updateActiveInvoice(inv => {
                                                  const it = inv.items[itemIdx];
                                                  it.taxableValue = parseFloat(e.target.value) || 0;
                                                })}
                                                className="w-full p-1.5 border border-zinc-200 rounded text-right text-xs font-semibold"
                                              />
                                            </td>
                                            <td className="p-2.5">
                                              <div className="space-y-1">
                                                <input
                                                  type="number"
                                                  value={item.gstRate}
                                                  onChange={(e) => updateActiveInvoice(inv => {
                                                    inv.items[itemIdx].gstRate = parseFloat(e.target.value) || 0;
                                                    inv.items[itemIdx].gstRateSource = 'User Entered';
                                                  })}
                                                  className="w-full p-1.5 border border-zinc-200 rounded text-center text-xs"
                                                />
                                                <span className={`text-[9px] block text-center font-bold uppercase rounded-full px-1 py-0.2 ${
                                                  item.gstRateSource === 'Stock Master' ? 'bg-emerald-50 text-emerald-600' : 'bg-zinc-100 text-zinc-600'
                                                }`}>
                                                  {item.gstRateSource === 'Stock Master' ? 'Master' : 'User'}
                                                </span>
                                              </div>
                                            </td>
                                            <td className="p-2.5 text-center">
                                              <button
                                                type="button"
                                                onClick={() => updateActiveInvoice(inv => {
                                                  inv.items.splice(itemIdx, 1);
                                                })}
                                                className="text-red-500 hover:text-red-700 transition-colors font-bold text-xs"
                                              >
                                                ✕
                                              </button>
                                            </td>
                                          </tr>

                                          {/* Custom manual taxes adjustment block or preview block */}
                                          <tr className="bg-zinc-50/20 border-b border-zinc-200/50">
                                            <td colSpan={9} className="p-2 text-[11px] text-zinc-500">
                                              <div className="flex flex-wrap gap-4 items-center pl-4 pb-2">
                                                <span className="font-bold text-[10px] text-zinc-400 uppercase tracking-wider">GST Distribution:</span>
                                                {activeInv.gstMode === 'Auto' ? (
                                                  <div className="flex gap-4 items-center text-zinc-700">
                                                    {item.cgstAmount > 0 && (
                                                      <span>
                                                        CGST ({item.gstRate/2}%): <strong className="font-semibold text-zinc-900">{item.cgstLedger || 'Auto-matched'}</strong> = ₹{item.cgstAmount.toFixed(2)}
                                                      </span>
                                                    )}
                                                    {item.sgstAmount > 0 && (
                                                      <span>
                                                        SGST ({item.gstRate/2}%): <strong className="font-semibold text-zinc-900">{item.sgstLedger || 'Auto-matched'}</strong> = ₹{item.sgstAmount.toFixed(2)}
                                                      </span>
                                                    )}
                                                    {item.igstAmount > 0 && (
                                                      <span>
                                                        IGST ({item.gstRate}%): <strong className="font-semibold text-zinc-900">{item.igstLedger || 'Auto-matched'}</strong> = ₹{item.igstAmount.toFixed(2)}
                                                      </span>
                                                    )}
                                                  </div>
                                                ) : (
                                                  <div className="flex flex-wrap gap-3.5 items-center">
                                                    {/* Manual Tax fields inputs */}
                                                    <div className="flex items-center gap-1 text-[11px]">
                                                      <span>CGST Ledger:</span>
                                                      <select
                                                        value={item.cgstLedger}
                                                        onChange={(e) => updateActiveInvoice(inv => { inv.items[itemIdx].cgstLedger = e.target.value; })}
                                                        className="p-1 border border-zinc-200 bg-white rounded text-[11px] max-w-[130px]"
                                                      >
                                                        <option value="">-- None --</option>
                                                        {ledgersList.map(l => <option key={l} value={l}>{l}</option>)}
                                                      </select>
                                                      <input
                                                        type="number"
                                                        value={item.cgstAmount}
                                                        onChange={(e) => updateActiveInvoice(inv => { inv.items[itemIdx].cgstAmount = parseFloat(e.target.value) || 0; })}
                                                        className="p-1 border border-zinc-200 rounded text-right text-[11px] w-16"
                                                      />
                                                    </div>

                                                    <div className="flex items-center gap-1 text-[11px]">
                                                      <span>SGST Ledger:</span>
                                                      <select
                                                        value={item.sgstLedger}
                                                        onChange={(e) => updateActiveInvoice(inv => { inv.items[itemIdx].sgstLedger = e.target.value; })}
                                                        className="p-1 border border-zinc-200 bg-white rounded text-[11px] max-w-[130px]"
                                                      >
                                                        <option value="">-- None --</option>
                                                        {ledgersList.map(l => <option key={l} value={l}>{l}</option>)}
                                                      </select>
                                                      <input
                                                        type="number"
                                                        value={item.sgstAmount}
                                                        onChange={(e) => updateActiveInvoice(inv => { inv.items[itemIdx].sgstAmount = parseFloat(e.target.value) || 0; })}
                                                        className="p-1 border border-zinc-200 rounded text-right text-[11px] w-16"
                                                      />
                                                    </div>

                                                    <div className="flex items-center gap-1 text-[11px]">
                                                      <span>IGST Ledger:</span>
                                                      <select
                                                        value={item.igstLedger}
                                                        onChange={(e) => updateActiveInvoice(inv => { inv.items[itemIdx].igstLedger = e.target.value; })}
                                                        className="p-1 border border-zinc-200 bg-white rounded text-[11px] max-w-[130px]"
                                                      >
                                                        <option value="">-- None --</option>
                                                        {ledgersList.map(l => <option key={l} value={l}>{l}</option>)}
                                                      </select>
                                                      <input
                                                        type="number"
                                                        value={item.igstAmount}
                                                        onChange={(e) => updateActiveInvoice(inv => { inv.items[itemIdx].igstAmount = parseFloat(e.target.value) || 0; })}
                                                        className="p-1 border border-zinc-200 rounded text-right text-[11px] w-16"
                                                      />
                                                    </div>
                                                  </div>
                                                )}
                                              </div>
                                            </td>
                                          </tr>
                                        </React.Fragment>
                                      );
                                    })}
                                  </tbody>
                                </table>
                              </div>
                            </div>

                            {/* Additional Charges / Expenses / Round-off */}
                            <div className="bg-zinc-50/50 p-5 rounded-2xl border border-zinc-200/60 space-y-4 text-xs">
                              <h4 className="font-bold text-xs text-zinc-800 border-b border-zinc-100 pb-2">Additional Accounting Ledger Lines</h4>
                              <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                                <div className="space-y-1">
                                  <span className="font-bold text-zinc-500">Freight Ledger</span>
                                  <select
                                    value={activeInv.freightLedger}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.freightLedger = e.target.value; })}
                                    className="w-full p-1.5 bg-white border border-zinc-200 rounded text-xs"
                                  >
                                    <option value="">-- Select Freight Ldg --</option>
                                    {ledgersList.map(l => <option key={l} value={l}>{l}</option>)}
                                  </select>
                                  <input
                                    type="number"
                                    value={activeInv.freightAmount}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.freightAmount = parseFloat(e.target.value) || 0; })}
                                    className="w-full p-1.5 border border-zinc-200 rounded text-right mt-1 font-mono text-xs"
                                    placeholder="Amount"
                                  />
                                </div>

                                <div className="space-y-1">
                                  <span className="font-bold text-zinc-500">Packing Ledger</span>
                                  <select
                                    value={activeInv.packingLedger}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.packingLedger = e.target.value; })}
                                    className="w-full p-1.5 bg-white border border-zinc-200 rounded text-xs"
                                  >
                                    <option value="">-- Select Packing Ldg --</option>
                                    {ledgersList.map(l => <option key={l} value={l}>{l}</option>)}
                                  </select>
                                  <input
                                    type="number"
                                    value={activeInv.packingAmount}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.packingAmount = parseFloat(e.target.value) || 0; })}
                                    className="w-full p-1.5 border border-zinc-200 rounded text-right mt-1 font-mono text-xs"
                                    placeholder="Amount"
                                  />
                                </div>

                                <div className="space-y-1">
                                  <span className="font-bold text-zinc-500">Loading Ledger</span>
                                  <select
                                    value={activeInv.loadingLedger}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.loadingLedger = e.target.value; })}
                                    className="w-full p-1.5 bg-white border border-zinc-200 rounded text-xs"
                                  >
                                    <option value="">-- Select Loading Ldg --</option>
                                    {ledgersList.map(l => <option key={l} value={l}>{l}</option>)}
                                  </select>
                                  <input
                                    type="number"
                                    value={activeInv.loadingAmount}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.loadingAmount = parseFloat(e.target.value) || 0; })}
                                    className="w-full p-1.5 border border-zinc-200 rounded text-right mt-1 font-mono text-xs"
                                    placeholder="Amount"
                                  />
                                </div>

                                <div className="space-y-1">
                                  <span className="font-bold text-zinc-500">Insurance Ledger</span>
                                  <select
                                    value={activeInv.insuranceLedger}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.insuranceLedger = e.target.value; })}
                                    className="w-full p-1.5 bg-white border border-zinc-200 rounded text-xs"
                                  >
                                    <option value="">-- Select Insurance Ldg --</option>
                                    {ledgersList.map(l => <option key={l} value={l}>{l}</option>)}
                                  </select>
                                  <input
                                    type="number"
                                    value={activeInv.insuranceAmount}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.insuranceAmount = parseFloat(e.target.value) || 0; })}
                                    className="w-full p-1.5 border border-zinc-200 rounded text-right mt-1 font-mono text-xs"
                                    placeholder="Amount"
                                  />
                                </div>

                                <div className="space-y-1">
                                  <span className="font-bold text-zinc-500">Bill Discount Ledger</span>
                                  <select
                                    value={activeInv.discountLedger}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.discountLedger = e.target.value; })}
                                    className="w-full p-1.5 bg-white border border-zinc-200 rounded text-xs"
                                  >
                                    <option value="">-- Select Discount Ldg --</option>
                                    {ledgersList.map(l => <option key={l} value={l}>{l}</option>)}
                                  </select>
                                  <input
                                    type="number"
                                    value={activeInv.billDiscountAmount}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.billDiscountAmount = parseFloat(e.target.value) || 0; })}
                                    className="w-full p-1.5 border border-zinc-200 rounded text-right mt-1 font-mono text-xs text-amber-700"
                                    placeholder="Discount Value"
                                  />
                                </div>

                                <div className="space-y-1">
                                  <span className="font-bold text-zinc-500">Round Off Ledger</span>
                                  <select
                                    value={activeInv.roundOffLedger}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.roundOffLedger = e.target.value; })}
                                    className="w-full p-1.5 bg-white border border-zinc-200 rounded text-xs"
                                  >
                                    <option value="">-- Select Roundoff Ldg --</option>
                                    {ledgersList.map(l => <option key={l} value={l}>{l}</option>)}
                                  </select>
                                  <input
                                    type="number"
                                    value={activeInv.roundOffAmount}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.roundOffAmount = parseFloat(e.target.value) || 0; })}
                                    className="w-full p-1.5 border border-zinc-200 rounded text-right mt-1 font-mono text-xs"
                                    placeholder="Amount (Auto)"
                                  />
                                </div>

                                <div className="space-y-1">
                                  <span className="font-bold text-zinc-500">Other Charges 1</span>
                                  <select
                                    value={activeInv.otherLedger1}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.otherLedger1 = e.target.value; })}
                                    className="w-full p-1.5 bg-white border border-zinc-200 rounded text-xs"
                                  >
                                    <option value="">-- Select Ledger --</option>
                                    {ledgersList.map(l => <option key={l} value={l}>{l}</option>)}
                                  </select>
                                  <input
                                    type="number"
                                    value={activeInv.otherAmount1}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.otherAmount1 = parseFloat(e.target.value) || 0; })}
                                    className="w-full p-1.5 border border-zinc-200 rounded text-right mt-1 font-mono text-xs"
                                  />
                                </div>

                                <div className="space-y-1">
                                  <span className="font-bold text-zinc-500">Other Charges 2</span>
                                  <select
                                    value={activeInv.otherLedger2}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.otherLedger2 = e.target.value; })}
                                    className="w-full p-1.5 bg-white border border-zinc-200 rounded text-xs"
                                  >
                                    <option value="">-- Select Ledger --</option>
                                    {ledgersList.map(l => <option key={l} value={l}>{l}</option>)}
                                  </select>
                                  <input
                                    type="number"
                                    value={activeInv.otherAmount2}
                                    onChange={(e) => updateActiveInvoice(inv => { inv.otherAmount2 = parseFloat(e.target.value) || 0; })}
                                    className="w-full p-1.5 border border-zinc-200 rounded text-right mt-1 font-mono text-xs"
                                  />
                                </div>
                              </div>
                            </div>

                            {/* Summary Calculations Footer Panel */}
                            <div className="bg-zinc-900 text-white p-6 rounded-2xl border border-zinc-800 flex flex-col md:flex-row md:items-center justify-between gap-6">
                              <div className="grid grid-cols-2 md:grid-cols-4 gap-6 text-xs flex-1">
                                <div>
                                  <span className="text-zinc-400 block mb-1">Total Taxable Value</span>
                                  <span className="font-mono text-base font-bold text-white">₹{activeInv.totalTaxableValue.toFixed(2)}</span>
                                </div>
                                <div>
                                  <span className="text-zinc-400 block mb-1">Total CGST + SGST</span>
                                  <span className="font-mono text-base font-bold text-white">₹{(activeInv.totalCGST + activeInv.totalSGST).toFixed(2)}</span>
                                </div>
                                <div>
                                  <span className="text-zinc-400 block mb-1">Total IGST</span>
                                  <span className="font-mono text-base font-bold text-white">₹{activeInv.totalIGST.toFixed(2)}</span>
                                </div>
                                <div>
                                  <span className="text-zinc-400 block mb-1">Additional Charges</span>
                                  <span className="font-mono text-base font-bold text-white">₹{activeInv.totalAdditionalCharges.toFixed(2)}</span>
                                </div>
                              </div>
                              <div className="border-t md:border-t-0 md:border-l border-zinc-800 pt-4 md:pt-0 md:pl-6 shrink-0 text-right">
                                <span className="text-zinc-400 block text-xs mb-1">Grand Invoice Total</span>
                                <span className="font-mono text-2xl font-black text-emerald-400">
                                  ₹{activeInv.invoiceTotal.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                </span>
                              </div>
                            </div>

                            {/* Alert Notifications list if any */}
                            {(activeInv.errors.length > 0 || activeInv.warnings.length > 0) && (
                              <div className="p-4 rounded-xl border space-y-2 bg-amber-50/25 border-amber-200/50 text-xs">
                                <h5 className="font-bold text-amber-800">Action Items & Validation Messages</h5>
                                <ul className="space-y-1.5 text-zinc-600 list-disc pl-5">
                                  {activeInv.errors.map((err, i) => (
                                    <li key={i} className="text-red-700 font-medium">{err}</li>
                                  ))}
                                  {activeInv.warnings.map((warn, i) => (
                                    <li key={i} className="text-amber-700">{warn}</li>
                                  ))}
                                </ul>
                              </div>
                            )}
                          </div>
                        ) : (
                          <div className="text-center py-20 text-zinc-400 border border-dashed border-zinc-200 rounded-2xl">
                            Select an invoice from the sidebar to view or customize.
                          </div>
                        )}
                      </div>
                    </div>

                    {/* Footer wizard controls */}
                    <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 pt-4 border-t border-zinc-100">
                      <div className="text-xs text-zinc-500">
                        Invoices are automatically compiled and validated according to strict Tally guidelines.
                      </div>
                      <div className="flex items-center gap-3 justify-end shrink-0">
                        <button
                          type="button"
                          onClick={() => setCurrentStep('upload')}
                          className="px-5 py-3 hover:bg-zinc-100 border border-zinc-200 text-zinc-700 rounded-xl text-xs font-bold transition-all flex items-center gap-2"
                        >
                          <RotateCcw className="w-4 h-4" />
                          Cancel & Back
                        </button>
                        <button
                          type="button"
                          onClick={generateSalesPurchaseXMLFromVerification}
                          disabled={isProcessing || salesPurchaseInvoices.some(i => !i.isValid)}
                          className="px-6 py-3 bg-zinc-900 hover:bg-zinc-850 text-white rounded-xl text-xs font-bold transition-all flex items-center gap-2 disabled:opacity-50 shadow-sm"
                        >
                          {isProcessing ? (
                            <>
                              <Loader2 className="w-4 h-4 animate-spin" />
                              Generating XML...
                            </>
                          ) : (
                            <>
                              <CheckCircle2 className="w-4 h-4" />
                              Confirm & Generate Tally XML
                            </>
                          )}
                        </button>
                      </div>
                    </div>
                  </motion.section>
                );
              })()}

              {currentStep === 'journal-verification' && (() => {
                const activeGroup = journalGroups[selectedJournalGroupIdx];
                const ledgersList = tallyContext?.ledgers || DEFAULT_LEDGERS;

                return (
                  <motion.section
                    key="journal-verification"
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, y: -10 }}
                    className="bg-white p-6 rounded-2xl border border-zinc-200 shadow-sm space-y-6"
                  >
                    <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-zinc-100 pb-5">
                      <div>
                        <h2 className="text-xl font-bold flex items-center gap-2 text-zinc-900">
                          <CheckCircle2 className="w-5 h-5 text-emerald-600" />
                          Journal Voucher Review and Verification
                        </h2>
                        <p className="text-xs text-zinc-500 mt-1 leading-relaxed">
                          Review and edit multi-line journal entries. All debit/credit totals must balance before generating the final XML.
                        </p>
                      </div>
                      <div className="flex flex-wrap gap-2">
                        <button
                          type="button"
                          onClick={() => generateJournalErrorExcel(journalGroups)}
                          className="px-3.5 py-2 border border-red-200 text-red-700 bg-red-50 hover:bg-red-100 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5"
                        >
                          <AlertCircle className="w-4 h-4" />
                          Download Error Log
                        </button>
                        <span className="text-xs bg-zinc-100 text-zinc-700 border border-zinc-200 px-3 py-1.5 rounded-lg font-bold flex items-center">
                          {journalGroups.length} Journal Vouchers Ready
                        </span>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                      {/* Left: Voucher List */}
                      <div className="lg:col-span-1 border border-zinc-200 rounded-xl overflow-hidden bg-zinc-50 flex flex-col max-h-[600px]">
                        <div className="p-3 bg-white border-b border-zinc-200 font-bold text-zinc-700 text-xs uppercase tracking-wider">
                          Voucher List
                        </div>
                        <div className="overflow-y-auto divide-y divide-zinc-200/60 flex-1">
                          {journalGroups.map((grp, idx) => {
                            const isSelected = idx === selectedJournalGroupIdx;
                            const hasErrors = !grp.isValid;
                            return (
                              <button
                                key={idx}
                                type="button"
                                onClick={() => setSelectedJournalGroupIdx(idx)}
                                className={`w-full p-3 text-left transition-all flex flex-col gap-1.5 hover:bg-zinc-100 ${
                                  isSelected ? 'bg-zinc-200/70 hover:bg-zinc-200/70 border-l-4 border-zinc-900 pl-2' : ''
                                }`}
                              >
                                <div className="flex items-center justify-between gap-2">
                                  <span className="font-bold text-zinc-800 text-xs">
                                    {grp.voucherNo || `Unnumbered (${idx + 1})`}
                                  </span>
                                  <span className="text-[10px] text-zinc-500 font-mono">
                                    {grp.voucherDate}
                                  </span>
                                </div>
                                <div className="flex items-center justify-between gap-2">
                                  <span className="text-[11px] text-zinc-600">
                                    Lines: {grp.lines.length}
                                  </span>
                                  <span className={`text-[10px] px-2 py-0.5 rounded-md font-bold ${
                                    hasErrors 
                                      ? 'bg-red-50 text-red-700 border border-red-100' 
                                      : 'bg-emerald-50 text-emerald-700 border border-emerald-100'
                                  }`}>
                                    {hasErrors ? 'Mismatch' : 'Balanced'}
                                  </span>
                                </div>
                                <div className="text-[10px] text-zinc-500 font-mono flex justify-between">
                                  <span>Dr: ₹{grp.totalDebit.toLocaleString('en-IN')}</span>
                                  <span>Cr: ₹{grp.totalCredit.toLocaleString('en-IN')}</span>
                                </div>
                              </button>
                            );
                          })}
                        </div>
                      </div>

                      {/* Right: Active Voucher Detail Form */}
                      <div className="lg:col-span-2 space-y-4">
                        {activeGroup ? (
                          <div className="border border-zinc-200 rounded-xl p-5 bg-white space-y-5">
                            {/* Header details */}
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                              <div>
                                <label className="text-[11px] font-bold text-zinc-500 block mb-1 uppercase tracking-wider">Voucher Date</label>
                                <input
                                  type="text"
                                  value={activeGroup.voucherDate}
                                  placeholder="DD-MM-YYYY"
                                  onChange={e => updateJournalGroupHeader(selectedJournalGroupIdx, { voucherDate: e.target.value })}
                                  className="w-full px-3 py-2 border border-zinc-200 rounded-xl focus:border-zinc-900 focus:ring-1 focus:ring-zinc-900 outline-none text-xs font-semibold text-zinc-800"
                                />
                              </div>
                              <div>
                                <label className="text-[11px] font-bold text-zinc-500 block mb-1 uppercase tracking-wider">Voucher No</label>
                                <input
                                  type="text"
                                  value={activeGroup.voucherNo}
                                  onChange={e => updateJournalGroupHeader(selectedJournalGroupIdx, { voucherNo: e.target.value })}
                                  className="w-full px-3 py-2 border border-zinc-200 rounded-xl focus:border-zinc-900 focus:ring-1 focus:ring-zinc-900 outline-none text-xs font-semibold text-zinc-800"
                                />
                              </div>
                            </div>

                            {/* Ledger lines table */}
                            <div className="space-y-3">
                              <div className="flex items-center justify-between pb-1 border-b border-zinc-100">
                                <h3 className="font-bold text-zinc-800 text-xs uppercase tracking-wider">Ledger Lines</h3>
                                <button
                                  type="button"
                                  onClick={() => addJournalLine(selectedJournalGroupIdx)}
                                  className="px-2.5 py-1 text-[11px] bg-zinc-900 text-white rounded-lg hover:bg-zinc-800 transition-colors font-bold"
                                >
                                  + Add Ledger Line
                                </button>
                              </div>

                              <div className="overflow-x-auto border border-zinc-200 rounded-xl">
                                <table className="w-full text-left border-collapse text-xs">
                                  <thead>
                                    <tr className="bg-zinc-50 border-b border-zinc-200 font-bold text-zinc-600">
                                      <th className="p-2.5 w-16 text-center">Dr/Cr</th>
                                      <th className="p-2.5 min-w-[200px]">Ledger Name</th>
                                      <th className="p-2.5 w-28">Amount</th>
                                      <th className="p-2.5 w-24">Cost Centre</th>
                                      <th className="p-2.5 w-24">Bill Ref</th>
                                      <th className="p-2.5 w-20">Remarks</th>
                                      <th className="p-2.5 w-12 text-center">Del</th>
                                    </tr>
                                  </thead>
                                  <tbody className="divide-y divide-zinc-100">
                                    {activeGroup.lines.map((line, lIdx) => (
                                      <tr key={lIdx} className={line.excluded ? 'opacity-40' : ''}>
                                        <td className="p-2">
                                          <select
                                            value={line.drCr}
                                            onChange={e => updateJournalLine(selectedJournalGroupIdx, lIdx, { drCr: e.target.value as any })}
                                            className="w-full p-1 border border-zinc-200 rounded-md bg-white text-xs font-bold"
                                          >
                                            <option value="Dr">Dr</option>
                                            <option value="Cr">Cr</option>
                                          </select>
                                        </td>
                                        <td className="p-2">
                                          <div className="relative">
                                            <input
                                              type="text"
                                              value={line.ledgerName}
                                              onChange={e => updateJournalLine(selectedJournalGroupIdx, lIdx, { ledgerName: e.target.value })}
                                              className="w-full px-2 py-1 border border-zinc-200 rounded-md text-xs font-semibold"
                                              placeholder="Search/type ledger..."
                                              list={`journal-ledgers-list-${selectedJournalGroupIdx}-${lIdx}`}
                                            />
                                            <datalist id={`journal-ledgers-list-${selectedJournalGroupIdx}-${lIdx}`}>
                                              {ledgersList.map(item => (
                                                <option key={item} value={item} />
                                              ))}
                                            </datalist>
                                            {line.ledgerName && !ledgersList.some(m => m.toLowerCase() === line.ledgerName.toLowerCase()) && (
                                              <span className="absolute right-2 top-1.5 text-[9px] bg-amber-50 text-amber-700 border border-amber-100 px-1.5 py-0.2 rounded font-bold">
                                                New
                                              </span>
                                            )}
                                          </div>
                                        </td>
                                        <td className="p-2">
                                          <input
                                            type="number"
                                            value={line.amount || ''}
                                            onChange={e => updateJournalLine(selectedJournalGroupIdx, lIdx, { amount: parseFloat(e.target.value) || 0 })}
                                            className="w-full px-2 py-1 border border-zinc-200 rounded-md text-xs font-semibold text-right"
                                          />
                                        </td>
                                        <td className="p-2">
                                          <input
                                            type="text"
                                            value={line.costCentre || ''}
                                            onChange={e => updateJournalLine(selectedJournalGroupIdx, lIdx, { costCentre: e.target.value })}
                                            className="w-full px-2 py-1 border border-zinc-200 rounded-md text-xs"
                                            placeholder="None"
                                          />
                                        </td>
                                        <td className="p-2">
                                          <input
                                            type="text"
                                            value={line.billReference || ''}
                                            onChange={e => updateJournalLine(selectedJournalGroupIdx, lIdx, { billReference: e.target.value })}
                                            className="w-full px-2 py-1 border border-zinc-200 rounded-md text-xs"
                                            placeholder="None"
                                          />
                                        </td>
                                        <td className="p-2">
                                          <input
                                            type="text"
                                            value={line.remarks || ''}
                                            onChange={e => updateJournalLine(selectedJournalGroupIdx, lIdx, { remarks: e.target.value })}
                                            className="w-full px-2 py-1 border border-zinc-200 rounded-md text-xs"
                                            placeholder="None"
                                          />
                                        </td>
                                        <td className="p-2 text-center">
                                          <button
                                            type="button"
                                            onClick={() => {
                                              const updatedLines = [...activeGroup.lines];
                                              updatedLines.splice(lIdx, 1);
                                              setJournalGroups(prev => {
                                                const copy = [...prev];
                                                copy[selectedJournalGroupIdx] = { ...copy[selectedJournalGroupIdx], lines: updatedLines };
                                                return copy;
                                              });
                                              setTimeout(() => {
                                                updateJournalLine(selectedJournalGroupIdx, 0, {});
                                              }, 10);
                                            }}
                                            className="p-1 text-red-500 hover:text-red-700 rounded transition-colors"
                                            title="Delete Line"
                                          >
                                            <Trash2 className="w-3.5 h-3.5" />
                                          </button>
                                        </td>
                                      </tr>
                                    ))}
                                  </tbody>
                                </table>
                              </div>
                            </div>

                            {/* Summary info */}
                            <div className="bg-zinc-50 border border-zinc-150 p-4 rounded-xl flex flex-col md:flex-row justify-between gap-4">
                              <div className="space-y-1">
                                <span className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider block">Voucher Totals</span>
                                <div className="flex gap-4 text-xs font-bold text-zinc-700">
                                  <div>Total Debit: <span className="text-zinc-900 font-mono">₹{activeGroup.totalDebit.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</span></div>
                                  <div>Total Credit: <span className="text-zinc-900 font-mono">₹{activeGroup.totalCredit.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</span></div>
                                </div>
                              </div>
                              <div className="text-right space-y-1">
                                <span className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider block">Difference</span>
                                <div className={`text-xs font-bold px-3 py-1 rounded-lg inline-block font-mono ${
                                  Math.abs(activeGroup.difference) <= 0.001
                                    ? 'bg-emerald-50 text-emerald-700 border border-emerald-200'
                                    : 'bg-red-50 text-red-700 border border-red-200'
                                }`}>
                                  ₹{activeGroup.difference.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                                </div>
                              </div>
                            </div>

                            {/* Error messaging list */}
                            {(activeGroup.errors.length > 0 || activeGroup.warnings.length > 0) && (
                              <div className="space-y-2 p-4 bg-zinc-50 border border-zinc-200 rounded-xl text-[11px]">
                                {activeGroup.errors.map((err, i) => (
                                  <div key={i} className="text-red-600 font-medium flex items-center gap-1.5">
                                    <span className="w-1.5 h-1.5 bg-red-500 rounded-full shrink-0"></span>
                                    {err}
                                  </div>
                                ))}
                                {activeGroup.warnings.map((warn, i) => (
                                  <div key={i} className="text-amber-700 font-medium flex items-center gap-1.5">
                                    <span className="w-1.5 h-1.5 bg-amber-500 rounded-full shrink-0"></span>
                                    {warn}
                                  </div>
                                ))}
                              </div>
                            )}

                            {/* Action Row */}
                            <div className="flex items-center justify-between pt-2 border-t border-zinc-100">
                              <span className="text-[11px] text-zinc-400">
                                Voucher Row Source: {activeGroup.lines.map(l => l.rowNo).join(', ')}
                              </span>
                              <button
                                type="button"
                                disabled={isProcessing || !journalGroups.some(g => g.isValid)}
                                onClick={() => checkMissingMastersAndProceedJournal(journalGroups)}
                                className="px-6 py-2.5 bg-zinc-950 hover:bg-zinc-800 disabled:opacity-40 text-white rounded-xl text-xs font-bold transition-all shadow-sm flex items-center gap-1.5"
                              >
                                {isProcessing ? (
                                  <>
                                    <Loader2 className="w-3.5 h-3.5 animate-spin" />
                                    Processing...
                                  </>
                                ) : (
                                  <>
                                    <CheckCircle2 className="w-3.5 h-3.5" />
                                    Confirm & Generate Journal XML
                                  </>
                                )}
                              </button>
                            </div>
                          </div>
                        ) : (
                          <div className="h-48 border border-zinc-200 rounded-xl bg-zinc-50 flex items-center justify-center text-xs text-zinc-400">
                            Select a voucher from the list to view and edit details
                          </div>
                        )}
                      </div>
                    </div>
                  </motion.section>
                );
              })()}

              {currentStep === 'complete' && (
                <motion.section 
                  key="complete"
                  initial={{ opacity: 0, scale: 0.95 }}
                  animate={{ opacity: 1, scale: 1 }}
                  className="bg-zinc-900 text-white p-12 rounded-2xl shadow-2xl text-center relative overflow-hidden"
                >
                  <div className="absolute top-0 right-0 p-8 opacity-10">
                    <FileCode className="w-64 h-64" />
                  </div>
                  <div className="relative z-10">
                    <div className="w-20 h-20 bg-green-500 rounded-full flex items-center justify-center mx-auto mb-6 shadow-lg shadow-green-500/20">
                      <CheckCircle2 className="w-10 h-10 text-white" />
                    </div>
                    <h2 className="text-3xl font-bold mb-2">Conversion Successful!</h2>
                    <p className="text-zinc-400 mb-8 max-w-md mx-auto">
                      {conversions[0]?.voucherType === 'SalesPurchase'
                        ? 'Your Sales & Purchase item invoices have been successfully grouped, calculated, and converted to Tally-compliant XML format.'
                        : importType === 'Voucher' 
                          ? 'Your bank statement has been successfully mapped and converted to Tally XML format.'
                          : `Your Tally ${importType} Masters have been successfully validated and converted to Tally XML format.`
                      }
                    </p>
                    
                    <div className="flex flex-col sm:flex-row gap-4 justify-center">
                      <button 
                        onClick={() => {
                          const lastConv = conversions[0];
                          if (lastConv?.xmlContent) downloadXML(lastConv.xmlContent, lastConv.fileName);
                        }}
                        className="flex items-center justify-center gap-2 bg-white text-zinc-900 py-3 px-8 rounded-xl font-bold hover:bg-zinc-100 transition-colors"
                      >
                        <FileCode className="w-5 h-5" />
                        Download XML
                      </button>

                      {getAppMode() === 'desktop-offline' && isDirectTallyAvailable() && tallyStatus === 'Connected' && (
                        <button 
                          onClick={() => {
                            const lastConv = conversions[0];
                            if (lastConv?.xmlContent) {
                              const type = lastConv.voucherType || 'Voucher';
                              const count = lastConv.voucherCount || 0;
                              initiateTallyPush(
                                lastConv.xmlContent,
                                type,
                                count,
                                0
                              );
                            }
                          }}
                          className="flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white py-3 px-8 rounded-xl font-bold transition-colors"
                        >
                          <RefreshCw className="w-5 h-5" />
                          Push to Tally
                        </button>
                      )}
                      
                      {conversions[0]?.voucherType === 'SalesPurchase' ? null : importType === 'Voucher' ? (
                        <button 
                          onClick={() => {
                            const lastConv = conversions[0];
                            if (aiMappedTransactions.length > 0) {
                              downloadMappedExcel(aiMappedTransactions, lastConv?.fileName || 'Statement.xlsx');
                            }
                          }}
                          className="flex items-center justify-center gap-2 bg-zinc-800 text-white py-3 px-8 rounded-xl font-bold hover:bg-zinc-700 transition-colors border border-zinc-700"
                        >
                          <FileSpreadsheet className="w-5 h-5" />
                          Download Excel
                        </button>
                      ) : (
                        <button 
                          onClick={downloadValidationReport}
                          className="flex items-center justify-center gap-2 bg-zinc-800 text-white py-3 px-8 rounded-xl font-bold hover:bg-zinc-700 transition-colors border border-zinc-700"
                        >
                          <FileSpreadsheet className="w-5 h-5" />
                          Download Validation Report
                        </button>
                      )}
                      
                      <button 
                        onClick={() => setCurrentStep('upload')}
                        className="flex items-center justify-center gap-2 bg-white/10 text-white py-3 px-8 rounded-xl font-medium hover:bg-white/20 transition-colors border border-white/10"
                      >
                        Convert Another
                      </button>
                    </div>
                  </div>
                </motion.section>
              )}
            </AnimatePresence>

            {/* Local Mapping Preview (Only show on upload step) */}
            {currentStep === 'upload' && (
              <section className="bg-zinc-900 text-white p-6 rounded-2xl shadow-xl overflow-hidden relative">
                <div className="absolute top-0 right-0 p-4 opacity-10">
                  <Sparkles className="w-32 h-32" />
                </div>
                <div className="relative z-10">
                  <div className="flex items-center gap-2 mb-4">
                    <Sparkles className="w-5 h-5 text-amber-400" />
                    <h2 className="text-xl font-bold">Local Mapping Engine</h2>
                  </div>
                  <p className="text-zinc-400 mb-6">Our local deterministic mapping engine automatically parses your column headers and maps them to Tally's XML schema.</p>
                  <div className="grid grid-cols-2 gap-4">
                    <div className="p-3 bg-white/5 rounded-lg border border-white/10">
                      <p className="text-xs text-zinc-500 uppercase tracking-wider mb-1">Detected Columns</p>
                      <p className="font-mono text-sm">Date, Particulars, Vch Type, Vch No, Amount</p>
                    </div>
                    <div className="p-3 bg-white/5 rounded-lg border border-white/10">
                      <p className="text-xs text-zinc-500 uppercase tracking-wider mb-1">Tally Schema</p>
                      <p className="font-mono text-sm">DATE, PARTYNAME, VOUCHERTYPENAME, VOUCHERNUMBER, AMOUNT</p>
                    </div>
                  </div>
                </div>
              </section>
            )}
          </div>

            {/* Right: History & Templates */}
            <div className="space-y-6">
              <section className="bg-white p-6 rounded-2xl border border-zinc-200 shadow-sm">
                <h2 className="text-xl font-bold mb-4 flex items-center gap-2">
                  <FileSpreadsheet className="w-5 h-5 text-zinc-900" />
                  Excel Templates
                </h2>
                <p className="text-sm text-zinc-500 mb-2 font-medium">Voucher Templates:</p>
                <div className="grid grid-cols-2 gap-2 mb-4">
                  {['Payment', 'Receipt', 'Contra', 'Journal'].map(type => (
                    <button 
                      key={type}
                      onClick={() => downloadTemplate(type)}
                      className="flex items-center justify-center gap-2 p-2 bg-zinc-50 hover:bg-zinc-100 border border-zinc-200 rounded-lg text-xs font-medium transition-colors"
                    >
                      <Download className="w-3 h-3" />
                      {type}
                    </button>
                  ))}
                </div>
                <p className="text-sm text-zinc-500 mb-2 font-medium">Sales & Purchase Templates:</p>
                <div className="grid grid-cols-2 gap-2 mb-4">
                  {[
                    { key: 'Sales Itemwise', label: 'Sales (Itemwise)' },
                    { key: 'Sales Voucherwise', label: 'Sales (Voucherwise)' },
                    { key: 'Purchase Itemwise', label: 'Purchase (Itemwise)' },
                    { key: 'Purchase Voucherwise', label: 'Purchase (Voucherwise)' }
                  ].map(item => (
                    <button 
                      key={item.key}
                      onClick={() => downloadTemplate(item.key)}
                      className="flex items-center justify-center gap-2 p-2 bg-zinc-50 hover:bg-zinc-100 border border-zinc-200 rounded-lg text-[10px] font-medium transition-colors"
                    >
                      <Download className="w-3 h-3 text-zinc-500" />
                      {item.label}
                    </button>
                  ))}
                </div>
                <p className="text-sm text-zinc-500 mb-2 font-medium">Master Templates:</p>
                <div className="grid grid-cols-2 gap-2">
                  {[
                    { key: 'Ledger', label: 'Ledger' },
                    { key: 'StockItem', label: 'Stock Item' },
                    { key: 'StockGroup', label: 'Stock Group' },
                    { key: 'Unit', label: 'Unit' }
                  ].map(m => (
                    <button 
                      key={m.key}
                      onClick={() => downloadMasterTemplate(m.key)}
                      className="flex items-center justify-center gap-2 p-2 bg-zinc-50 hover:bg-zinc-100 border border-zinc-200 rounded-lg text-xs font-medium transition-colors"
                    >
                      <Download className="w-3 h-3" />
                      {m.label}
                    </button>
                  ))}
                </div>
              </section>

              {/* Receipt / Payment Voucher Template Guide */}
              {importType === 'Voucher' && (
                <section className="bg-white p-6 rounded-2xl border border-zinc-200 shadow-sm" id="voucher-template-guide">
                  <h2 className="text-lg font-bold mb-3 flex items-center gap-2" id="voucher-template-guide-title">
                    <HelpCircle className="w-5 h-5 text-zinc-900" id="voucher-template-guide-icon" />
                    Voucher Template Guide
                  </h2>
                  <p className="text-xs text-zinc-500 mb-4 leading-relaxed" id="voucher-template-guide-desc">
                    Learn about each field in your Receipt or Payment Excel template and how it maps to Tally's XML tags. Click a field to expand its details.
                  </p>
                  
                  <div className="space-y-2" id="voucher-template-guide-fields">
                    {[
                      {
                        field: 'Date',
                        desc: 'The transaction date. Supports multiple formats which are normalized automatically to Tally format.',
                        input: 'DD/MM/YYYY, YYYY-MM-DD, or Excel serial number (e.g. "08/07/2026", "2026-07-08")',
                        xml: '<DATE>20260708</DATE>'
                      },
                      {
                        field: 'Particulars',
                        desc: 'The ledger name in Tally to be debited or credited. Must match an existing ledger in your Tally Masters for proper synchronization.',
                        input: 'Ledger names (e.g. "Acme Corporation", "HDFC Bank", "Office Expenses")',
                        xml: '<PARTYNAME>Acme Corporation</PARTYNAME>\n<LEDGERNAME>Acme Corporation</LEDGERNAME>'
                      },
                      {
                        field: 'Voucher Number',
                        desc: 'Optional unique transaction identifier. If omitted, Tally auto-numbers the voucher.',
                        input: 'Any text/numeric ID (e.g. "VCH-101", "PMT/2026/001")',
                        xml: '<VOUCHERNUMBER>VCH-101</VOUCHERNUMBER>'
                      },
                      {
                        field: 'Amount',
                        desc: 'The monetary transaction value. Credit ledger entries are marked negative, and debits are positive.',
                        input: 'Numeric amounts (e.g. "1250", "-500.50")',
                        xml: '<AMOUNT>-1250.00</AMOUNT> (Credit)\n<AMOUNT>1250.00</AMOUNT> (Debit)'
                      },
                      {
                        field: 'Narration',
                        desc: 'A description or remark describing the nature or details of the transaction.',
                        input: 'Any explanation text (e.g. "Payment for office stationary")',
                        xml: '<NARRATION>Payment for office stationary</NARRATION>'
                      },
                      {
                        field: 'Reference',
                        desc: 'Optional unique banking or payment ID. It is added as a dedicated XML reference tag and also appended safely to the Narration.',
                        input: 'UTR number, Cheque number, bank transaction ID (e.g. "UTR12345678", "CHQ-002341")',
                        xml: '<REFERENCE>UTR12345678</REFERENCE>\n<NARRATION>... | Ref: UTR12345678</NARRATION>'
                      }
                    ].map(item => {
                      const isOpen = activeGuideField === item.field;
                      return (
                        <div key={item.field} className="border border-zinc-100 rounded-xl overflow-hidden bg-zinc-50" id={`guide-field-${item.field.toLowerCase().replace(/\s+/g, '-')}`}>
                          <button
                            type="button"
                            onClick={() => setActiveGuideField(isOpen ? null : item.field)}
                            className="w-full flex items-center justify-between p-3.5 text-left font-semibold text-sm text-zinc-900 hover:bg-zinc-100/80 transition-colors"
                            id={`guide-field-${item.field.toLowerCase().replace(/\s+/g, '-')}-btn`}
                          >
                            <span>{item.field}</span>
                            <span className="text-zinc-400 text-xs font-medium bg-white px-2 py-1 rounded border border-zinc-200/60 shadow-sm" id={`guide-field-${item.field.toLowerCase().replace(/\s+/g, '-')}-status`}>
                              {isOpen ? 'Collapse' : 'Expand'}
                            </span>
                          </button>
                          
                          {isOpen && (
                            <div className="p-3.5 border-t border-zinc-200/60 bg-white text-xs text-zinc-600 space-y-2.5 leading-relaxed" id={`guide-field-${item.field.toLowerCase().replace(/\s+/g, '-')}-content`}>
                              <div>
                                <span className="font-semibold text-zinc-800">Description:</span> {item.desc}
                              </div>
                              <div>
                                <span className="font-semibold text-zinc-800">What to enter:</span> <span className="text-zinc-700 italic">{item.input}</span>
                              </div>
                              <div className="space-y-1">
                                <span className="font-semibold text-zinc-800">XML Tag Mapping:</span>
                                <pre className="p-2 bg-zinc-950 text-emerald-400 font-mono text-[11px] rounded overflow-x-auto leading-relaxed border border-zinc-800 shadow-inner">
                                  {item.xml}
                                </pre>
                              </div>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </section>
              )}

              <section className="bg-white p-6 rounded-2xl border border-zinc-200 shadow-sm h-full max-h-[600px] overflow-y-auto">
              <h2 className="text-xl font-bold mb-4 flex items-center gap-2">
                <History className="w-5 h-5" />
                Recent Activity
              </h2>
              
              <div className="space-y-4">
                {conversions.length === 0 ? (
                  <div className="text-center py-12 text-zinc-400">
                    <p>No recent conversions</p>
                  </div>
                ) : (
                  conversions.map((conv) => (
                    <div key={conv.id} className="flex items-center justify-between p-3 hover:bg-zinc-50 rounded-xl transition-colors border border-transparent hover:border-zinc-100">
                      <div className="flex items-center gap-3">
                        <div className={`p-2 rounded-lg ${
                          conv.status === 'completed' ? 'bg-green-50 text-green-600' :
                          conv.status === 'failed' ? 'bg-red-50 text-red-600' :
                          'bg-zinc-100 text-zinc-600'
                        }`}>
                          {conv.status === 'completed' ? <CheckCircle2 className="w-4 h-4" /> :
                           conv.status === 'failed' ? <AlertCircle className="w-4 h-4" /> :
                           <Loader2 className="w-4 h-4 animate-spin" />}
                        </div>
                        <div className="min-w-0">
                          <p className="text-sm font-medium truncate max-w-[120px]">{conv.fileName}</p>
                          <p className="text-xs text-zinc-500">
                            {(() => {
                              if (!conv.timestamp) return '';
                              if (typeof conv.timestamp.toDate === 'function') {
                                return conv.timestamp.toDate().toLocaleDateString();
                              }
                              if (conv.timestamp.seconds) {
                                return new Date(conv.timestamp.seconds * 1000).toLocaleDateString();
                              }
                              return new Date(conv.timestamp).toLocaleDateString();
                            })()}
                          </p>
                        </div>
                      </div>
                      {conv.status === 'completed' && conv.xmlContent && (
                        <button 
                          onClick={() => downloadXML(conv.xmlContent!, conv.fileName)}
                          className="p-2 hover:bg-zinc-200 rounded-full transition-colors"
                          title="Download XML"
                        >
                          <Download className="w-4 h-4" />
                        </button>
                      )}
                    </div>
                  ))
                )}
              </div>
            </section>
          </div>
        </div>
        )}
      </main>

      {/* Restart / Reset Confirmation Modal */}
      <AnimatePresence>
        {showRestartConfirm && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4" id="restart-modal">
            {/* Backdrop */}
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowRestartConfirm(false)}
              className="fixed inset-0 bg-zinc-900/40 backdrop-blur-sm"
              id="restart-modal-backdrop"
            />
            
            {/* Modal Content */}
            <motion.div
              initial={{ opacity: 0, scale: 0.95, y: 10 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 10 }}
              transition={{ duration: 0.2, ease: "easeOut" }}
              className="bg-white w-full max-w-md rounded-2xl p-6 shadow-xl border border-zinc-200 relative z-10"
              id="restart-modal-container"
            >
              <div className="flex items-start gap-4 mb-4">
                <div className="w-10 h-10 bg-amber-50 rounded-xl flex items-center justify-center text-amber-600 flex-shrink-0" id="restart-modal-icon-bg">
                  <AlertCircle className="w-6 h-6" id="restart-modal-icon" />
                </div>
                <div>
                  <h3 className="text-lg font-bold text-zinc-900" id="restart-modal-title">Restart Workspace?</h3>
                  <p className="text-zinc-500 text-sm mt-1 leading-relaxed" id="restart-modal-message">
                    This will clear the current loaded Tally context, uploaded files, mappings, review data, generated XML, and active task progress. Do you want to continue?
                  </p>
                </div>
              </div>

              <div className="flex items-center justify-end gap-3" id="restart-modal-actions">
                <button
                  type="button"
                  onClick={() => setShowRestartConfirm(false)}
                  className="px-4 py-2 bg-zinc-50 hover:bg-zinc-100 border border-zinc-200 rounded-xl text-zinc-700 font-medium text-sm transition-colors"
                  id="restart-modal-cancel-btn"
                >
                  Cancel
                </button>
                <button
                  type="button"
                  onClick={handleRestartWorkspace}
                  disabled={isProcessing}
                  className="flex items-center gap-2 px-4 py-2 bg-zinc-900 hover:bg-zinc-800 disabled:bg-zinc-400 text-white rounded-xl font-medium text-sm transition-colors"
                  id="restart-modal-confirm-btn"
                >
                  {isProcessing ? (
                    <>
                      <Loader2 className="w-4 h-4 animate-spin" />
                      Clearing...
                    </>
                  ) : (
                    'Yes, Restart'
                  )}
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Direct Tally Push Confirmation Modal */}
      <AnimatePresence>
        {isPushModalOpen && pushModalData && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4" id="tally-push-modal">
            {/* Backdrop */}
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => {
                if (!isPushing) setIsPushModalOpen(false);
              }}
              className="fixed inset-0 bg-zinc-900/40 backdrop-blur-sm"
              id="tally-push-backdrop"
            />
            
            {/* Modal Content */}
            <motion.div
              initial={{ opacity: 0, scale: 0.95, y: 10 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 10 }}
              transition={{ duration: 0.2, ease: "easeOut" }}
              className="bg-white w-full max-w-md rounded-2xl p-6 shadow-xl border border-zinc-200 relative z-10"
              id="tally-push-container"
            >
              <div className="flex items-start gap-4 mb-4">
                <div className="w-10 h-10 bg-emerald-50 rounded-xl flex items-center justify-center text-emerald-600 flex-shrink-0" id="tally-push-icon-bg">
                  <RefreshCw className="w-5 h-5 animate-spin-slow" id="tally-push-icon" />
                </div>
                <div>
                  <h3 className="text-lg font-bold text-zinc-900" id="tally-push-title">Push XML directly to TallyPrime</h3>
                  <p className="text-zinc-500 text-xs mt-1 leading-relaxed">
                    Review import details before pushing generated entries directly into your local running Tally instance.
                  </p>
                </div>
              </div>

              {/* Tally Details Grid */}
              <div className="bg-zinc-50 border border-zinc-200 rounded-xl p-4 space-y-2.5 mb-5 text-xs text-zinc-700" id="tally-push-details">
                <div className="flex justify-between">
                  <span className="font-semibold text-zinc-500">Target Company:</span>
                  <span className="font-bold text-emerald-700 break-all text-right">{pushModalData.companyName}</span>
                </div>
                <div className="flex justify-between">
                  <span className="font-semibold text-zinc-500">XML Type:</span>
                  <span className="font-bold text-zinc-800">{pushModalData.xmlType}</span>
                </div>
                <div className="flex justify-between">
                  <span className="font-semibold text-zinc-500">Voucher Count:</span>
                  <span className="font-bold text-zinc-800">{pushModalData.voucherCount}</span>
                </div>
                <div className="flex justify-between">
                  <span className="font-semibold text-zinc-500">Missing Masters Created:</span>
                  <span className="font-bold text-zinc-800">{pushModalData.masterCount}</span>
                </div>
              </div>

              {/* Warning Alert */}
              <div className="bg-amber-50 border border-amber-200 rounded-xl p-3.5 text-xs text-amber-800 flex items-start gap-2 mb-5" id="tally-push-warning">
                <AlertCircle className="w-4 h-4 shrink-0 mt-0.5 text-amber-600" />
                <div>
                  <span className="font-bold block">Safety Recommendation</span>
                  <span className="text-amber-700 mt-0.5 block leading-relaxed">
                    Please ensure you have a backup of your Tally company data before continuing with the direct XML import.
                  </span>
                </div>
              </div>

              {/* Buttons */}
              <div className="flex flex-col sm:flex-row items-center justify-end gap-2.5" id="tally-push-actions">
                <button
                  type="button"
                  onClick={() => setIsPushModalOpen(false)}
                  disabled={isPushing}
                  className="w-full sm:w-auto px-4 py-2.5 bg-zinc-50 hover:bg-zinc-100 border border-zinc-200 rounded-xl text-zinc-700 font-bold text-xs transition-colors"
                  id="tally-push-cancel-btn"
                >
                  Cancel
                </button>
                <button
                  type="button"
                  onClick={() => {
                    downloadXML(pushModalData.xmlContent, 'TallyGen_Backup_' + Date.now() + '.xml');
                  }}
                  disabled={isPushing}
                  className="w-full sm:w-auto px-4 py-2.5 bg-white hover:bg-zinc-50 border border-zinc-200 rounded-xl text-zinc-700 font-bold text-xs transition-colors flex items-center justify-center gap-1.5"
                  id="tally-push-backup-btn"
                >
                  <Download className="w-3.5 h-3.5" />
                  Download Backup XML
                </button>
                <button
                  type="button"
                  onClick={handleDirectPushToTally}
                  disabled={isPushing}
                  className="w-full sm:w-auto flex items-center justify-center gap-1.5 px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 disabled:bg-emerald-400 text-white rounded-xl font-bold text-xs transition-colors shadow-md shadow-emerald-500/10"
                  id="tally-push-confirm-btn"
                >
                  {isPushing ? (
                    <>
                      <Loader2 className="w-3.5 h-3.5 animate-spin" />
                      Pushing...
                    </>
                  ) : (
                    <>
                      <Check className="w-3.5 h-3.5" />
                      Push to Tally
                    </>
                  )}
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
