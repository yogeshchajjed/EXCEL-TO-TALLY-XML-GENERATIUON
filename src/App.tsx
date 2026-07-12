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
  clearOfflineWorkspace 
} from './lib/storageAdapter';
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
  FileCode,
  Database,
  FileUp,
  FileText,
  Search,
  RotateCcw,
  Upload,
  HelpCircle,
  Building
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
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
  TallyVoucher, 
  TallyLedgerMaster, 
  TallyStockItemMaster, 
  TallyStockGroupMaster, 
  TallyUnitMaster 
} from './lib/tallyXml';

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

export default function App() {
  const [user, setUser] = useState<User | null>(null);
  const [loading, setLoading] = useState(true);
  const [conversions, setConversions] = useState<ConversionRecord[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [mappings, setMappings] = useState<ColumnMapping[]>([]);
  const [currentStep, setCurrentStep] = useState<'upload' | 'mapping' | 'complete' | 'context' | 'master-review' | 'bank-statement-review'>('upload');
  const [pendingData, setPendingData] = useState<any[]>([]);
  const [pendingFileName, setPendingFileName] = useState('');
  const [tallyContext, setTallyContext] = useState<TallyContext | null>(null);
  const [isContextLoading, setIsContextLoading] = useState(false);
  const [selectedVoucherType, setSelectedVoucherType] = useState('Payment');
  const [selectedBankLedger, setSelectedBankLedger] = useState('');
  const [aiMappedTransactions, setAiMappedTransactions] = useState<MappedTransaction[]>([]);

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

  // Bank Statement Import States
  const [voucherImportMethod, setVoucherImportMethod] = useState<'template' | 'bankStatement'>('template');
  const [voucherMode, setVoucherMode] = useState<'auto' | 'payment' | 'receipt'>('auto');
  const [bankStatementRows, setBankStatementRows] = useState<any[]>([]);

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

  const downloadTemplate = (type: string) => {
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

    const ws = XLSX.utils.json_to_sheet(sampleData, { header: headers });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    XLSX.writeFile(wb, `Tally_${type}_Template.xlsx`);
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
        list.forEach((l: any) => {
          const name = l['@_NAME'] || l.NAME;
          if (name) {
            const nameStr = String(name);
            ledgers.push(nameStr);
            const parent = l.PARENT || l['@_PARENT'];
            const parentStr = getParentName(parent);
            if (parentStr) ledgerGroupMap[nameStr] = parentStr;
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
        list.forEach((si: any) => {
          const name = si['@_NAME'] || si.NAME;
          if (name) {
            const nameStr = String(name);
            stockItems.push(nameStr);
            if (si.PARENT) stockItemStockGroupMap[nameStr] = String(si.PARENT);
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
          historicalMappings: existingMappings
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
          historicalMappings: mergedHistorical
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
      historicalMappings: historicalMappings.slice(0, 500)
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
  const suggestLedgerForNarration = (narration: string, context: TallyContext | null): string => {
    if (!context || !narration) return '';
    const cleanNarr = narration.toLowerCase().trim();
    
    // 1. Exact old narration mapping from historicalMappings
    if (context.historicalMappings) {
      const exactHist = context.historicalMappings.find(h => h.narration.toLowerCase().trim() === cleanNarr);
      if (exactHist && context.ledgers.includes(exactHist.ledger)) {
        return exactHist.ledger;
      }
    }

    // 2. Keyword matching against ledger names
    for (const ledger of context.ledgers) {
      const ledgerLower = ledger.toLowerCase();
      if (cleanNarr.includes(ledgerLower) && ledgerLower.length > 3) {
        return ledger;
      }
    }

    // 3. Similar narration mapping from historicalMappings
    if (context.historicalMappings) {
      let bestMatch: { ledger: string; score: number } | null = null;
      for (const h of context.historicalMappings) {
        const score = getSimilarity(h.narration, narration);
        if (score >= 0.8) {
          if (!bestMatch || score > bestMatch.score) {
            bestMatch = { ledger: h.ledger, score };
          }
        }
      }
      if (bestMatch && context.ledgers.includes(bestMatch.ledger)) {
        return bestMatch.ledger;
      }
    }

    // 4. Fallback: Keyword matches of individual ledger name words
    for (const ledger of context.ledgers) {
      const ledgerWords = ledger.toLowerCase().split(/[^a-zA-Z0-9]/).filter(w => w.length > 3);
      if (ledgerWords.length > 0 && ledgerWords.every(w => cleanNarr.includes(w))) {
        return ledger;
      }
    }

    return '';
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
    const NARRATION_HEADERS = ['narration', 'description', 'particulars', 'transaction details', 'details', 'remarks'];
    const DEBIT_HEADERS = ['debit', 'withdrawal', 'withdrawals', 'paid out', 'dr', 'debit amount'];
    const CREDIT_HEADERS = ['credit', 'deposit', 'deposits', 'paid in', 'cr', 'credit amount'];
    const AMOUNT_HEADERS = ['amount', 'transaction amount'];
    const REF_HEADERS = ['reference', 'ref no', 'utr', 'cheque no', 'instrument no', 'transaction id'];

    const dateCol = getHeaderKey(DATE_HEADERS);
    const descCol = getHeaderKey(NARRATION_HEADERS);
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

    try {
      const validRows = bankStatementRows.filter(row => !row.excluded && row.status !== 'invalid');
      if (validRows.length === 0) {
        throw new Error("No valid, non-excluded bank statement rows available to generate XML.");
      }

      const missingLedgerRow = validRows.find(row => !row.userLedger);
      if (missingLedgerRow) {
        throw new Error(`Row ${missingLedgerRow.rowNo} has no Particulars Ledger selected. Please select a ledger.`);
      }

      const conversionPayload: any = {
        fileName: pendingFileName,
        voucherType: 'Voucher',
        status: 'processing',
        bankLedger: selectedBankLedger
      };
      if (getAppMode() === 'web') {
        conversionPayload.timestamp = serverTimestamp();
      }
      const conversionId = await saveConversion(user?.uid || OFFLINE_USER.uid, conversionPayload);

      const vouchers: TallyVoucher[] = validRows.map(row => {
        const dateNorm = normalizeTallyDate(row.date);
        if (!dateNorm.isValid) {
          throw new Error(`Row ${row.rowNo}: Date "${row.date}" could not be parsed: ${dateNorm.error}`);
        }

        const vType = row.detectedVoucherType === 'Unknown' ? 'Payment' : row.detectedVoucherType;

        const voucher: TallyVoucher = {
          date: dateNorm.value,
          voucherType: vType,
          partyName: row.userLedger,
          voucherNumber: undefined,
          narration: row.description || undefined,
          reference: row.reference || undefined,
          ledgerEntries: []
        };

        const isPartyPositive = vType === 'Payment' ? 'Yes' : 'No';
        voucher.ledgerEntries.push({
          ledgerName: row.userLedger,
          isDeemedPositive: isPartyPositive,
          amount: Math.abs(row.amount)
        });

        const isBankPositive = vType === 'Payment' ? 'No' : 'Yes';
        voucher.ledgerEntries.push({
          ledgerName: selectedBankLedger,
          isDeemedPositive: isBankPositive,
          amount: Math.abs(row.amount)
        });

        return voucher;
      });

      const xml = generateTallyXML(vouchers);
      await updateConversion(user?.uid || OFFLINE_USER.uid, conversionId, {
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
    } catch (err) {
      setError(err instanceof Error ? err.message : "Conversion failed");
      setIsProcessing(false);
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
          const results = rawRows.map((row, idx) => {
            const suggestedLedger = suggestLedgerForNarration(row.description, tallyContext);
            let detectedVoucherType: 'Payment' | 'Receipt' | 'Unknown' = 'Unknown';
            let status: 'valid' | 'invalid' | 'warning' = 'valid';
            let errorMsg = '';
            let finalAmount = row.amount || 0;

            const hasDebit = row.debit !== null && row.debit > 0;
            const hasCredit = row.credit !== null && row.credit > 0;

            if (hasDebit && hasCredit) {
              status = 'invalid';
              errorMsg = 'Both debit and credit found in same row. Please verify.';
            } else if (!hasDebit && !hasCredit && !finalAmount) {
              status = 'invalid';
              errorMsg = 'Debit, Credit, and Amount are all blank.';
            } else if (hasDebit) {
              detectedVoucherType = 'Payment';
              finalAmount = row.debit!;
            } else if (hasCredit) {
              detectedVoucherType = 'Receipt';
              finalAmount = row.credit!;
            } else {
              if (selectedVoucherType === 'Payment') {
                detectedVoucherType = 'Payment';
                status = 'warning';
                errorMsg = 'Single Amount column found. Voucher type has been applied using selected screen mode.';
              } else if (selectedVoucherType === 'Receipt') {
                detectedVoucherType = 'Receipt';
                status = 'warning';
                errorMsg = 'Single Amount column found. Voucher type has been applied using selected screen mode.';
              }
            }

            let excluded = false;
            if (status === 'valid' || status === 'warning') {
              if (voucherMode === 'payment') {
                if (detectedVoucherType === 'Receipt') {
                  status = 'warning';
                  errorMsg = 'Skipped: Credit row in Payment only mode';
                  excluded = true;
                }
              } else if (voucherMode === 'receipt') {
                if (detectedVoucherType === 'Payment') {
                  status = 'warning';
                  errorMsg = 'Skipped: Debit row in Receipt only mode';
                  excluded = true;
                }
              }
            }

            return {
              rowNo: idx + 1,
              rawRow: row,
              date: row.date,
              rawDate: row.date,
              description: row.description,
              debit: row.debit,
              credit: row.credit,
              amount: finalAmount,
              detectedVoucherType,
              suggestedLedger,
              userLedger: suggestedLedger,
              reference: row.reference || '',
              status,
              errorMsg,
              excluded
            };
          });

          setBankStatementRows(results);
          setCurrentStep('bank-statement-review');
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

              const results = parseBankStatementExcelOrCsv(jsonData, voucherMode);
              setBankStatementRows(results);
              setCurrentStep('bank-statement-review');
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
            let descCol = aiMappings.find(m => m.tallyField === 'NARRATION')?.excelColumn || 
                          aiMappings.find(m => m.tallyField === 'PARTYNAME')?.excelColumn;
            const amtCol = aiMappings.find(m => m.tallyField === 'AMOUNT')?.excelColumn;

            // Defensive: If AI mapped Date to Narration, try to find another candidate for Narration
            if (descCol === dateCol) {
              descCol = headers.find(h => 
                !['date', 'amount', 'balance', 'vch', 'no'].some(k => h.toLowerCase().includes(k)) &&
                ['particulars', 'description', 'narration', 'remarks', 'details'].some(k => h.toLowerCase().includes(k))
              ) || descCol;
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
                  const val = getValueByHeader(row, h);
                  if (val && String(val).trim() !== '') {
                    narration = val;
                    break;
                  }
                }
              }

              return {
                date: String(dateVal),
                description: String(narration || 'No Narration Found'),
                amount: isNaN(parsedAmount) ? 0 : parsedAmount
              };
            });

            if (tallyContext) {
              const mapped = await mapBankTransactions(bankTransactions, tallyContext.ledgers, tallyContext.historicalMappings);
              setAiMappedTransactions(mapped);
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
      const conversionPayload: any = {
        fileName: pendingFileName,
        status: 'processing',
        voucherType: selectedVoucherType,
        bankLedger: selectedBankLedger
      };
      if (getAppMode() === 'web') {
        conversionPayload.timestamp = serverTimestamp();
      }
      const conversionId = await saveConversion(user.uid, conversionPayload);

      const vouchers: TallyVoucher[] = pendingData.map((row: any, rowIndex: number) => {
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

        const voucher: TallyVoucher = {
          date: dateNorm.value,
          voucherType: selectedVoucherType,
          partyName: partyNameVal || 'Unknown',
          voucherNumber: voucherNumberVal || undefined,
          narration: narrationVal || undefined,
          reference: referenceVal || undefined,
          ledgerEntries: []
        };

        voucher.ledgerEntries.push({
          ledgerName: voucher.partyName,
          isDeemedPositive: amountVal > 0 ? 'Yes' : 'No',
          amount: Math.abs(amountVal)
        });

        if (voucher.ledgerEntries.length === 1) {
          const mainEntry = voucher.ledgerEntries[0];
          
          // Logic for Receipt: Bank is Debited (Positive), Mapped Account is Credited (Negative)
          // Logic for Payment: Bank is Credited (Negative), Mapped Account is Debited (Positive)
          if (selectedVoucherType === 'Receipt') {
            mainEntry.isDeemedPositive = 'No'; // Credit the mapped account
            voucher.ledgerEntries.push({
              ledgerName: selectedBankLedger || 'Bank Account',
              isDeemedPositive: 'Yes', // Debit the bank
              amount: mainEntry.amount
            });
          } else {
            // Default (Payment) logic
            mainEntry.isDeemedPositive = 'Yes'; // Debit the mapped account
            voucher.ledgerEntries.push({
              ledgerName: selectedBankLedger || 'Bank Account',
              isDeemedPositive: 'No', // Credit the bank
              amount: mainEntry.amount
            });
          }
        }

        return voucher;
      });

      const xml = generateTallyXML(vouchers);
      await updateConversion(user.uid, conversionId, {
        status: 'completed',
        xmlContent: xml
      });

      setCurrentStep('complete');
      setIsProcessing(false);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Conversion failed");
      setIsProcessing(false);
    }
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
                              </select>
                            </div>

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
                                {['Payment', 'Receipt', 'Contra', 'Journal'].map(type => (
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
                        importType === 'Voucher' && !selectedBankLedger 
                          ? 'border-zinc-100 bg-zinc-50/50 cursor-not-allowed' 
                          : 'border-zinc-200 hover:border-zinc-400 cursor-pointer'
                      }`}>
                        <input 
                          type="file" 
                          accept=".xlsx, .xls, .csv, .pdf" 
                          onChange={handleFileUpload}
                          className="absolute inset-0 opacity-0 cursor-pointer disabled:cursor-not-allowed"
                          disabled={isProcessing || (importType === 'Voucher' && !selectedBankLedger)}
                        />
                        <div className="space-y-4">
                          <div className={`w-12 h-12 rounded-full flex items-center justify-center mx-auto transition-colors ${
                            importType === 'Voucher' && !selectedBankLedger 
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

              {currentStep === 'bank-statement-review' && (
                <motion.section 
                  key="bank-statement-review"
                  initial={{ opacity: 0, x: -20 }}
                  animate={{ opacity: 1, x: 0 }}
                  exit={{ opacity: 0, x: 20 }}
                  className="bg-white p-6 rounded-2xl border border-zinc-200 shadow-sm space-y-6"
                >
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-zinc-100 pb-5">
                    <div>
                      <h2 className="text-xl font-bold flex items-center gap-2 text-zinc-900">
                        <Building className="w-5 h-5 text-zinc-800" />
                        Step 2: Bank Statement Review
                      </h2>
                      <p className="text-xs text-zinc-500 mt-1 leading-relaxed">
                        Verify transactions, select particulars ledgers, and exclude rows before generating Tally XML.
                      </p>
                    </div>
                    <div className="text-right">
                      <span className="text-xs bg-zinc-100 px-3 py-1.5 rounded-lg text-zinc-700 font-semibold block md:inline-block">
                        {pendingFileName}
                      </span>
                    </div>
                  </div>

                  {/* Summary Dashboard widgets */}
                  <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                    <div className="bg-zinc-50 p-4 rounded-xl border border-zinc-200/50">
                      <p className="text-[10px] uppercase font-bold tracking-wider text-zinc-400">Total Rows</p>
                      <p className="text-xl font-bold text-zinc-800 mt-1">{bankStatementRows.length}</p>
                    </div>
                    <div className="bg-emerald-50/50 p-4 rounded-xl border border-emerald-100">
                      <p className="text-[10px] uppercase font-bold tracking-wider text-emerald-600">Valid to Import</p>
                      <p className="text-xl font-bold text-emerald-700 mt-1">
                        {bankStatementRows.filter(r => r.status !== 'invalid' && !r.excluded).length}
                      </p>
                    </div>
                    <div className="bg-amber-50/50 p-4 rounded-xl border border-amber-100">
                      <p className="text-[10px] uppercase font-bold tracking-wider text-amber-600">Warnings/Skipped</p>
                      <p className="text-xl font-bold text-amber-700 mt-1">
                        {bankStatementRows.filter(r => r.status === 'warning' || r.excluded).length}
                      </p>
                    </div>
                    <div className="bg-rose-50/50 p-4 rounded-xl border border-rose-100">
                      <p className="text-[10px] uppercase font-bold tracking-wider text-rose-600">Invalid Rows</p>
                      <p className="text-xl font-bold text-rose-700 mt-1">
                        {bankStatementRows.filter(r => r.status === 'invalid').length}
                      </p>
                    </div>
                  </div>

                  {/* Desktop view scroll container */}
                  <div className="border border-zinc-200 rounded-xl overflow-hidden bg-white shadow-sm max-h-[500px] overflow-y-auto">
                    <table className="w-full text-xs text-left border-collapse">
                      <thead className="bg-zinc-50 text-zinc-600 uppercase font-bold border-b border-zinc-200 sticky top-0 z-10">
                        <tr>
                          <th className="p-3 w-12 text-center">Row</th>
                          <th className="p-3 w-24">Date</th>
                          <th className="p-3 max-w-xs">Narration / Description</th>
                          <th className="p-3 text-right">Debit (Dr)</th>
                          <th className="p-3 text-right">Credit (Cr)</th>
                          <th className="p-3 text-right">Amount</th>
                          <th className="p-3 w-24 text-center">Vch Type</th>
                          <th className="p-3 w-56">Particulars Ledger (Tally)</th>
                          <th className="p-3 w-24">Reference</th>
                          <th className="p-3 w-32">Status / Message</th>
                          <th className="p-3 w-16 text-center">Exclude</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-zinc-100">
                        {bankStatementRows.map((row, idx) => {
                          const isInvalid = row.status === 'invalid';
                          const isWarning = row.status === 'warning';
                          const isExcluded = row.excluded;

                          return (
                            <tr 
                              key={idx} 
                              className={`transition-colors ${
                                isExcluded 
                                  ? 'bg-zinc-50/60 text-zinc-400' 
                                  : isInvalid 
                                    ? 'bg-rose-50/30' 
                                    : isWarning 
                                      ? 'bg-amber-50/20' 
                                      : 'hover:bg-zinc-50/50'
                              }`}
                            >
                              <td className="p-3 text-center font-mono font-medium">{row.rowNo}</td>
                              <td className="p-3 font-medium whitespace-nowrap">{row.date}</td>
                              <td className="p-3 max-w-xs truncate font-medium" title={row.description}>
                                {row.description}
                              </td>
                              <td className="p-3 text-right font-mono font-medium text-rose-600">
                                {row.debit !== null && row.debit !== undefined ? row.debit.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'}
                              </td>
                              <td className="p-3 text-right font-mono font-medium text-emerald-600">
                                {row.credit !== null && row.credit !== undefined ? row.credit.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'}
                              </td>
                              <td className="p-3 text-right font-mono font-bold text-zinc-900">
                                {row.amount.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                              </td>
                              <td className="p-3 text-center">
                                <span className={`inline-block px-2 py-1 rounded text-[10px] font-bold uppercase tracking-wider ${
                                  row.detectedVoucherType === 'Payment'
                                    ? 'bg-rose-100 text-rose-800'
                                    : row.detectedVoucherType === 'Receipt'
                                      ? 'bg-emerald-100 text-emerald-800'
                                      : 'bg-zinc-100 text-zinc-600'
                                }`}>
                                  {row.detectedVoucherType}
                                </span>
                              </td>
                              <td className="p-2">
                                {tallyContext ? (
                                  <select
                                    value={row.userLedger}
                                    onChange={(e) => {
                                      const updated = [...bankStatementRows];
                                      updated[idx].userLedger = e.target.value;
                                      setBankStatementRows(updated);
                                    }}
                                    disabled={isExcluded}
                                    className="w-full p-2 bg-zinc-50 border border-zinc-200 rounded-lg text-xs outline-none focus:ring-1 focus:ring-zinc-950 disabled:opacity-50"
                                  >
                                    <option value="">-- Select Ledger --</option>
                                    {tallyContext.ledgers.map(l => (
                                      <option key={l} value={l}>{l}</option>
                                    ))}
                                  </select>
                                ) : (
                                  <input
                                    type="text"
                                    value={row.userLedger}
                                    onChange={(e) => {
                                      const updated = [...bankStatementRows];
                                      updated[idx].userLedger = e.target.value;
                                      setBankStatementRows(updated);
                                    }}
                                    disabled={isExcluded}
                                    placeholder="Enter ledger name"
                                    className="w-full p-2 bg-zinc-50 border border-zinc-200 rounded-lg text-xs outline-none focus:ring-1 focus:ring-zinc-950 disabled:opacity-50 font-medium"
                                  />
                                )}
                              </td>
                              <td className="p-3 font-medium whitespace-nowrap">{row.reference}</td>
                              <td className="p-3">
                                {isInvalid && (
                                  <span className="text-rose-600 font-semibold flex items-center gap-1" title={row.errorMsg}>
                                    <AlertCircle className="w-3.5 h-3.5 shrink-0" />
                                    Error
                                  </span>
                                )}
                                {isWarning && (
                                  <span className="text-amber-600 font-semibold flex items-center gap-1" title={row.errorMsg}>
                                    <AlertCircle className="w-3.5 h-3.5 shrink-0" />
                                    Warning
                                  </span>
                                )}
                                {!isInvalid && !isWarning && !isExcluded && (
                                  <span className="text-emerald-600 font-semibold flex items-center gap-1">
                                    <CheckCircle2 className="w-3.5 h-3.5 shrink-0" />
                                    Ready
                                  </span>
                                )}
                                {isExcluded && (
                                  <span className="text-zinc-400 font-semibold">
                                    Excluded
                                  </span>
                                )}
                              </td>
                              <td className="p-3 text-center">
                                <input
                                  type="checkbox"
                                  checked={row.excluded}
                                  onChange={(e) => {
                                    const updated = [...bankStatementRows];
                                    updated[idx].excluded = e.target.checked;
                                    setBankStatementRows(updated);
                                  }}
                                  className="rounded border-zinc-300 text-zinc-900 focus:ring-zinc-950 w-4 h-4 cursor-pointer"
                                />
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>

                  {/* Bottom validation and actions block */}
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 pt-4 border-t border-zinc-100">
                    <div className="text-xs text-zinc-500 leading-relaxed max-w-xl">
                      <span className="font-bold text-zinc-700">Rules applied:</span> Valid rows and non-excluded warnings will be written to the generated Tally XML. Invalid rows or manually excluded rows will be skipped. Ledger dropdown utilizes uploaded Tally masters context.
                    </div>
                    <div className="flex items-center gap-3 justify-end shrink-0">
                      <button
                        type="button"
                        onClick={() => {
                          setBankStatementRows([]);
                          setCurrentStep('upload');
                        }}
                        className="px-5 py-3 hover:bg-zinc-100 border border-zinc-200 text-zinc-700 rounded-xl text-xs font-bold transition-all flex items-center gap-2"
                      >
                        <RotateCcw className="w-4 h-4" />
                        Cancel / Restart
                      </button>
                      <button
                        type="button"
                        onClick={handleBankStatementGenerateXML}
                        disabled={isProcessing}
                        className="px-6 py-3 bg-zinc-900 hover:bg-zinc-850 text-white rounded-xl text-xs font-bold transition-all flex items-center gap-2 disabled:opacity-50 shadow-sm"
                      >
                        {isProcessing ? (
                          <>
                            <Loader2 className="w-4 h-4 animate-spin" />
                            Processing...
                          </>
                        ) : (
                          <>
                            <FileCode className="w-4 h-4" />
                            Generate Tally XML
                          </>
                        )}
                      </button>
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
                              <ArrowRight className="w-3 h-3 text-zinc-400" />
                              <span className="font-bold text-zinc-900">{tx.tallyLedger}</span>
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
                      {importType === 'Voucher' 
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
                      
                      {importType === 'Voucher' ? (
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
    </div>
  );
}
