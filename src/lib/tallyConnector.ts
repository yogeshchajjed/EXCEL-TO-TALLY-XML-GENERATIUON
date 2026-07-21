import { getAppMode } from './storageAdapter';
import { ElectronTallyConfig, ElectronTallyResponse } from '../types/electron';
import * as XLSX from 'xlsx';
import { normalizeTallyDate } from './tallyXml';

export interface LedgerMasterRow {
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
  excluded: boolean;
}

export interface StockMasterRow {
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
}

export interface TallyDaybookTransaction {
  date: string;
  voucherType: string;
  narration: string;
  ledger: string;
  amount: number;
  reference: string;
}

export interface TallyResponseSummary {
  success: boolean;
  companyName?: string;
  ledgers?: any[];
  groups?: any[];
  stockItems?: any[];
  units?: any[];
  vouchers?: any[];
  createdCount?: number;
  alteredCount?: number;
  errorCount?: number;
  errors?: string[];
}

/**
 * Checks if direct local Tally connection is available.
 * Enabled only in Electron desktop offline mode.
 */
export function isDirectTallyAvailable(): boolean {
  return (
    getAppMode() === 'desktop-offline' &&
    typeof window !== 'undefined' &&
    !!(window as any).electron?.tally
  );
}

/**
 * XML builder to fetch all Ledgers
 */
export function buildExportLedgersRequest(): string {
  return `
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>LedgerCollection</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="LedgerCollection">
            <TYPE>Ledger</TYPE>
            <FETCH>Name, Parent, OpeningBalance, MailingName, Address, LEDState, PINCode, IncomeTaxNumber, GSTRegistrationType, PartyGSTIN, Email, Phone</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>`.trim();
}

/**
 * XML builder to fetch all Groups
 */
export function buildExportGroupsRequest(): string {
  return `
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>GroupCollection</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="GroupCollection">
            <TYPE>Group</TYPE>
            <FETCH>Name, Parent</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>`.trim();
}

/**
 * XML builder to fetch all Stock Items
 */
export function buildExportStockItemsRequest(): string {
  return `
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>StockItemCollection</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="StockItemCollection">
            <TYPE>StockItem</TYPE>
            <FETCH>Name, Parent, BaseUnits, OpeningBalance, OpeningRate, OpeningValue, HSNCode, GSTApplicable, TAXABILITY, Description, IGSTRate, CGSTRate, SGSTRate</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>`.trim();
}

/**
 * XML builder to fetch all Units
 */
export function buildExportUnitsRequest(): string {
  return `
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>UnitCollection</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="UnitCollection">
            <TYPE>Unit</TYPE>
            <FETCH>Name, OriginalName, DecimalPlaces</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>`.trim();
}

/**
 * XML builder to fetch Daybook Vouchers with date range
 */
export function buildExportDaybookRequest(fromDate: string, toDate: string): string {
  const fDateObj = normalizeTallyDate(fromDate);
  const fDate = fDateObj.isValid ? fDateObj.value : fromDate.replace(/[-/]/g, '');
  const tDateObj = normalizeTallyDate(toDate);
  const tDate = tDateObj.isValid ? tDateObj.value : toDate.replace(/[-/]/g, '');
  return `
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>DaybookCollection</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVFROMDATE>${fDate}</SVFROMDATE>
        <SVTODATE>${tDate}</SVTODATE>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="DaybookCollection">
            <TYPE>Voucher</TYPE>
            <FETCH>Date, VoucherTypeName, VoucherNumber, Narration, Reference, LedgerEntries, AllLedgerEntries</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>`.trim();
}

/**
 * XML builder to import payload
 */
export function buildImportXmlRequest(xml: string): string {
  if (xml.trim().startsWith('<ENVELOPE>')) {
    return xml.trim();
  }
  return `
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Import Data</TALLYREQUEST>
  </HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>All Masters</REPORTNAME>
      </REQUESTDESC>
      <REQUESTDATA>
        ${xml}
      </REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>`.trim();
}

/**
 * Tests direct connection to local TallyPrime
 */
export async function testTallyConnection(config: ElectronTallyConfig): Promise<ElectronTallyResponse> {
  if (!isDirectTallyAvailable()) {
    return { success: false, error: 'Direct Tally connection is only available in Desktop Offline App mode.' };
  }
  return await window.electron.tally.testConnection(config);
}

/**
 * Fetches the currently loaded active company from TallyPrime
 */
export async function fetchTallyCompany(config: ElectronTallyConfig): Promise<ElectronTallyResponse> {
  if (!isDirectTallyAvailable()) {
    return { success: false, error: 'Direct Tally connection is only available in Desktop Offline App mode.' };
  }
  return await window.electron.tally.fetchCompany(config);
}

/**
 * Parses general XML response from Tally to summarize action/import status
 */
export function parseTallyResponse(responseXml: string): TallyResponseSummary {
  try {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(responseXml, "text/xml");
    
    const parserError = xmlDoc.getElementsByTagName("parsererror");
    if (parserError.length > 0) {
      return { success: false, errors: [parserError[0].textContent || "XML parsing error"] };
    }

    // Check for Import Result
    const createdEl = xmlDoc.getElementsByTagName("CREATED");
    const alteredEl = xmlDoc.getElementsByTagName("ALTERED");
    const errorsEl = xmlDoc.getElementsByTagName("ERRORS");
    const lineErrors = xmlDoc.getElementsByTagName("LINEERROR");
    
    if (createdEl.length > 0 || alteredEl.length > 0 || errorsEl.length > 0) {
      const created = parseInt(createdEl[0]?.textContent || "0", 10);
      const altered = parseInt(alteredEl[0]?.textContent || "0", 10);
      const errorCount = parseInt(errorsEl[0]?.textContent || "0", 10);
      
      const errorsList: string[] = [];
      for (let i = 0; i < lineErrors.length; i++) {
        if (lineErrors[i].textContent) {
          errorsList.push(lineErrors[i].textContent!);
        }
      }
      
      return {
        success: errorCount === 0,
        createdCount: created,
        alteredCount: altered,
        errorCount: errorCount,
        errors: errorsList
      };
    }
    
    // Check for Company name indicators
    const cmpNameTags = ["SVCURRENTCOMPANY", "SVCOMPANYNAME", "COMPANYNAME", "REDCURRENTCOMPANY", "CURRENTCOMPANY"];
    for (const tag of cmpNameTags) {
      const els = xmlDoc.getElementsByTagName(tag);
      if (els.length > 0 && els[0].textContent) {
        return { success: true, companyName: els[0].textContent.trim() };
      }
    }
    
    const companyEls = xmlDoc.getElementsByTagName("COMPANY");
    if (companyEls.length > 0 && companyEls[0].textContent) {
      return { success: true, companyName: companyEls[0].textContent.trim() };
    }

    return { success: true };
  } catch (err: any) {
    return { success: false, errors: [err.message || String(err)] };
  }
}

/**
 * Parses Ledgers from Tally XML
 */
export function parseTallyLedgersXml(xml: string): LedgerMasterRow[] {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "text/xml");
  const ledgers: LedgerMasterRow[] = [];
  const ledgerNodes = doc.getElementsByTagName("LEDGER");
  
  for (let i = 0; i < ledgerNodes.length; i++) {
    const node = ledgerNodes[i];
    const ledgerName = node.getAttribute("NAME") || node.getElementsByTagName("NAME")[0]?.textContent || "";
    if (!ledgerName) continue;
    
    const underGroup = node.getElementsByTagName("PARENT")[0]?.textContent || "Primary";
    const openingBalance = node.getElementsByTagName("OPENINGBALANCE")[0]?.textContent || "";
    const mailingName = node.getElementsByTagName("MAILINGNAME")[0]?.textContent || ledgerName;
    
    const addressNodes = node.getElementsByTagName("ADDRESS");
    const address1 = addressNodes[0]?.textContent || "";
    const address2 = addressNodes[1]?.textContent || "";
    
    const state = node.getElementsByTagName("LEDSTATENAME")[0]?.textContent || "";
    const country = node.getElementsByTagName("COUNTRYNAME")[0]?.textContent || "India";
    const pincode = node.getElementsByTagName("PINCODE")[0]?.textContent || "";
    const pan = node.getElementsByTagName("INCOMETAXNUMBER")[0]?.textContent || "";
    const gstin = node.getElementsByTagName("PARTYGSTIN")[0]?.textContent || "";
    const registrationType = node.getElementsByTagName("GSTREGISTRATIONTYPE")[0]?.textContent || "";
    const email = node.getElementsByTagName("EMAIL")[0]?.textContent || "";
    
    ledgers.push({
      rowNum: i + 1,
      ledgerName,
      underGroup,
      openingBalance,
      mailingName,
      address1,
      address2,
      state,
      country,
      pincode,
      pan,
      gstin,
      registrationType,
      email,
      isValid: true,
      errors: [],
      warnings: [],
      isDuplicate: false,
      isPossibleDuplicate: false,
      excluded: false
    });
  }
  return ledgers;
}

/**
 * Parses Groups from Tally XML
 */
export function parseTallyGroupsXml(xml: string): { groups: string[], groupParentMap: Record<string, string> } {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "text/xml");
  const groups: string[] = [];
  const groupParentMap: Record<string, string> = {};
  
  const groupNodes = doc.getElementsByTagName("GROUP");
  for (let i = 0; i < groupNodes.length; i++) {
    const node = groupNodes[i];
    const groupName = node.getAttribute("NAME") || node.getElementsByTagName("NAME")[0]?.textContent || "";
    if (!groupName) continue;
    
    groups.push(groupName);
    const parent = node.getElementsByTagName("PARENT")[0]?.textContent || "";
    if (parent) {
      groupParentMap[groupName] = parent;
    }
  }
  return { groups, groupParentMap };
}

/**
 * Parses Stock Items from Tally XML
 */
export function parseTallyStockItemsXml(xml: string): StockMasterRow[] {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "text/xml");
  const items: StockMasterRow[] = [];
  const stockNodes = doc.getElementsByTagName("STOCKITEM");
  
  for (let i = 0; i < stockNodes.length; i++) {
    const node = stockNodes[i];
    const itemName = node.getAttribute("NAME") || node.getElementsByTagName("NAME")[0]?.textContent || "";
    if (!itemName) continue;
    
    const underGroup = node.getElementsByTagName("PARENT")[0]?.textContent || "Primary";
    const unit = node.getElementsByTagName("BASEUNITS")[0]?.textContent || "";
    const openingQty = node.getElementsByTagName("OPENINGBALANCE")[0]?.textContent || "";
    const openingRate = node.getElementsByTagName("OPENINGRATE")[0]?.textContent || "";
    const openingValue = node.getElementsByTagName("OPENINGVALUE")[0]?.textContent || "";
    const hsn = node.getElementsByTagName("HSNCODE")[0]?.textContent || "";
    const gstApplicable = node.getElementsByTagName("GSTAPPLICABLE")[0]?.textContent || "";
    const taxability = node.getElementsByTagName("TAXABILITY")[0]?.textContent || "";
    const description = node.getElementsByTagName("DESCRIPTION")[0]?.textContent || "";
    
    const igst = node.getElementsByTagName("IGSTRATE")[0]?.textContent || "";
    const cgst = node.getElementsByTagName("CGSTRATE")[0]?.textContent || "";
    const sgst = node.getElementsByTagName("SGSTRATE")[0]?.textContent || "";
    
    items.push({
      rowNum: i + 1,
      itemName,
      underGroup,
      unit,
      openingQty,
      openingRate,
      openingValue,
      hsn,
      gstApplicable,
      taxability,
      gstRate: igst,
      cgstRate: cgst,
      sgstRate: sgst,
      igstRate: igst,
      description,
      isValid: true,
      errors: [],
      warnings: []
    });
  }
  return items;
}

/**
 * Parses Units from Tally XML
 */
export function parseTallyUnitsXml(xml: string): string[] {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "text/xml");
  const units: string[] = [];
  const unitNodes = doc.getElementsByTagName("UNIT");
  
  for (let i = 0; i < unitNodes.length; i++) {
    const node = unitNodes[i];
    const name = node.getAttribute("NAME") || node.getElementsByTagName("NAME")[0]?.textContent || "";
    if (name) {
      units.push(name);
    }
  }
  return units;
}

/**
 * Parses Daybook Vouchers from Tally XML
 */
export function parseTallyDaybookXml(xml: string): TallyDaybookTransaction[] {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "text/xml");
  const transactions: TallyDaybookTransaction[] = [];
  const voucherNodes = doc.getElementsByTagName("VOUCHER");
  
  for (let i = 0; i < voucherNodes.length; i++) {
    const node = voucherNodes[i];
    const date = node.getElementsByTagName("DATE")[0]?.textContent || "";
    const voucherType = node.getElementsByTagName("VOUCHERTYPENAME")[0]?.textContent || "";
    const narration = node.getElementsByTagName("NARRATION")[0]?.textContent || "";
    const reference = node.getElementsByTagName("REFERENCE")[0]?.textContent || "";
    
    const ledgerList = node.getElementsByTagName("ALLLEDGERENTRIES.LIST");
    for (let j = 0; j < ledgerList.length; j++) {
      const entry = ledgerList[j];
      const ledger = entry.getElementsByTagName("LEDGERNAME")[0]?.textContent || "";
      const amountStr = entry.getElementsByTagName("AMOUNT")[0]?.textContent || "0";
      const amount = parseFloat(amountStr);
      
      if (ledger) {
        transactions.push({
          date,
          voucherType,
          narration,
          ledger,
          amount,
          reference
        });
      }
    }
  }
  return transactions;
}

/**
 * Fetches and normalizes all masters from local Tally to fit existing TallyContext structure
 */
export async function fetchTallyMasters(config: ElectronTallyConfig): Promise<any> {
  if (!isDirectTallyAvailable()) {
    throw new Error("Direct Tally connection is not available in Web mode.");
  }
  
  // 1. Fetch groups
  const groupsRes = await window.electron.tally.fetchGroups(config, buildExportGroupsRequest());
  if (!groupsRes.success) throw new Error("Failed to fetch Groups from Tally: " + groupsRes.error);
  
  // 2. Fetch ledgers
  const ledgersRes = await window.electron.tally.fetchLedgers(config, buildExportLedgersRequest());
  if (!ledgersRes.success) throw new Error("Failed to fetch Ledgers from Tally: " + ledgersRes.error);
  
  // 3. Fetch stock items
  const stockItemsRes = await window.electron.tally.fetchStockItems(config, buildExportStockItemsRequest());
  if (!stockItemsRes.success) throw new Error("Failed to fetch Stock Items from Tally: " + stockItemsRes.error);
  
  // 4. Fetch units
  const unitsRes = await window.electron.tally.fetchUnits(config, buildExportUnitsRequest());
  if (!unitsRes.success) throw new Error("Failed to fetch Units from Tally: " + unitsRes.error);
  
  // Parse them
  const parsedGroups = parseTallyGroupsXml(groupsRes.response || "");
  const parsedLedgers = parseTallyLedgersXml(ledgersRes.response || "");
  const parsedStockItems = parseTallyStockItemsXml(stockItemsRes.response || "");
  const parsedUnits = parseTallyUnitsXml(unitsRes.response || "");
  
  const ledgerGroupMap: Record<string, string> = {};
  parsedLedgers.forEach(l => {
    ledgerGroupMap[l.ledgerName] = l.underGroup;
  });
  
  const stockItemStockGroupMap: Record<string, string> = {};
  parsedStockItems.forEach(si => {
    stockItemStockGroupMap[si.itemName] = si.underGroup;
  });
  
  return {
    ledgers: parsedLedgers.map(l => l.ledgerName),
    groups: parsedGroups.groups,
    groupParentMap: parsedGroups.groupParentMap,
    stockItems: parsedStockItems.map(si => si.itemName),
    stockGroups: Array.from(new Set(parsedStockItems.map(si => si.underGroup))),
    units: parsedUnits,
    ledgerGroupMap,
    stockItemStockGroupMap,
    ledgerDetails: parsedLedgers,
    stockItemDetails: parsedStockItems
  };
}

/**
 * Fetches Daybook transactions and converts to mapping transactions
 */
export async function fetchTallyDaybook(
  config: ElectronTallyConfig,
  fromDate: string,
  toDate: string
): Promise<TallyDaybookTransaction[]> {
  if (!isDirectTallyAvailable()) {
    throw new Error("Direct Tally connection is not available in Web mode.");
  }
  const xml = buildExportDaybookRequest(fromDate, toDate);
  const response = await window.electron.tally.fetchDaybook(config, xml);
  if (!response.success) {
    throw new Error("Failed to fetch Daybook: " + response.error);
  }
  return parseTallyDaybookXml(response.response || "");
}

/**
 * Pushes generated XML payload directly to local TallyPrime
 */
export async function pushXmlToTally(config: ElectronTallyConfig, xml: string): Promise<ElectronTallyResponse> {
  if (!isDirectTallyAvailable()) {
    return { success: false, error: 'Direct Tally connection is only available in Desktop Offline App mode.' };
  }
  const payload = buildImportXmlRequest(xml);
  return await window.electron.tally.pushXml(config, payload);
}

export interface FetchMasterFailure {
  type: 'Ledger' | 'Group' | 'StockItem' | 'Unit' | 'VoucherType';
  name: string;
  reason: string;
}

export interface MastersFetchSummary {
  companyName: string;
  fetchedAt: number;
  ledgers: { total: number; fetched: number; failed: number };
  groups: { total: number; fetched: number; failed: number };
  stockItems: { total: number; fetched: number; failed: number };
  units: { total: number; fetched: number; failed: number };
  voucherTypes: { total: number; fetched: number; failed: number };
  failures: FetchMasterFailure[];
}

export interface FetchTransactionFailure {
  date: string;
  voucherNo: string;
  voucherType: string;
  reason: string;
}

export interface DaybookFetchSummary {
  companyName: string;
  fromDate: string;
  toDate: string;
  totalVouchers: number;
  fetchedVouchers: number;
  failedVouchers: number;
  total: number;
  fetched: number;
  failed: number;
  failures: FetchTransactionFailure[];
}

/**
 * XML builder to fetch all loaded/open companies from TallyPrime
 */
export function buildExportCompaniesRequest(): string {
  return `
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>ListofCompanies</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="ListofCompanies">
            <TYPE>Company</TYPE>
            <FETCH>Name, GUID</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>`.trim();
}

/**
 * XML builder to fetch all Voucher Types for a company
 */
export function buildExportVoucherTypesRequest(companyName?: string): string {
  const companyVar = companyName ? `<SVCOMPANYNAME>${companyName}</SVCOMPANYNAME>` : '';
  return `
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>VoucherTypeCollection</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        ${companyVar}
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="VoucherTypeCollection">
            <TYPE>VoucherType</TYPE>
            <FETCH>Name</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>`.trim();
}

/**
 * Parses company list response from Tally
 */
export function parseCompanyListResponse(responseXml: string): { name: string; guid: string; isActive: boolean }[] {
  try {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(responseXml, "text/xml");
    
    const parserError = xmlDoc.getElementsByTagName("parsererror");
    if (parserError.length > 0) {
      console.error("XML parse error in company list", parserError[0].textContent);
    }
    
    const companies: { name: string; guid: string; isActive: boolean }[] = [];
    const companyNodes = xmlDoc.getElementsByTagName("COMPANY");
    
    if (companyNodes.length > 0) {
      for (let i = 0; i < companyNodes.length; i++) {
        const node = companyNodes[i];
        const name = node.getAttribute("NAME") || node.getElementsByTagName("NAME")[0]?.textContent || "";
        const guid = node.getElementsByTagName("GUID")[0]?.textContent || "";
        if (name) {
          companies.push({ name: name.trim(), guid: guid.trim(), isActive: true });
        }
      }
    } else {
      const activeCompanyTags = ["SVCURRENTCOMPANY", "SVCOMPANYNAME", "COMPANYNAME", "REDCURRENTCOMPANY", "CURRENTCOMPANY"];
      for (const tag of activeCompanyTags) {
        const els = xmlDoc.getElementsByTagName(tag);
        if (els.length > 0 && els[0].textContent) {
          companies.push({ name: els[0].textContent.trim(), guid: "", isActive: true });
          break;
        }
      }
    }
    
    const uniqueMap = new Map<string, { name: string; guid: string; isActive: boolean }>();
    companies.forEach(c => uniqueMap.set(c.name.toLowerCase(), c));
    return Array.from(uniqueMap.values());
  } catch (err) {
    console.error("Error parsing company list response", err);
    return [];
  }
}

/**
 * Fetches all available companies in local Tally
 */
export async function fetchCompaniesFromTally(config: ElectronTallyConfig): Promise<{ name: string; guid: string; isActive: boolean }[]> {
  if (!isDirectTallyAvailable()) {
    throw new Error("Direct Tally connection is only available in Desktop Offline App.");
  }
  const xml = buildExportCompaniesRequest();
  const res = await window.electron.tally.fetchCompanies(config, xml);
  if (!res.success) {
    throw new Error(res.error || "Failed to fetch companies from Tally.");
  }
  return parseCompanyListResponse(res.response || "");
}

/**
 * Builds scoped company XML request
 */
export function buildCompanyScopedExportRequest(
  companyName: string,
  reportType: 'Ledgers' | 'Groups' | 'StockItems' | 'Units' | 'Daybook',
  params?: any
): string {
  const companyVar = companyName ? `<SVCOMPANYNAME>${companyName}</SVCOMPANYNAME>` : '';
  
  if (reportType === 'Ledgers') {
    return `
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>LedgerCollection</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        ${companyVar}
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="LedgerCollection">
            <TYPE>Ledger</TYPE>
            <FETCH>Name, Parent, OpeningBalance, MailingName, Address, LEDState, PINCode, IncomeTaxNumber, GSTRegistrationType, PartyGSTIN, Email, Phone</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>`.trim();
  } else if (reportType === 'Groups') {
    return `
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>GroupCollection</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        ${companyVar}
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="GroupCollection">
            <TYPE>Group</TYPE>
            <FETCH>Name, Parent</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>`.trim();
  } else if (reportType === 'StockItems') {
    return `
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>StockItemCollection</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        ${companyVar}
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="StockItemCollection">
            <TYPE>StockItem</TYPE>
            <FETCH>Name, Parent, BaseUnits, OpeningBalance, OpeningRate, OpeningValue, HSNCode, GSTApplicable, TAXABILITY, Description, IGSTRate, CGSTRate, SGSTRate</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>`.trim();
  } else if (reportType === 'Units') {
    return `
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>UnitCollection</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        ${companyVar}
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="UnitCollection">
            <TYPE>Unit</TYPE>
            <FETCH>Name, OriginalName, DecimalPlaces</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>`.trim();
  } else {
    const fromDate = params?.fromDate || '';
    const toDate = params?.toDate || '';
    const fDateObj = normalizeTallyDate(fromDate);
    const fDate = fDateObj.isValid ? fDateObj.value : fromDate.replace(/[-/]/g, '');
    const tDateObj = normalizeTallyDate(toDate);
    const tDate = tDateObj.isValid ? tDateObj.value : toDate.replace(/[-/]/g, '');
    return `
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>DaybookCollection</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVFROMDATE>${fDate}</SVFROMDATE>
        <SVTODATE>${tDate}</SVTODATE>
        ${companyVar}
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="DaybookCollection">
            <TYPE>Voucher</TYPE>
            <FETCH>Date, VoucherTypeName, VoucherNumber, Narration, Reference, LedgerEntries, AllLedgerEntries</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>`.trim();
  }
}

/**
 * Parses fetched masters xmls with complete failure/skip details
 */
export function parseMastersWithSummary(
  companyName: string,
  xmls: { groupsXml: string; ledgersXml: string; stockItemsXml: string; unitsXml: string; voucherTypesXml: string }
): {
  parsedMasters: any;
  summary: MastersFetchSummary;
} {
  const failures: FetchMasterFailure[] = [];
  
  // 1. Groups
  const groupParser = new DOMParser();
  const groupDoc = groupParser.parseFromString(xmls.groupsXml, "text/xml");
  const groupNodes = groupDoc.getElementsByTagName("GROUP");
  const groups: string[] = [];
  const groupParentMap: Record<string, string> = {};
  
  let totalGroups = groupNodes.length;
  let fetchedGroups = 0;
  
  for (let i = 0; i < groupNodes.length; i++) {
    const node = groupNodes[i];
    const name = node.getAttribute("NAME") || node.getElementsByTagName("NAME")[0]?.textContent || "";
    if (!name) {
      failures.push({ type: 'Group', name: `Group #${i+1}`, reason: 'Missing name' });
      continue;
    }
    const parent = node.getElementsByTagName("PARENT")[0]?.textContent || "";
    if (groups.includes(name)) {
      failures.push({ type: 'Group', name, reason: 'Duplicate master ignored' });
      continue;
    }
    groups.push(name);
    if (parent) {
      groupParentMap[name] = parent;
    }
    fetchedGroups++;
  }
  
  // 2. Ledgers
  const ledgerParser = new DOMParser();
  const ledgerDoc = ledgerParser.parseFromString(xmls.ledgersXml, "text/xml");
  const ledgerNodes = ledgerDoc.getElementsByTagName("LEDGER");
  const ledgers: LedgerMasterRow[] = [];
  const ledgerNames: string[] = [];
  const ledgerGroupMap: Record<string, string> = {};
  
  let totalLedgers = ledgerNodes.length;
  let fetchedLedgers = 0;
  
  for (let i = 0; i < ledgerNodes.length; i++) {
    const node = ledgerNodes[i];
    const name = node.getAttribute("NAME") || node.getElementsByTagName("NAME")[0]?.textContent || "";
    if (!name) {
      failures.push({ type: 'Ledger', name: `Ledger #${i+1}`, reason: 'Missing name' });
      continue;
    }
    const underGroup = node.getElementsByTagName("PARENT")[0]?.textContent || "";
    if (!underGroup) {
      failures.push({ type: 'Ledger', name, reason: 'Parent group missing' });
      continue;
    }
    if (ledgerNames.includes(name)) {
      failures.push({ type: 'Ledger', name, reason: 'Duplicate master ignored' });
      continue;
    }
    
    ledgerNames.push(name);
    ledgerGroupMap[name] = underGroup;
    
    const gstin = node.getElementsByTagName("PARTYGSTIN")[0]?.textContent || "";
    const regType = node.getElementsByTagName("GSTREGISTRATIONTYPE")[0]?.textContent || "";
    if (regType && regType !== 'Unregistered' && gstin && gstin.length !== 15) {
      failures.push({ type: 'Ledger', name, reason: 'Invalid GSTIN length (warning)' });
    }
    
    const openingBalance = node.getElementsByTagName("OPENINGBALANCE")[0]?.textContent || "";
    const mailingName = node.getElementsByTagName("MAILINGNAME")[0]?.textContent || name;
    const addressNodes = node.getElementsByTagName("ADDRESS");
    const address1 = addressNodes[0]?.textContent || "";
    const address2 = addressNodes[1]?.textContent || "";
    const state = node.getElementsByTagName("LEDSTATENAME")[0]?.textContent || "";
    const country = node.getElementsByTagName("COUNTRYNAME")[0]?.textContent || "India";
    const pincode = node.getElementsByTagName("PINCODE")[0]?.textContent || "";
    const pan = node.getElementsByTagName("INCOMETAXNUMBER")[0]?.textContent || "";
    const email = node.getElementsByTagName("EMAIL")[0]?.textContent || "";
    
    ledgers.push({
      rowNum: ledgers.length + 1,
      ledgerName: name,
      underGroup,
      openingBalance,
      mailingName,
      address1,
      address2,
      state,
      country,
      pincode,
      pan,
      gstin,
      registrationType: regType,
      isValid: true,
      errors: [],
      warnings: [],
      isDuplicate: false,
      isPossibleDuplicate: false,
      excluded: false
    });
    fetchedLedgers++;
  }
  
  // 3. Stock Items
  const stockParser = new DOMParser();
  const stockDoc = stockParser.parseFromString(xmls.stockItemsXml, "text/xml");
  const stockNodes = stockDoc.getElementsByTagName("STOCKITEM");
  const stockItems: StockMasterRow[] = [];
  const stockItemNames: string[] = [];
  const stockItemStockGroupMap: Record<string, string> = {};
  
  let totalStockItems = stockNodes.length;
  let fetchedStockItems = 0;
  
  for (let i = 0; i < stockNodes.length; i++) {
    const node = stockNodes[i];
    const name = node.getAttribute("NAME") || node.getElementsByTagName("NAME")[0]?.textContent || "";
    if (!name) {
      failures.push({ type: 'StockItem', name: `StockItem #${i+1}`, reason: 'Missing name' });
      continue;
    }
    const underGroup = node.getElementsByTagName("PARENT")[0]?.textContent || "Primary";
    const unit = node.getElementsByTagName("BASEUNITS")[0]?.textContent || "";
    if (!unit) {
      failures.push({ type: 'StockItem', name, reason: 'Invalid stock unit / missing unit mapping' });
      continue;
    }
    if (stockItemNames.includes(name)) {
      failures.push({ type: 'StockItem', name, reason: 'Duplicate master ignored' });
      continue;
    }
    
    stockItemNames.push(name);
    stockItemStockGroupMap[name] = underGroup;
    
    const openingQty = node.getElementsByTagName("OPENINGBALANCE")[0]?.textContent || "";
    const openingRate = node.getElementsByTagName("OPENINGRATE")[0]?.textContent || "";
    const openingValue = node.getElementsByTagName("OPENINGVALUE")[0]?.textContent || "";
    const hsn = node.getElementsByTagName("HSNCODE")[0]?.textContent || "";
    const gstApplicable = node.getElementsByTagName("GSTAPPLICABLE")[0]?.textContent || "";
    const taxability = node.getElementsByTagName("TAXABILITY")[0]?.textContent || "";
    const description = node.getElementsByTagName("DESCRIPTION")[0]?.textContent || "";
    const igst = node.getElementsByTagName("IGSTRATE")[0]?.textContent || "";
    const cgst = node.getElementsByTagName("CGSTRATE")[0]?.textContent || "";
    const sgst = node.getElementsByTagName("SGSTRATE")[0]?.textContent || "";
    
    stockItems.push({
      rowNum: stockItems.length + 1,
      itemName: name,
      underGroup,
      unit,
      openingQty,
      openingRate,
      openingValue,
      hsn,
      gstApplicable,
      taxability,
      gstRate: igst,
      cgstRate: cgst,
      sgstRate: sgst,
      igstRate: igst,
      description,
      isValid: true,
      errors: [],
      warnings: []
    });
    fetchedStockItems++;
  }
  
  // 4. Units
  const unitParser = new DOMParser();
  const unitDoc = unitParser.parseFromString(xmls.unitsXml, "text/xml");
  const unitNodes = unitDoc.getElementsByTagName("UNIT");
  const units: string[] = [];
  
  let totalUnits = unitNodes.length;
  let fetchedUnits = 0;
  
  for (let i = 0; i < unitNodes.length; i++) {
    const node = unitNodes[i];
    const name = node.getAttribute("NAME") || node.getElementsByTagName("NAME")[0]?.textContent || "";
    if (!name) {
      failures.push({ type: 'Unit', name: `Unit #${i+1}`, reason: 'Missing unit name' });
      continue;
    }
    if (units.includes(name)) {
      failures.push({ type: 'Unit', name, reason: 'Duplicate master ignored' });
      continue;
    }
    units.push(name);
    fetchedUnits++;
  }
  
  // 5. Voucher Types
  const vtParser = new DOMParser();
  const vtDoc = vtParser.parseFromString(xmls.voucherTypesXml, "text/xml");
  const vtNodes = vtDoc.getElementsByTagName("VOUCHERTYPE");
  const voucherTypes: string[] = [];
  
  let totalVoucherTypes = vtNodes.length;
  let fetchedVoucherTypes = 0;
  
  for (let i = 0; i < vtNodes.length; i++) {
    const node = vtNodes[i];
    const name = node.getAttribute("NAME") || node.getElementsByTagName("NAME")[0]?.textContent || "";
    if (!name) {
      failures.push({ type: 'VoucherType', name: `VoucherType #${i+1}`, reason: 'Missing voucher type name' });
      continue;
    }
    if (voucherTypes.includes(name)) {
      failures.push({ type: 'VoucherType', name, reason: 'Duplicate master ignored' });
      continue;
    }
    voucherTypes.push(name);
    fetchedVoucherTypes++;
  }
  
  const summary: MastersFetchSummary = {
    companyName,
    fetchedAt: Date.now(),
    ledgers: { total: totalLedgers, fetched: fetchedLedgers, failed: totalLedgers - fetchedLedgers },
    groups: { total: totalGroups, fetched: fetchedGroups, failed: totalGroups - fetchedGroups },
    stockItems: { total: totalStockItems, fetched: fetchedStockItems, failed: totalStockItems - fetchedStockItems },
    units: { total: totalUnits, fetched: fetchedUnits, failed: totalUnits - fetchedUnits },
    voucherTypes: { total: totalVoucherTypes, fetched: fetchedVoucherTypes, failed: totalVoucherTypes - fetchedVoucherTypes },
    failures
  };
  
  const parsedMasters = {
    ledgers: ledgerNames,
    groups,
    groupParentMap,
    stockItems: stockItemNames,
    stockGroups: Array.from(new Set(stockItems.map(si => si.underGroup))),
    units,
    ledgerGroupMap,
    stockItemStockGroupMap,
    ledgerDetails: ledgers,
    stockItemDetails: stockItems
  };
  
  return { parsedMasters, summary };
}

/**
 * Parses fetched Daybook vouchers with failure details
 */
export function parseDaybookWithSummary(
  companyName: string,
  fromDate: string,
  toDate: string,
  xml: string
): {
  transactions: TallyDaybookTransaction[];
  summary: DaybookFetchSummary;
} {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "text/xml");
  const transactions: TallyDaybookTransaction[] = [];
  const failures: FetchTransactionFailure[] = [];
  
  const voucherNodes = doc.getElementsByTagName("VOUCHER");
  let totalVouchers = voucherNodes.length;
  let fetchedVouchers = 0;
  
  for (let i = 0; i < voucherNodes.length; i++) {
    const node = voucherNodes[i];
    const date = node.getElementsByTagName("DATE")[0]?.textContent || "";
    const voucherType = node.getElementsByTagName("VOUCHERTYPENAME")[0]?.textContent || "";
    const voucherNo = node.getElementsByTagName("VOUCHERNUMBER")[0]?.textContent || `Index #${i+1}`;
    const narration = node.getElementsByTagName("NARRATION")[0]?.textContent || "";
    const reference = node.getElementsByTagName("REFERENCE")[0]?.textContent || "";
    
    const normalizedResult = normalizeTallyDate(date);
    if (!normalizedResult.isValid) {
      failures.push({
        date: date || 'N/A',
        voucherNo,
        voucherType: voucherType || 'N/A',
        reason: normalizedResult.error || 'Invalid date format (must be YYYYMMDD)'
      });
      continue;
    }
    const finalDate = normalizedResult.value;
    
    if (!voucherType) {
      failures.push({
        date: finalDate,
        voucherNo,
        voucherType: 'N/A',
        reason: 'Missing voucher type'
      });
      continue;
    }
    
    const ledgerList = node.getElementsByTagName("ALLLEDGERENTRIES.LIST");
    if (ledgerList.length === 0) {
      failures.push({
        date: finalDate,
        voucherNo,
        voucherType,
        reason: 'Ledger entries missing'
      });
      continue;
    }
    
    let hasLedger = false;
    let hasZeroAmount = false;
    const tempEntries: TallyDaybookTransaction[] = [];
    
    for (let j = 0; j < ledgerList.length; j++) {
      const entry = ledgerList[j];
      const ledger = entry.getElementsByTagName("LEDGERNAME")[0]?.textContent || "";
      const amountStr = entry.getElementsByTagName("AMOUNT")[0]?.textContent || "0";
      const amount = parseFloat(amountStr);
      
      if (!ledger) {
        continue;
      }
      
      hasLedger = true;
      if (amount === 0 || isNaN(amount)) {
        hasZeroAmount = true;
      }
      
      tempEntries.push({
        date: finalDate,
        voucherType,
        narration,
        ledger,
        amount,
        reference
      });
    }
    
    if (!hasLedger) {
      failures.push({
        date: finalDate,
        voucherNo,
        voucherType,
        reason: 'Missing ledger names in entries'
      });
      continue;
    }
    
    if (hasZeroAmount) {
      failures.push({
        date: finalDate,
        voucherNo,
        voucherType,
        reason: 'Zero amount voucher ignored'
      });
      continue;
    }
    
    transactions.push(...tempEntries);
    fetchedVouchers++;
  }
  
  const summary: DaybookFetchSummary = {
    companyName,
    fromDate,
    toDate,
    totalVouchers,
    fetchedVouchers,
    failedVouchers: totalVouchers - fetchedVouchers,
    total: totalVouchers,
    fetched: fetchedVouchers,
    failed: totalVouchers - fetchedVouchers,
    failures
  };
  
  return { transactions, summary };
}

/**
 * Fetches all Tally masters for selected company with detailed summary
 */
export async function fetchMastersForCompany(
  config: ElectronTallyConfig,
  companyName: string
): Promise<{ parsedMasters: any; summary: MastersFetchSummary }> {
  if (!isDirectTallyAvailable()) {
    throw new Error("Direct Tally connection is not available in Web mode.");
  }
  
  const groupsRes = await window.electron.tally.fetchMastersForCompany(config, buildCompanyScopedExportRequest(companyName, 'Groups'));
  if (!groupsRes.success) throw new Error("Failed to fetch Groups: " + groupsRes.error);
  
  const ledgersRes = await window.electron.tally.fetchMastersForCompany(config, buildCompanyScopedExportRequest(companyName, 'Ledgers'));
  if (!ledgersRes.success) throw new Error("Failed to fetch Ledgers: " + ledgersRes.error);
  
  const stockItemsRes = await window.electron.tally.fetchMastersForCompany(config, buildCompanyScopedExportRequest(companyName, 'StockItems'));
  if (!stockItemsRes.success) throw new Error("Failed to fetch Stock Items: " + stockItemsRes.error);
  
  const unitsRes = await window.electron.tally.fetchMastersForCompany(config, buildCompanyScopedExportRequest(companyName, 'Units'));
  if (!unitsRes.success) throw new Error("Failed to fetch Units: " + unitsRes.error);
  
  const vtRes = await window.electron.tally.testConnection(config, buildExportVoucherTypesRequest(companyName));
  if (!vtRes.success) throw new Error("Failed to fetch Voucher Types: " + vtRes.error);
  
  return parseMastersWithSummary(companyName, {
    groupsXml: groupsRes.response || "",
    ledgersXml: ledgersRes.response || "",
    stockItemsXml: stockItemsRes.response || "",
    unitsXml: unitsRes.response || "",
    voucherTypesXml: vtRes.response || ""
  });
}

/**
 * Fetches company scoped Daybook transactions with detailed summary
 */
export async function fetchDaybookForCompany(
  config: ElectronTallyConfig,
  companyName: string,
  fromDate: string,
  toDate: string
): Promise<{ transactions: TallyDaybookTransaction[]; summary: DaybookFetchSummary }> {
  if (!isDirectTallyAvailable()) {
    throw new Error("Direct Tally connection is not available in Web mode.");
  }
  const xml = buildCompanyScopedExportRequest(companyName, 'Daybook', { fromDate, toDate });
  const response = await window.electron.tally.fetchDaybookForCompany(config, xml);
  if (!response.success) {
    throw new Error("Failed to fetch Daybook: " + response.error);
  }
  return parseDaybookWithSummary(companyName, fromDate, toDate, response.response || "");
}

/**
 * Exports Fetch Report as Excel workbook with Summary, Fetched and Failed sheets
 */
export function generateFetchReportExcel(
  mastersSummary?: MastersFetchSummary,
  daybookSummary?: DaybookFetchSummary,
  parsedMastersData?: any,
  parsedTransactions?: TallyDaybookTransaction[]
) {
  const wb = XLSX.utils.book_new();
  
  // Sheet 1: Summary
  const summaryRows: any[] = [];
  summaryRows.push(['TALLYGEN PRO - FETCH REPORT', '']);
  summaryRows.push(['Generated At', new Date().toLocaleString()]);
  summaryRows.push(['', '']);
  
  if (mastersSummary) {
    summaryRows.push(['MASTERS FETCH SUMMARY', '']);
    summaryRows.push(['Company Name', mastersSummary.companyName]);
    summaryRows.push(['Fetched At', new Date(mastersSummary.fetchedAt).toLocaleString()]);
    summaryRows.push(['Master Type', 'Available in Tally', 'Fetched Successfully', 'Failed/Skipped']);
    summaryRows.push(['Ledgers', mastersSummary.ledgers.total, mastersSummary.ledgers.fetched, mastersSummary.ledgers.failed]);
    summaryRows.push(['Groups', mastersSummary.groups.total, mastersSummary.groups.fetched, mastersSummary.groups.failed]);
    summaryRows.push(['Stock Items', mastersSummary.stockItems.total, mastersSummary.stockItems.fetched, mastersSummary.stockItems.failed]);
    summaryRows.push(['Units', mastersSummary.units.total, mastersSummary.units.fetched, mastersSummary.units.failed]);
    summaryRows.push(['Voucher Types', mastersSummary.voucherTypes.total, mastersSummary.voucherTypes.fetched, mastersSummary.voucherTypes.failed]);
    summaryRows.push(['', '', '', '']);
  }
  
  if (daybookSummary) {
    summaryRows.push(['DAYBOOK TRANSACTION FETCH SUMMARY', '']);
    summaryRows.push(['Company Name', daybookSummary.companyName]);
    summaryRows.push(['Date Range', `${daybookSummary.fromDate} to ${daybookSummary.toDate}`]);
    summaryRows.push(['Metric', 'Count']);
    summaryRows.push(['Total Vouchers Found', daybookSummary.totalVouchers]);
    summaryRows.push(['Successfully Parsed', daybookSummary.fetchedVouchers]);
    summaryRows.push(['Skipped / Failed', daybookSummary.failedVouchers]);
  }
  
  const summaryWs = XLSX.utils.aoa_to_sheet(summaryRows);
  XLSX.utils.book_append_sheet(wb, summaryWs, 'Summary');
  
  // Sheet 2: Fetched Masters (Dynamic)
  if (parsedMastersData && (
    (parsedMastersData.ledgers && parsedMastersData.ledgers.length > 0) ||
    (parsedMastersData.groups && parsedMastersData.groups.length > 0) ||
    (parsedMastersData.stockItems && parsedMastersData.stockItems.length > 0) ||
    (parsedMastersData.units && parsedMastersData.units.length > 0)
  )) {
    const fetchedMastersRows: any[] = [];
    fetchedMastersRows.push(['Master Type', 'Master Name', 'Parent Group/Unit/Details']);
    if (parsedMastersData.ledgers) {
      parsedMastersData.ledgers.forEach((l: string) => {
        const parent = parsedMastersData.ledgerGroupMap?.[l] || '';
        fetchedMastersRows.push(['Ledger', l, parent]);
      });
    }
    if (parsedMastersData.groups) {
      parsedMastersData.groups.forEach((g: string) => {
        const parent = parsedMastersData.groupParentMap?.[g] || '';
        fetchedMastersRows.push(['Group', g, parent]);
      });
    }
    if (parsedMastersData.stockItems) {
      parsedMastersData.stockItems.forEach((si: string) => {
        const parent = parsedMastersData.stockItemStockGroupMap?.[si] || '';
        fetchedMastersRows.push(['Stock Item', si, parent]);
      });
    }
    if (parsedMastersData.units) {
      parsedMastersData.units.forEach((u: string) => {
        fetchedMastersRows.push(['Unit', u, '']);
      });
    }
    const fetchedMastersWs = XLSX.utils.aoa_to_sheet(fetchedMastersRows);
    XLSX.utils.book_append_sheet(wb, fetchedMastersWs, 'Fetched Masters');
  }
  
  // Sheet 3: Failed Masters (Dynamic)
  if (mastersSummary && mastersSummary.failures && mastersSummary.failures.length > 0) {
    const failedMastersRows: any[] = [];
    failedMastersRows.push(['Master Type', 'Name', 'Reason for Failure / Skip']);
    mastersSummary.failures.forEach(f => {
      failedMastersRows.push([f.type, f.name, f.reason]);
    });
    const failedMastersWs = XLSX.utils.aoa_to_sheet(failedMastersRows);
    XLSX.utils.book_append_sheet(wb, failedMastersWs, 'Failed Masters');
  }
  
  // Sheet 4: Fetched Transactions (Dynamic)
  if (parsedTransactions && parsedTransactions.length > 0) {
    const fetchedTxRows: any[] = [];
    fetchedTxRows.push(['Date', 'Voucher Type', 'Ledger Name', 'Amount', 'Narration', 'Reference']);
    parsedTransactions.forEach(t => {
      fetchedTxRows.push([t.date, t.voucherType, t.ledger, t.amount, t.narration, t.reference]);
    });
    const fetchedTxWs = XLSX.utils.aoa_to_sheet(fetchedTxRows);
    XLSX.utils.book_append_sheet(wb, fetchedTxWs, 'Fetched Transactions');
  }
  
  // Sheet 5: Failed Transactions (Dynamic)
  if (daybookSummary && daybookSummary.failures && daybookSummary.failures.length > 0) {
    const failedTxRows: any[] = [];
    failedTxRows.push(['Voucher Date', 'Voucher Number', 'Voucher Type', 'Reason for Failure / Skip']);
    daybookSummary.failures.forEach(f => {
      failedTxRows.push([f.date, f.voucherNo, f.voucherType, f.reason]);
    });
    const failedTxWs = XLSX.utils.aoa_to_sheet(failedTxRows);
    XLSX.utils.book_append_sheet(wb, failedTxWs, 'Failed Transactions');
  }
  
  XLSX.writeFile(wb, `Tally_Fetch_Report_${Date.now()}.xlsx`);
}
