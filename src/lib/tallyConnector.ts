import { getAppMode } from './storageAdapter';
import { ElectronTallyConfig, ElectronTallyResponse } from '../types/electron';

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
  const fDate = fromDate.replace(/-/g, '');
  const tDate = toDate.replace(/-/g, '');
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
