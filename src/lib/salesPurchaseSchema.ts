export interface SalesPurchaseItem {
  stockItem: string;
  description?: string;
  quantity: number;
  unit: string;
  rate: number;
  itemAmount: number;
  discountPercent?: number;
  discountAmount?: number;
  taxableValue?: number;
  hsn?: string;
  gstRate: number;
  cgstLedger?: string;
  cgstAmount: number;
  sgstLedger?: string;
  sgstAmount: number;
  igstLedger?: string;
  igstAmount: number;
}

export interface SalesPurchaseCharge {
  ledger: string;
  amount: number;
  isDiscount?: boolean;
}

export interface SalesPurchaseInvoice {
  rowNum: number;
  voucherType: string;         // 'Sales' or 'Purchase'
  voucherMode: string;         // 'Accounting' or 'Item Invoice' or 'Auto'
  inventoryMode: string;       // 'No Inventory' or 'Inventory Optional' or 'Inventory Mandatory'
  invoiceDate: string;
  invoiceNo: string;
  partyLedger: string;
  salesPurchaseLedger: string;
  items: SalesPurchaseItem[];
  charges: SalesPurchaseCharge[];
  
  // Tax details
  gstMode: 'Auto' | 'Manual';
  gstRate?: number;
  hsn?: string;
  
  // Freight / extra charges
  freightLedger?: string;
  freightAmount?: number;
  packingLedger?: string;
  packingAmount?: number;
  loadingLedger?: string;
  loadingAmount?: number;
  insuranceLedger?: string;
  insuranceAmount?: number;
  otherLedger1?: string;
  otherAmount1?: number;
  otherLedger2?: string;
  otherAmount2?: number;
  discountLedger?: string;
  discountAmount?: number;
  roundOffLedger?: string;
  roundOffAmount?: number;

  // Party Details
  partyGSTIN?: string;
  partyAddress1?: string;
  partyAddress2?: string;
  partyState?: string;
  placeOfSupply?: string;
  partyRegistrationType?: string; // 'Regular', 'Unregistered', etc.

  // Dispatch / Transport
  dispatchDate?: string;
  deliveryNoteNo?: string;
  dispatchDocNo?: string;
  biltyLRNo?: string;
  transporterName?: string;
  transporterGSTIN?: string;
  vehicleNo?: string;
  destination?: string;
  modeOfTransport?: string;
  ewayBillNo?: string;

  // Narration & Reference
  narration?: string;
  reference?: string;

  // Error tracking & verification fields
  errors: string[];
  warnings: string[];
  isValid: boolean;
  totalTaxableValue: number;
  totalCGST: number;
  totalSGST: number;
  totalIGST: number;
  totalAdditionalCharges: number;
  invoiceTotal: number;
}

export interface ProfileDefinition {
  id: string;
  name: string;
  description: string;
  type: 'itemwise' | 'voucherwise' | 'custom';
  columns: string[];
}
