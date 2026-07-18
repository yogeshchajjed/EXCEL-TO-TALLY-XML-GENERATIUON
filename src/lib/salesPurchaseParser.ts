import { getNormalizedField } from './columnAliases';
import { SalesPurchaseInvoice, SalesPurchaseItem, SalesPurchaseCharge } from './salesPurchaseSchema';

export function parseSalesPurchaseExcel(
  rawData: any[],
  profileId: string
): SalesPurchaseInvoice[] {
  if (!rawData || rawData.length === 0) return [];

  // Map each raw row to a dictionary with normalized keys
  const normalizedRows = rawData.map((rawRow, index) => {
    const rowDict: Record<string, any> = {};
    for (const [key, val] of Object.entries(rawRow)) {
      const field = getNormalizedField(key);
      if (field) {
        rowDict[field] = val;
      }
    }
    rowDict._rawRow = rawRow;
    rowDict._rowNum = index + 2;
    return rowDict;
  });

  const getNum = (v: any): number => {
    if (v === undefined || v === null || String(v).trim() === '') return 0;
    const n = Number(v);
    return isNaN(n) ? 0 : n;
  };

  const getStr = (v: any): string => {
    if (v === undefined || v === null) return '';
    return String(v).trim();
  };

  const invoicesMap = new Map<string, SalesPurchaseInvoice>();

  for (const row of normalizedRows) {
    const invoiceNo = getStr(row.invoiceNo);
    const voucherType = getStr(row.voucherType || 'Sales');
    
    if (!invoiceNo) continue;

    const key = `${voucherType}_${invoiceNo}`.toLowerCase();

    let inv = invoicesMap.get(key);
    if (!inv) {
      inv = {
        rowNum: row._rowNum,
        voucherType: voucherType,
        voucherMode: getStr(row.voucherMode || 'Auto'),
        inventoryMode: getStr(row.inventoryMode || 'Inventory Optional'),
        invoiceDate: getStr(row.invoiceDate),
        invoiceNo: invoiceNo,
        partyLedger: getStr(row.partyLedger),
        salesPurchaseLedger: getStr(row.salesPurchaseLedger),
        items: [],
        charges: [],
        gstMode: getStr(row.gstMode).toLowerCase() === 'manual' ? 'Manual' : 'Auto',
        gstRate: getNum(row.gstRate),
        hsn: getStr(row.hsn),
        
        freightLedger: getStr(row.freightLedger),
        freightAmount: getNum(row.freightAmount),
        packingLedger: getStr(row.packingLedger),
        packingAmount: getNum(row.packingAmount),
        loadingLedger: getStr(row.loadingLedger),
        loadingAmount: getNum(row.loadingAmount),
        insuranceLedger: getStr(row.insuranceLedger),
        insuranceAmount: getNum(row.insuranceAmount),
        otherLedger1: getStr(row.otherLedger1),
        otherAmount1: getNum(row.otherAmount1),
        otherLedger2: getStr(row.otherLedger2),
        otherAmount2: getNum(row.otherAmount2),
        discountLedger: getStr(row.discountLedger),
        discountAmount: getNum(row.discountAmount),
        roundOffLedger: getStr(row.roundOffLedger),
        roundOffAmount: getNum(row.roundOffAmount),

        partyGSTIN: getStr(row.partyGSTIN),
        partyAddress1: getStr(row.partyAddress1),
        partyAddress2: getStr(row.partyAddress2),
        partyState: getStr(row.partyState),
        placeOfSupply: getStr(row.placeOfSupply),
        partyRegistrationType: getStr(row.partyRegistrationType),

        dispatchDate: getStr(row.dispatchDate),
        deliveryNoteNo: getStr(row.deliveryNoteNo),
        dispatchDocNo: getStr(row.dispatchDocNo),
        biltyLRNo: getStr(row.biltyLRNo),
        transporterName: getStr(row.transporterName),
        transporterGSTIN: getStr(row.transporterGSTIN),
        vehicleNo: getStr(row.vehicleNo),
        destination: getStr(row.destination),
        modeOfTransport: getStr(row.modeOfTransport),
        ewayBillNo: getStr(row.ewayBillNo),

        narration: getStr(row.narration),
        reference: getStr(row.reference),
        errors: [],
        warnings: [],
        isValid: true,
        totalTaxableValue: 0,
        totalCGST: 0,
        totalSGST: 0,
        totalIGST: 0,
        totalAdditionalCharges: 0,
        invoiceTotal: 0
      };
      invoicesMap.set(key, inv);
    }

    const stockItem = getStr(row.stockItem);
    const quantity = getNum(row.quantity);
    const rate = getNum(row.rate);
    const itemAmount = getNum(row.itemAmount) || (quantity * rate);
    const taxableValue = getNum(row.taxableValue) || itemAmount;

    // We capture row details even if there is no stockItem name, as long as there is some financial data
    const hasRowData = stockItem || quantity > 0 || rate > 0 || itemAmount > 0 || taxableValue > 0 || getNum(row.cgstAmount) > 0 || getNum(row.sgstAmount) > 0 || getNum(row.igstAmount) > 0;

    if (hasRowData) {
      const item: SalesPurchaseItem = {
        stockItem: stockItem, // Might be empty for Accounting mode
        description: getStr(row.description),
        quantity: quantity,
        unit: getStr(row.unit || 'Nos'),
        rate: rate,
        itemAmount: itemAmount,
        discountPercent: getNum(row.discountPercent),
        discountAmount: getNum(row.discountAmount),
        taxableValue: taxableValue,
        hsn: getStr(row.hsn),
        gstRate: getNum(row.gstRate),
        cgstLedger: getStr(row.cgstLedger),
        cgstAmount: getNum(row.cgstAmount),
        sgstLedger: getStr(row.sgstLedger),
        sgstAmount: getNum(row.sgstAmount),
        igstLedger: getStr(row.igstLedger),
        igstAmount: getNum(row.igstAmount),
      };
      inv.items.push(item);
    }
  }

  const invoices = Array.from(invoicesMap.values());

  invoices.forEach(inv => {
    // Determine Voucher Mode
    if (inv.voucherMode === 'Auto' || !inv.voucherMode) {
      const hasPhysicalItems = inv.items.some(item => item.stockItem && item.stockItem.trim().length > 0);
      if (hasPhysicalItems) {
        inv.voucherMode = 'Item Invoice';
      } else {
        inv.voucherMode = 'Accounting';
      }
    }

    // Accumulate charge details
    const addCharge = (ledger: string, amount: number, isDiscount = false) => {
      if (ledger && amount !== 0) {
        inv.charges.push({ ledger, amount, isDiscount });
      }
    };

    addCharge(inv.freightLedger, inv.freightAmount || 0);
    addCharge(inv.packingLedger, inv.packingAmount || 0);
    addCharge(inv.loadingLedger, inv.loadingAmount || 0);
    addCharge(inv.insuranceLedger, inv.insuranceAmount || 0);
    addCharge(inv.otherLedger1, inv.otherAmount1 || 0);
    addCharge(inv.otherLedger2, inv.otherAmount2 || 0);
    addCharge(inv.discountLedger, inv.discountAmount || 0, true);
    addCharge(inv.roundOffLedger, inv.roundOffAmount || 0);
  });

  return invoices;
}
