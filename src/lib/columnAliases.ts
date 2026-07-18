export const COLUMN_ALIASES: Record<string, string[]> = {
  voucherType: ['voucher type', 'vouchertype', 'vchtype', 'type'],
  invoiceDate: ['invoice date', 'invoicedate', 'date', 'invdate', 'vchdate'],
  invoiceNo: ['invoice no', 'invoiceno', 'invno', 'vchno', 'voucher number', 'voucherno'],
  partyLedger: ['party ledger', 'partyledger', 'party name', 'partyname', 'customer', 'supplier'],
  salesPurchaseLedger: ['sales/purchase ledger', 'sales purchase ledger', 'sales ledger', 'purchase ledger', 'ledger name', 'sales/purchase a/c', 'sales a/c', 'purchase a/c'],
  voucherMode: ['voucher mode', 'vouchermode', 'mode'],
  inventoryMode: ['inventory mode', 'inventorymode'],
  stockItem: ['stock item', 'stockitem', 'item name', 'itemname', 'product', 'particulars'],
  quantity: ['quantity', 'qty', 'units'],
  unit: ['unit', 'uom', 'symbol', 'units symbol'],
  rate: ['rate', 'price', 'unit price'],
  itemAmount: ['item amount', 'itemamount', 'amount', 'value'],
  taxableValue: ['taxable value (purchase value)', 'taxable value', 'taxablevalue', 'assessable value', 'assessablevalue'],
  discountPercent: ['discount %', 'discountpercent', 'discount percent', 'disc %'],
  description: ['description', 'item description', 'desc'],
  gstMode: ['gst mode', 'gstmode'],
  gstRate: ['gst rate %', 'gst rate', 'gstrate', 'tax rate', 'taxrate'],
  hsn: ['hsn/sac', 'hsnsac', 'hsn', 'sac'],
  cgstLedger: ['cgst ledger', 'cgstledger', 'central tax ledger'],
  cgstAmount: ['cgst amount', 'cgstamount', 'cgst'],
  sgstLedger: ['sgst ledger', 'sgstledger', 'state tax ledger'],
  sgstAmount: ['sgst amount', 'sgstamount', 'sgst'],
  igstLedger: ['igst ledger', 'igstledger', 'integrated tax ledger'],
  igstAmount: ['igst amount', 'igstamount', 'igst'],
  freightLedger: ['freight ledger', 'freightledger'],
  freightAmount: ['freight amount', 'freightamount'],
  packingLedger: ['packing ledger', 'packingledger'],
  packingAmount: ['packing amount', 'packingamount'],
  loadingLedger: ['loading/unloading ledger', 'loading ledger', 'unloading ledger', 'loading/unloading amount', 'loading amount', 'loading/unloading ledger name'],
  loadingAmount: ['loading/unloading amount', 'loading amount', 'unloading amount'],
  insuranceLedger: ['insurance ledger', 'insuranceledger'],
  insuranceAmount: ['insurance amount', 'insuranceamount'],
  otherLedger1: ['other charges ledger 1', 'other ledger 1', 'otherledger1'],
  otherAmount1: ['other charges amount 1', 'other amount 1', 'otheramount1'],
  otherLedger2: ['other charges ledger 2', 'other ledger 2', 'otherledger2'],
  otherAmount2: ['other charges amount 2', 'other amount 2', 'otheramount2'],
  discountLedger: ['discount ledger', 'discountledger'],
  discountAmount: ['bill discount amount', 'discount amount', 'discount'],
  roundOffLedger: ['round off ledger', 'roundoffledger'],
  roundOffAmount: ['round off amount', 'roundoffamount', 'round off', 'roundoff'],
  partyGSTIN: ['party gstin', 'partygstin', 'gstin', 'gst number', 'gstno'],
  partyAddress1: ['party address 1', 'partyaddress1', 'address 1', 'address1'],
  partyAddress2: ['party address 2', 'partyaddress2', 'address 2', 'address2'],
  partyState: ['party state', 'partystate', 'state'],
  placeOfSupply: ['place of supply', 'placeofsupply', 'pos'],
  partyRegistrationType: ['party registration type', 'partyregistrationtype', 'registration type'],
  dispatchDate: ['dispatch date', 'dispatchdate'],
  deliveryNoteNo: ['delivery note no', 'deliverynoteno', 'delivery note', 'deliverynote'],
  dispatchDocNo: ['dispatch doc no', 'dispatchdocno', 'dispatch doc', 'dispatchdoc'],
  biltyLRNo: ['bilty/lr no', 'bilty lr no', 'biltyno', 'lrno', 'lr number', 'bilty number'],
  transporterName: ['transporter name', 'transportername', 'transporter'],
  transporterGSTIN: ['transporter gstin', 'transportergstin'],
  vehicleNo: ['vehicle no', 'vehicleno', 'vehicle number', 'vehiclenumber'],
  destination: ['destination', 'ship to', 'shipto'],
  modeOfTransport: ['mode of transport', 'modeoftransport', 'transport mode', 'transportmode'],
  ewayBillNo: ['e-way bill no', 'eway bill no', 'ewaybillno', 'eway bill', 'ewaybill'],
  narration: ['narration', 'nar', 'remarks', 'remark'],
  reference: ['reference', 'ref', 'ref no', 'refno']
};

export function getNormalizedField(header: string): string | null {
  const cleanHeader = header.toLowerCase().trim().replace(/[^a-z0-9/% _-]/g, '');
  for (const [field, aliases] of Object.entries(COLUMN_ALIASES)) {
    if (aliases.some(alias => {
      const cleanAlias = alias.toLowerCase().trim();
      return cleanHeader === cleanAlias || cleanHeader.includes(cleanAlias) || cleanAlias.includes(cleanHeader);
    })) {
      return field;
    }
  }
  return null;
}
