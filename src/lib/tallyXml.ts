import { XMLBuilder } from 'fast-xml-parser';

export interface TallyVoucher {
  date: string;
  voucherType: string;
  partyName: string;
  bankLedger: string;
  voucherNumber?: string;
  narration?: string;
  reference?: string;
  ledgerEntries: {
    ledgerName: string;
    isDeemedPositive: 'Yes' | 'No';
    amount: number;
    isPartyLedger?: 'Yes' | 'No';
    isLastDeemedPositive?: 'Yes' | 'No';
  }[];
}

export function validateAndFormatDateParts(y: number, m: number, d: number): string {
  if (isNaN(y) || isNaN(m) || isNaN(d)) {
    throw new Error(`Invalid date parts: Year=${y}, Month=${m}, Day=${d}`);
  }
  let finalYear = y;
  if (finalYear < 100) {
    finalYear += 2000;
  }
  if (finalYear < 1900 || finalYear > 2100) {
    throw new Error(`Year ${finalYear} is out of realistic range (1900-2100)`);
  }
  if (m < 1 || m > 12) {
    throw new Error(`Month ${m} is out of range (1-12)`);
  }
  const isLeap = finalYear % 4 === 0 && (finalYear % 100 !== 0 || finalYear % 400 === 0);
  const daysInMonths = [31, isLeap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  if (d < 1 || d > daysInMonths[m - 1]) {
    throw new Error(`Day ${d} is out of range for month ${m} in year ${finalYear}`);
  }
  return `${finalYear}${String(m).padStart(2, '0')}${String(d).padStart(2, '0')}`;
}

export function normalizeTallyDate(input: any): { isValid: boolean; value: string; error?: string } {
  if (input === undefined || input === null) {
    return { isValid: false, value: '', error: 'Date is empty.' };
  }

  // If input is already in YYYYMMDD format (string, length 8, all digits)
  if (typeof input === 'string' && /^\d{8}$/.test(input)) {
    const y = parseInt(input.substring(0, 4));
    const m = parseInt(input.substring(4, 6));
    const d = parseInt(input.substring(6, 8));
    try {
      const val = validateAndFormatDateParts(y, m, d);
      return { isValid: true, value: val };
    } catch (e: any) {
      return { isValid: false, value: '', error: e.message };
    }
  }

  let y: number, m: number, d: number;

  try {
    if (input instanceof Date) {
      if (isNaN(input.getTime())) throw new Error("Invalid Date object");
      y = input.getFullYear();
      m = input.getMonth() + 1;
      d = input.getDate();
      return { isValid: true, value: validateAndFormatDateParts(y, m, d) };
    }

    if (typeof input === 'number') {
      if (input > 0 && input < 100000) {
        const date = new Date((input - 25569) * 86400 * 1000);
        if (isNaN(date.getTime())) throw new Error("Invalid numeric Excel date serial");
        y = date.getFullYear();
        m = date.getMonth() + 1;
        d = date.getDate();
        return { isValid: true, value: validateAndFormatDateParts(y, m, d) };
      }
      return { isValid: false, value: '', error: `Invalid date serial number: ${input}` };
    }

    const str = String(input).trim();
    if (!str) {
      return { isValid: false, value: '', error: 'Date is empty.' };
    }

    const monthsMap: Record<string, number> = {
      jan: 1, january: 1,
      feb: 2, february: 2,
      mar: 3, march: 3,
      apr: 4, april: 4,
      may: 5,
      jun: 6, june: 6,
      jul: 7, july: 7,
      aug: 8, august: 8,
      sep: 9, september: 9,
      oct: 10, october: 10,
      nov: 11, november: 11,
      dec: 12, december: 12
    };

    // Try parsing YYYY-MM-DD or YYYY/MM/DD
    const ymdMatch = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (ymdMatch) {
      const [_, yearStr, monthStr, dayStr] = ymdMatch;
      return { isValid: true, value: validateAndFormatDateParts(parseInt(yearStr), parseInt(monthStr), parseInt(dayStr)) };
    }

    // Try parsing DD-MM-YYYY or DD/MM/YYYY
    const dmyMatch = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
    if (dmyMatch) {
      const [_, dayStr, monthStr, yearStr] = dmyMatch;
      return { isValid: true, value: validateAndFormatDateParts(parseInt(yearStr), parseInt(monthStr), parseInt(dayStr)) };
    }

    // Try split by delimiters (month words like Apr, April)
    const parts = str.toLowerCase().split(/[-/\s,]+/);
    if (parts.length === 3) {
      let day = 0;
      let month = 0;
      let year = 0;

      if (monthsMap[parts[1]] !== undefined) {
        month = monthsMap[parts[1]];
        day = parseInt(parts[0]);
        year = parseInt(parts[2]);
      } else if (monthsMap[parts[0]] !== undefined) {
        month = monthsMap[parts[0]];
        day = parseInt(parts[1]);
        year = parseInt(parts[2]);
      } else if (monthsMap[parts[2]] !== undefined) {
        month = monthsMap[parts[2]];
        day = parseInt(parts[0]);
        year = parseInt(parts[1]);
      } else {
        // Direct numeric split fallback (Indian DD/MM/YYYY preferred)
        day = parseInt(parts[0]);
        month = parseInt(parts[1]);
        year = parseInt(parts[2]);
      }

      return { isValid: true, value: validateAndFormatDateParts(year, month, day) };
    }

    // Last-ditch parse with JS native Date
    const parsedTs = Date.parse(str);
    if (!isNaN(parsedTs)) {
      const date = new Date(parsedTs);
      return { isValid: true, value: validateAndFormatDateParts(date.getFullYear(), date.getMonth() + 1, date.getDate()) };
    }

    return { isValid: false, value: '', error: `Could not parse date format: "${input}"` };
  } catch (e: any) {
    return { isValid: false, value: '', error: e.message };
  }
}

export interface TallyLedgerMaster {
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
}

export interface TallyStockItemMaster {
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
}

export interface TallyStockGroupMaster {
  groupName: string;
  underGroup?: string;
}

export interface TallyUnitMaster {
  symbol: string;
  formalName?: string;
  uqc?: string;
  decimalPlaces?: string;
}

export function generateTallyXML(vouchers: TallyVoucher[]): string {
  const builder = new XMLBuilder({
    ignoreAttributes: false,
    format: true,
    suppressEmptyNode: true,
  });

  const tallyData = {
    ENVELOPE: {
      HEADER: {
        TALLYREQUEST: 'Import Data',
      },
      BODY: {
        IMPORTDATA: {
          REQUESTDESC: {
            REPORTNAME: 'Vouchers',
          },
          REQUESTDATA: {
            TALLYMESSAGE: vouchers.map(v => {
              // 1. Get reference if available
              const refVal = v.reference ? v.reference.trim() : '';
              
              // 2. Prepare narration and append reference if it doesn't already exist in narration
              let finalNarration = v.narration ? v.narration.trim() : '';
              if (refVal) {
                const refSnippet = `Ref: ${refVal}`;
                if (finalNarration) {
                   if (!finalNarration.includes(refVal)) {
                     finalNarration = `${finalNarration} | ${refSnippet}`;
                   }
                } else {
                  finalNarration = refSnippet;
                }
              }

              return {
                VOUCHER: {
                  '@_VCHTYPE': v.voucherType,
                  '@_ACTION': 'Create',
                  '@_OBJVIEW': 'Accounting Voucher View',
                  DATE: v.date.replace(/-/g, ''),
                  REFERENCEDATE: v.date.replace(/-/g, ''),
                  VCHSTATUSDATE: v.date.replace(/-/g, ''),
                  EFFECTIVEDATE: v.date.replace(/-/g, ''),
                  VOUCHERTYPENAME: v.voucherType,
                  VCHSTATUSVOUCHERTYPE: v.voucherType,
                  VOUCHERTYPEORIGNAME: v.voucherType,
                  PARTYLEDGERNAME: v.bankLedger, // Selected Bank/Cash Ledger as Party
                  VOUCHERNUMBER: v.voucherNumber || undefined,
                  NARRATION: finalNarration || undefined,
                  REFERENCE: refVal || undefined,
                  'ALLLEDGERENTRIES.LIST': v.ledgerEntries.map(le => ({
                    LEDGERNAME: le.ledgerName,
                    ISDEEMEDPOSITIVE: le.isDeemedPositive,
                    ISLASTDEEMEDPOSITIVE: le.isLastDeemedPositive || le.isDeemedPositive,
                    ISPARTYLEDGER: le.isPartyLedger || 'No',
                    AMOUNT: le.amount.toFixed(2), // Format with 2 decimal places with sign
                  })),
                },
              };
            }),
          },
        },
      },
    },
  };

  return builder.build(tallyData);
}

export function generateLedgersXML(ledgers: TallyLedgerMaster[]): string {
  const builder = new XMLBuilder({
    ignoreAttributes: false,
    format: true,
    suppressEmptyNode: true,
  });

  const tallyData = {
    ENVELOPE: {
      HEADER: {
        TALLYREQUEST: 'Import Data',
      },
      BODY: {
        IMPORTDATA: {
          REQUESTDESC: {
            REPORTNAME: 'All Masters',
          },
          REQUESTDATA: {
            TALLYMESSAGE: ledgers.map(l => {
              const openingBal = parseFloat(l.openingBalance || '0');
              const sign = l.drCr === 'Cr' ? '-' : '';
              const openingBalanceStr = openingBal > 0 ? `${sign}${openingBal}` : undefined;

              const addresses: string[] = [];
              if (l.address1) addresses.push(l.address1);
              if (l.address2) addresses.push(l.address2);

              return {
                LEDGER: {
                  '@_NAME': l.ledgerName,
                  '@_ACTION': 'Create',
                  'NAME.LIST': {
                    NAME: l.ledgerName,
                  },
                  PARENT: l.underGroup,
                  COUNTRYNAME: l.country || 'INDIA',
                  GSTREGISTRATIONTYPE: l.registrationType || undefined,
                  PARTYGSTIN: l.gstin || undefined,
                  LEDSTATENAME: l.state || undefined,
                  PINCODE: l.pincode || undefined,
                  INCOMETAXNUMBER: l.pan || undefined,
                  OPENINGBALANCE: openingBalanceStr,
                  ISBILLWISEON: l.isBillwiseOn || 'No',
                  ISCOSTCENTRESON: l.isCostCentreOn || 'No',
                  EMAIL: l.email || undefined,
                  MAILINGNAME: l.mailingName || l.ledgerName,
                  ...(addresses.length > 0 ? {
                    'ADDRESS.LIST': {
                      ADDRESS: addresses
                    }
                  } : {})
                }
              };
            })
          }
        }
      }
    }
  };

  return builder.build(tallyData);
}

export function generateStockItemsXML(items: TallyStockItemMaster[]): string {
  const builder = new XMLBuilder({
    ignoreAttributes: false,
    format: true,
    suppressEmptyNode: true,
  });

  const tallyData = {
    ENVELOPE: {
      HEADER: {
        TALLYREQUEST: 'Import Data',
      },
      BODY: {
        IMPORTDATA: {
          REQUESTDESC: {
            REPORTNAME: 'All Masters',
          },
          REQUESTDATA: {
            TALLYMESSAGE: items.map(item => {
              const qty = parseFloat(item.openingQty || '0');
              const rate = parseFloat(item.openingRate || '0');
              const val = parseFloat(item.openingValue || '0');

              return {
                STOCKITEM: {
                  '@_NAME': item.itemName,
                  '@_ACTION': 'Create',
                  'NAME.LIST': {
                    NAME: item.itemName,
                  },
                  PARENT: item.underGroup,
                  BASEUNITS: item.unit,
                  OPENINGBALANCE: qty > 0 ? `${qty} ${item.unit}` : undefined,
                  OPENINGRATE: rate > 0 ? `${rate}/${item.unit}` : undefined,
                  OPENINGVALUE: val > 0 ? `-${val}` : undefined,
                  DESCRIPTION: item.description || undefined,
                  HSNCODE: item.hsn || undefined,
                  GSTAPPLICABLE: item.gstApplicable || undefined,
                  TAXABILITY: item.taxability || undefined,
                  ...(item.gstRate ? {
                    'GSTDETAILS.LIST': {
                      GSTTAXABILITY: item.taxability || 'Taxable',
                      'STATEGSTDETAILS.LIST': {
                        'RATEDETAILS.LIST': [
                          {
                            GSTRATEDUTYHEAD: 'Integrated Tax',
                            GSTRATE: parseFloat(item.igstRate || item.gstRate || '0')
                          },
                          ...(item.cgstRate ? [{
                            GSTRATEDUTYHEAD: 'Central Tax',
                            GSTRATE: parseFloat(item.cgstRate)
                          }] : []),
                          ...(item.sgstRate ? [{
                            GSTRATEDUTYHEAD: 'State Tax',
                            GSTRATE: parseFloat(item.sgstRate)
                          }] : [])
                        ]
                      }
                    }
                  } : {})
                }
              };
            })
          }
        }
      }
    }
  };

  return builder.build(tallyData);
}

export function generateStockGroupsXML(groups: TallyStockGroupMaster[]): string {
  const builder = new XMLBuilder({
    ignoreAttributes: false,
    format: true,
    suppressEmptyNode: true,
  });

  const tallyData = {
    ENVELOPE: {
      HEADER: {
        TALLYREQUEST: 'Import Data',
      },
      BODY: {
        IMPORTDATA: {
          REQUESTDESC: {
            REPORTNAME: 'All Masters',
          },
          REQUESTDATA: {
            TALLYMESSAGE: groups.map(g => ({
              STOCKGROUP: {
                '@_NAME': g.groupName,
                '@_ACTION': 'Create',
                'NAME.LIST': {
                  NAME: g.groupName,
                },
                PARENT: g.underGroup || 'Primary'
              }
            }))
          }
        }
      }
    }
  };

  return builder.build(tallyData);
}

export function generateUnitsXML(units: TallyUnitMaster[]): string {
  const builder = new XMLBuilder({
    ignoreAttributes: false,
    format: true,
    suppressEmptyNode: true,
  });

  const tallyData = {
    ENVELOPE: {
      HEADER: {
        TALLYREQUEST: 'Import Data',
      },
      BODY: {
        IMPORTDATA: {
          REQUESTDESC: {
            REPORTNAME: 'All Masters',
          },
          REQUESTDATA: {
            TALLYMESSAGE: units.map(u => ({
              UNIT: {
                '@_NAME': u.symbol,
                '@_ACTION': 'Create',
                NAME: u.symbol,
                ISSIMPLEUNIT: 'Yes',
                DECIMALPLACES: parseInt(u.decimalPlaces || '0', 10),
                ORIGINALNAME: u.formalName || undefined,
                REPORTINGUOM: u.uqc || undefined
              }
            }))
          }
        }
      }
    }
  };

  return builder.build(tallyData);
}

export interface SalesPurchaseInvoice {
  errors: string[];
  warnings: string[];
  isValid: boolean;
  voucherType: 'Sales' | 'Purchase';
  invoiceDate: string;
  invoiceNo: string;
  partyLedger: string;
  salesPurchaseLedger: string;
  partyGSTIN?: string;
  partyAddress1?: string;
  partyAddress2?: string;
  partyState?: string;
  placeOfSupply?: string;
  partyRegistrationType?: string;
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
  narration?: string;
  reference?: string;
  items: {
    stockItem: string;
    description?: string;
    quantity: number;
    unit: string;
    rate: number;
    itemAmount: number;
    discountPercent?: number;
    discountAmount?: number;
    taxableValue: number;
    hsn?: string;
    gstRate: number;
    cgstLedger?: string;
    cgstAmount?: number;
    sgstLedger?: string;
    sgstAmount?: number;
    igstLedger?: string;
    igstAmount?: number;
    gstRateSource?: string;
  }[];
  gstMode: 'Auto' | 'Manual';
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
  billDiscountAmount?: number;
  roundOffLedger?: string;
  roundOffAmount?: number;
  totalTaxableValue: number;
  totalCGST: number;
  totalSGST: number;
  totalIGST: number;
  totalAdditionalCharges: number;
  invoiceTotal: number;
}

export function generateSalesPurchaseXML(invoices: SalesPurchaseInvoice[], companyState?: string): string {
  const builder = new XMLBuilder({
    ignoreAttributes: false,
    format: true,
    suppressEmptyNode: true,
  });

  const tallyData = {
    ENVELOPE: {
      HEADER: {
        TALLYREQUEST: 'Import Data',
      },
      BODY: {
        IMPORTDATA: {
          REQUESTDESC: {
            REPORTNAME: 'Vouchers',
          },
          REQUESTDATA: {
            TALLYMESSAGE: invoices.map(inv => {
              const isSales = inv.voucherType === 'Sales';

              // Normalize and validate dates inside generator
              const dateNorm = normalizeTallyDate(inv.invoiceDate);
              if (!dateNorm.isValid) {
                throw new Error(`Invoice No "${inv.invoiceNo}": Invalid Invoice Date (${inv.invoiceDate}). Error: ${dateNorm.error}`);
              }
              const dateStr = dateNorm.value;

              let dispatchDateStr = '';
              if (inv.dispatchDate) {
                const dispNorm = normalizeTallyDate(inv.dispatchDate);
                if (!dispNorm.isValid) {
                  throw new Error(`Invoice No "${inv.invoiceNo}": Invalid Dispatch Date (${inv.dispatchDate}). Error: ${dispNorm.error}`);
                }
                dispatchDateStr = dispNorm.value;
              }
              
              const refVal = inv.reference ? inv.reference.trim() : '';
              let finalNarration = inv.narration ? inv.narration.trim() : '';
              if (refVal) {
                const refSnippet = `Ref: ${refVal}`;
                if (finalNarration) {
                  if (!finalNarration.includes(refVal)) {
                    finalNarration = `${finalNarration} | ${refSnippet}`;
                  }
                } else {
                  finalNarration = refSnippet;
                }
              }

              // Append E-Way Bill Number to narration safely
              if (inv.ewayBillNo) {
                const ewaySnippet = `E-Way Bill: ${inv.ewayBillNo.trim()}`;
                if (finalNarration) {
                  if (!finalNarration.includes(inv.ewayBillNo)) {
                    finalNarration = `${finalNarration} | ${ewaySnippet}`;
                  }
                } else {
                  finalNarration = ewaySnippet;
                }
              }

              // Address lines
              const addresses: string[] = [];
              if (inv.partyAddress1) addresses.push(inv.partyAddress1);
              if (inv.partyAddress2) addresses.push(inv.partyAddress2);

              // 1. Party ledger entry (Debited for Sales, Credited for Purchase)
              const partyAmount = isSales ? -inv.invoiceTotal : inv.invoiceTotal;
              const partyDeemedPositive = isSales ? 'Yes' : 'No';

              const ledgerEntriesList: any[] = [
                {
                  LEDGERNAME: inv.partyLedger,
                  ISDEEMEDPOSITIVE: partyDeemedPositive,
                  ISLASTDEEMEDPOSITIVE: partyDeemedPositive,
                  ISPARTYLEDGER: 'Yes',
                  AMOUNT: partyAmount.toFixed(2),
                }
              ];

              // 2. GST ledger entries
              const gstLedgerSums: Record<string, number> = {};
              inv.items.forEach(item => {
                if (item.cgstLedger && item.cgstAmount) {
                  gstLedgerSums[item.cgstLedger] = (gstLedgerSums[item.cgstLedger] || 0) + item.cgstAmount;
                }
                if (item.sgstLedger && item.sgstAmount) {
                  gstLedgerSums[item.sgstLedger] = (gstLedgerSums[item.sgstLedger] || 0) + item.sgstAmount;
                }
                if (item.igstLedger && item.igstAmount) {
                  gstLedgerSums[item.igstLedger] = (gstLedgerSums[item.igstLedger] || 0) + item.igstAmount;
                }
              });

              Object.entries(gstLedgerSums).forEach(([ledgerName, amount]) => {
                if (amount > 0) {
                  const gstAmountSign = isSales ? amount : -amount;
                  const gstDeemedPositive = isSales ? 'No' : 'Yes';
                  ledgerEntriesList.push({
                    LEDGERNAME: ledgerName,
                    ISDEEMEDPOSITIVE: gstDeemedPositive,
                    ISLASTDEEMEDPOSITIVE: gstDeemedPositive,
                    ISPARTYLEDGER: 'No',
                    AMOUNT: gstAmountSign.toFixed(2),
                  });
                }
              });

              // 3. Additional charges
              const additionalCharges = [
                { ledger: inv.freightLedger, amount: inv.freightAmount, isDiscount: false },
                { ledger: inv.packingLedger, amount: inv.packingAmount, isDiscount: false },
                { ledger: inv.loadingLedger, amount: inv.loadingAmount, isDiscount: false },
                { ledger: inv.insuranceLedger, amount: inv.insuranceAmount, isDiscount: false },
                { ledger: inv.otherLedger1, amount: inv.otherAmount1, isDiscount: false },
                { ledger: inv.otherLedger2, amount: inv.otherAmount2, isDiscount: false },
                { ledger: inv.discountLedger, amount: inv.billDiscountAmount, isDiscount: true },
                { ledger: inv.roundOffLedger, amount: inv.roundOffAmount, isDiscount: false, isRoundOff: true }
              ];

              additionalCharges.forEach(ac => {
                if (ac.ledger && ac.amount !== undefined && ac.amount !== 0) {
                  let signValue = ac.amount;
                  if (ac.isDiscount) {
                     signValue = -ac.amount;
                  }

                  let chargeAmountSign = isSales ? signValue : -signValue;
                  const chargeDeemedPositive = chargeAmountSign < 0 ? 'Yes' : 'No';

                  ledgerEntriesList.push({
                    LEDGERNAME: ac.ledger,
                    ISDEEMEDPOSITIVE: chargeDeemedPositive,
                    ISLASTDEEMEDPOSITIVE: chargeDeemedPositive,
                    ISPARTYLEDGER: 'No',
                    AMOUNT: chargeAmountSign.toFixed(2),
                  });
                }
              });

              // 4. Stock items list
              const inventoryEntriesList = inv.items.map(item => {
                const itemAmountVal = item.taxableValue;
                const inventoryAmountSign = isSales ? itemAmountVal : -itemAmountVal;
                const invDeemedPositive = isSales ? 'No' : 'Yes';

                // Inter-state or Intra-state check for RATEDETAILS.LIST
                const isInterState = inv.placeOfSupply && companyState && (inv.placeOfSupply.toLowerCase().trim() !== companyState.toLowerCase().trim());
                
                const rateDetailsList = [];
                rateDetailsList.push({
                  GSTRATEDUTYHEAD: 'CGST',
                  GSTRATEVALUATIONTYPE: 'Based on Value',
                  GSTRATE: isInterState ? 0 : (item.gstRate / 2)
                });
                rateDetailsList.push({
                  GSTRATEDUTYHEAD: 'SGST/UTGST',
                  GSTRATEVALUATIONTYPE: 'Based on Value',
                  GSTRATE: isInterState ? 0 : (item.gstRate / 2)
                });
                rateDetailsList.push({
                  GSTRATEDUTYHEAD: 'IGST',
                  GSTRATEVALUATIONTYPE: 'Based on Value',
                  GSTRATE: isInterState ? item.gstRate : 0
                });
                rateDetailsList.push({
                  GSTRATEDUTYHEAD: 'Cess',
                  GSTRATEVALUATIONTYPE: 'Not Applicable'
                });
                rateDetailsList.push({
                  GSTRATEDUTYHEAD: 'State Cess',
                  GSTRATEVALUATIONTYPE: 'Based on Value'
                });

                const hsnVal = (item.hsn || '').trim();
                if (!hsnVal) {
                  throw new Error(`Invoice No "${inv.invoiceNo}": Item "${item.stockItem}" is missing a valid HSN/SAC.`);
                }

                return {
                  STOCKITEMNAME: item.stockItem,
                  DESCRIPTION: item.description || undefined,
                  GSTOVRDNTAXABILITY: 'Taxable',
                  GSTSOURCETYPE: 'Stock Item',
                  GSTITEMSOURCE: item.stockItem,
                  HSNSOURCETYPE: 'Stock Item',
                  HSNITEMSOURCE: item.stockItem,
                  GSTOVRDNTYPEOFSUPPLY: 'Goods',
                  GSTRATEINFERAPPLICABILITY: 'As per Masters/Company',
                  GSTHSNNAME: hsnVal,
                  GSTHSNDESCRIPTION: item.description || undefined,
                  GSTHSNINFERAPPLICABILITY: 'As per Masters/Company',
                  ISDEEMEDPOSITIVE: invDeemedPositive,
                  ISLASTDEEMEDPOSITIVE: invDeemedPositive,
                  RATE: `${item.rate.toFixed(2)}/${item.unit}`,
                  AMOUNT: inventoryAmountSign.toFixed(2),
                  ACTUALQTY: `${item.quantity} ${item.unit}`,
                  BILLEDQTY: `${item.quantity} ${item.unit}`,
                  'BATCHALLOCATIONS.LIST': [
                    {
                      GODOWNNAME: 'Main Location',
                      BATCHNAME: 'Primary Batch',
                      AMOUNT: inventoryAmountSign.toFixed(2),
                      ACTUALQTY: `${item.quantity} ${item.unit}`,
                      BILLEDQTY: `${item.quantity} ${item.unit}`,
                    }
                  ],
                  'ACCOUNTINGALLOCATIONS.LIST': [
                    {
                      LEDGERNAME: inv.salesPurchaseLedger,
                      ISDEEMEDPOSITIVE: invDeemedPositive,
                      ISLASTDEEMEDPOSITIVE: invDeemedPositive,
                      AMOUNT: inventoryAmountSign.toFixed(2),
                    }
                  ],
                  'RATEDETAILS.LIST': rateDetailsList
                };
              });

              const registrationTypeVal = inv.partyRegistrationType || 'Regular';

              return {
                VOUCHER: {
                  '@_VCHTYPE': inv.voucherType,
                  '@_ACTION': 'Create',
                  '@_OBJVIEW': 'Invoice Voucher View',
                  DATE: dateStr,
                  REFERENCEDATE: dateStr,
                  VCHSTATUSDATE: dateStr,
                  EFFECTIVEDATE: dateStr,
                  GSTREGISTRATIONTYPE: registrationTypeVal,
                  VATDEALERTYPE: registrationTypeVal,
                  STATENAME: inv.partyState || undefined,
                  COUNTRYOFRESIDENCE: 'India',
                  PARTYGSTIN: inv.partyGSTIN || undefined,
                  PLACEOFSUPPLY: inv.placeOfSupply || undefined,
                  VOUCHERTYPENAME: inv.voucherType,
                  PARTYNAME: inv.partyLedger,
                  PARTYLEDGERNAME: inv.partyLedger,
                  VOUCHERNUMBER: inv.invoiceNo,
                  BASICBUYERNAME: inv.partyLedger,
                  PARTYMAILINGNAME: inv.partyLedger,
                  CONSIGNEEGSTIN: inv.partyGSTIN || undefined,
                  CONSIGNEEMAILINGNAME: inv.partyLedger,
                  CONSIGNEESTATENAME: inv.partyState || undefined,
                  CMPGSTSTATE: companyState || undefined,
                  CONSIGNEECOUNTRYNAME: 'India',
                  BASICBASEPARTYNAME: inv.partyLedger,
                  PERSISTEDVIEW: 'Invoice Voucher View',
                  VCHSTATUSVOUCHERTYPE: inv.voucherType,
                  VCHENTRYMODE: 'Item Invoice',
                  ISINVOICE: 'Yes',
                  HASCASHFLOW: 'No',
                  
                  ...(addresses.length > 0 ? {
                    'ADDRESS.LIST': {
                      ADDRESS: addresses
                    },
                    'BASICBUYERADDRESS.LIST': {
                      BASICBUYERADDRESS: addresses
                    }
                  } : {}),
                  
                  BILLOFLADINGDATE: dispatchDateStr || undefined,
                  BASICSHIPDELIVERYNOTE: inv.deliveryNoteNo || undefined,
                  BASICSHIPDOCUMENTNO: inv.dispatchDocNo || undefined,
                  BILLOFLADINGNO: inv.biltyLRNo || undefined,
                  BASICSHIPTRANSPORTNAME: inv.transporterName || undefined,
                  TRANSPORTERGSTIN: inv.transporterGSTIN || undefined,
                  BASICSHIPVESSELNO: inv.vehicleNo || undefined,
                  BASICFINALDESTINATION: inv.destination || undefined,
                  BASICMODEOFTRANSPORT: inv.modeOfTransport || undefined,
                  NARRATION: finalNarration || undefined,

                  'ALLINVENTORYENTRIES.LIST': inventoryEntriesList,
                  'LEDGERENTRIES.LIST': ledgerEntriesList,
                }
              };
            })
          }
        }
      }
    }
  };

  return builder.build(tallyData);
}
