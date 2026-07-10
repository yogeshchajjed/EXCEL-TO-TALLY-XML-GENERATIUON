import { XMLBuilder } from 'fast-xml-parser';

export interface TallyVoucher {
  date: string;
  voucherType: string;
  partyName: string;
  voucherNumber?: string;
  narration?: string;
  reference?: string;
  ledgerEntries: {
    ledgerName: string;
    isDeemedPositive: 'Yes' | 'No';
    amount: number;
  }[];
}

export function normalizeTallyDate(input: any): { isValid: boolean; value: string; error?: string } {
  if (input === undefined || input === null) {
    return { isValid: false, value: '', error: 'Date is empty.' };
  }

  let dateObj: Date | null = null;

  if (input instanceof Date) {
    dateObj = input;
  } else if (typeof input === 'number') {
    // If it's a number, check if it's a valid Excel serial date
    if (input > 0 && input < 100000) {
      dateObj = new Date((input - 25569) * 86400 * 1000);
    } else {
      return { isValid: false, value: '', error: `Invalid date serial number: ${input}` };
    }
  } else {
    const trimmed = String(input).trim();
    if (!trimmed) {
      return { isValid: false, value: '', error: 'Date is empty.' };
    }

    // Try parsing YYYY-MM-DD or YYYY/MM/DD
    const ymdRegex = /^(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})/;
    const ymdMatch = trimmed.match(ymdRegex);
    if (ymdMatch) {
      const year = parseInt(ymdMatch[1], 10);
      const month = parseInt(ymdMatch[2], 10) - 1;
      const day = parseInt(ymdMatch[3], 10);
      dateObj = new Date(year, month, day);
    } else {
      // Try parsing DD/MM/YYYY or DD-MM-YYYY (Ambiguous treated as Indian format DD/MM/YYYY)
      const dmyRegex = /^(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})/;
      const dmyMatch = trimmed.match(dmyRegex);
      if (dmyMatch) {
        const day = parseInt(dmyMatch[1], 10);
        const month = parseInt(dmyMatch[2], 10) - 1;
        const year = parseInt(dmyMatch[3], 10);
        dateObj = new Date(year, month, day);
      } else {
        // Fallback to JS native parser if it's another string format
        const parsed = Date.parse(trimmed);
        if (!isNaN(parsed)) {
          dateObj = new Date(parsed);
        }
      }
    }
  }

  if (!dateObj || isNaN(dateObj.getTime())) {
    return { isValid: false, value: '', error: `Could not parse date format: "${input}"` };
  }

  const y = dateObj.getFullYear();
  const m = dateObj.getMonth() + 1;
  const d = dateObj.getDate();

  // Validate range sanity
  if (y < 1990 || y > 2100) {
    return { isValid: false, value: '', error: `Year out of range: ${y}` };
  }

  const formattedValue = `${y}${String(m).padStart(2, '0')}${String(d).padStart(2, '0')}`;
  return { isValid: true, value: formattedValue };
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
                  DATE: v.date.replace(/-/g, ''),
                  VOUCHERTYPENAME: v.voucherType,
                  PARTYNAME: v.partyName,
                  VOUCHERNUMBER: v.voucherNumber || undefined,
                  NARRATION: finalNarration || undefined,
                  REFERENCE: refVal || undefined,
                  'ALLLEDGERENTRIES.LIST': v.ledgerEntries.map(le => ({
                    LEDGERNAME: le.ledgerName,
                    ISDEEMEDPOSITIVE: le.isDeemedPositive,
                    AMOUNT: le.amount,
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
