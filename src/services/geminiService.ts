import { GoogleGenAI, Type, ThinkingLevel } from "@google/genai";

let aiInstance: any = null;

function getAI(): any {
  if (!aiInstance) {
    const key = typeof process !== 'undefined' && process.env ? process.env.GEMINI_API_KEY : '';
    if (!key) {
      throw new Error('An API Key must be set when running in a browser');
    }
    aiInstance = new GoogleGenAI({ apiKey: key });
  }
  return aiInstance;
}

export interface ColumnMapping {
  excelColumn: string;
  tallyField: string;
  confidence: number;
  reasoning: string;
}

export function isElectron(): boolean {
  return typeof window !== 'undefined' && (window as any).electron?.isElectron === true;
}

export function hasApiKey(): boolean {
  return !!(typeof process !== 'undefined' && process.env && process.env.GEMINI_API_KEY);
}

export function isGeminiAvailable(): boolean {
  if (isElectron()) return false;
  return hasApiKey();
}

export function getLocalColumnMapping(headers: string[]): ColumnMapping[] {
  const standardFields = {
    DATE: ['date', 'txn date', 'transaction date', 'value date', 'posting date', 'txn_date'],
    PARTYNAME: ['party name', 'partyname', 'particulars', 'ledger', 'ledger name', 'account'],
    VOUCHERTYPENAME: ['voucher type', 'vouchertype', 'type', 'vch type', 'vouchertype name'],
    VOUCHERNUMBER: ['voucher number', 'vouchernumber', 'vch no', 'voucher no', 'no', 'vch_no'],
    AMOUNT: ['amount', 'transaction amount', 'credit', 'debit', 'value', 'total'],
    NARRATION: ['narration', 'description', 'particulars', 'transaction details', 'details', 'remarks', 'narr'],
    REFERENCE: ['reference', 'ref no', 'utr', 'cheque no', 'instrument no', 'transaction id', 'ref', 'chq no']
  };

  const mappings: ColumnMapping[] = [];

  for (const [field, aliases] of Object.entries(standardFields)) {
    // Try exact match first
    let foundHeader = headers.find(h => {
      const lower = h.toLowerCase().trim();
      return aliases.some(alias => lower === alias);
    });

    // If not found, try partial match (includes or startsWith)
    if (!foundHeader) {
      foundHeader = headers.find(h => {
        const lower = h.toLowerCase().trim();
        return aliases.some(alias => lower.includes(alias) || alias.includes(lower));
      });
    }

    if (foundHeader) {
      mappings.push({
        excelColumn: foundHeader,
        tallyField: field,
        confidence: 0.95,
        reasoning: `Deterministic match for ${field} using local mapping engine.`
      });
    }
  }

  return mappings;
}

export async function getAIColumnMapping(headers: string[]): Promise<ColumnMapping[]> {
  if (!isGeminiAvailable()) {
    return getLocalColumnMapping(headers);
  }

  const prompt = `
    Analyze the following Excel column headers and map them to standard Tally XML fields.
    Standard Tally fields include: DATE, PARTYNAME, VOUCHERTYPENAME, VOUCHERNUMBER, AMOUNT, NARRATION, REFERENCE.
    
    CRITICAL RULES:
    1. DATE: Map the column containing transaction dates to DATE.
    2. NARRATION: Map the column containing transaction details/particulars/narration to NARRATION. This is the MOST IMPORTANT column for accounting logic.
    3. DISTINGUISH DATE vs NARRATION: Do NOT map a column that contains dates (like "Date", "Txn Date", "Value Date") to NARRATION.
    4. NARRATION IDENTIFICATION: The "NARRATION" field is often called "Description", "Particulars", "Narration", "Transaction Details", "Remarks", "Comments", or "Details". Map the most descriptive text column to NARRATION.
    5. UNIQUE MAPPING: Each Tally field should ideally be mapped to a unique Excel column.
    6. CASE INSENSITIVITY: Treat headers as case-insensitive.
    
    Headers: ${headers.join(', ')}
    
    Return a JSON array of mappings.
  `;

  const response = await getAI().models.generateContent({
    model: "gemini-3.1-pro-preview",
    contents: prompt,
    config: {
      thinkingConfig: {
        thinkingLevel: ThinkingLevel.HIGH
      },
      responseMimeType: "application/json",
      responseSchema: {
        type: Type.ARRAY,
        items: {
          type: Type.OBJECT,
          properties: {
            excelColumn: { type: Type.STRING },
            tallyField: { type: Type.STRING },
            confidence: { type: Type.NUMBER },
            reasoning: { type: Type.STRING }
          },
          required: ["excelColumn", "tallyField", "confidence", "reasoning"]
        }
      }
    }
  });

  return JSON.parse(response.text);
}

export interface BankTransaction {
  date: string;
  description: string;
  reference?: string;
  debit?: number;
  credit?: number;
  amount: number;
  particulars?: string;
}

export interface MappedTransaction extends BankTransaction {
  tallyLedger: string;
  confidence: number;
  reasoning: string;
}

export async function parseBankStatementText(text: string): Promise<BankTransaction[]> {
  if (!isGeminiAvailable()) {
    throw new Error("Gemini is not available. Please use local bank statement parser.");
  }

  const prompt = `
    Extract bank transactions from the following text extracted from a PDF bank statement.
    Return a JSON array of objects with the following keys: date, description, amount.
    The "description" field should contain the full transaction narration/particulars.
    Ensure amounts are numbers. If a transaction has both debit and credit, use a single amount field (positive for credit/deposit, negative for debit/withdrawal if that's the convention, or just provide the absolute value and we will map it later).
    
    Text: ${text}
  `;

  const response = await getAI().models.generateContent({
    model: "gemini-3-flash-preview",
    contents: prompt,
    config: {
      responseMimeType: "application/json",
      responseSchema: {
        type: Type.ARRAY,
        items: {
          type: Type.OBJECT,
          properties: {
            date: { type: Type.STRING },
            description: { type: Type.STRING },
            amount: { type: Type.NUMBER }
          },
          required: ["date", "description", "amount"]
        }
      }
    }
  });

  return JSON.parse(response.text);
}

export async function mapBankTransactions(
  transactions: BankTransaction[],
  tallyLedgers: string[],
  historicalMappings?: { narration: string; ledger: string }[]
): Promise<MappedTransaction[]> {
  const hasSuspense = tallyLedgers.includes('Suspense Account');

  // We should process transactions and apply the Rules!
  // Rule I: If Particulars / Ledger field is already filled in standard template, preserve it!
  const results = transactions.map(tx => {
    if (tx.particulars && tx.particulars.trim() !== '') {
      return {
        ...tx,
        tallyLedger: tx.particulars.trim(),
        confidence: 1.0,
        reasoning: 'Preserved user-provided ledger from Excel file.'
      };
    }

    // Otherwise, try to suggest ledger using local deterministic keyword / historical mappings logic!
    const cleanNarr = tx.description.toLowerCase().trim();
    let matchedLedger = '';
    let confidence = 0;
    let reasoning = '';

    // 1. Exact historical
    if (historicalMappings) {
      const exact = historicalMappings.find(h => h.narration.toLowerCase().trim() === cleanNarr);
      if (exact && tallyLedgers.includes(exact.ledger)) {
        matchedLedger = exact.ledger;
        confidence = 1.0;
        reasoning = 'Matched exact historical narration from Daybook';
      }
    }

    // 2. Keyword match
    if (!matchedLedger) {
      for (const ledger of tallyLedgers) {
        const ledgerLower = ledger.toLowerCase().trim();
        if (cleanNarr.includes(ledgerLower) && ledgerLower.length > 3) {
          matchedLedger = ledger;
          confidence = 0.90;
          reasoning = `Found exact ledger name "${ledger}" in narration`;
          break;
        }
      }
    }

    // If confidence is high (>= 90%), keep it. Otherwise, use Suspense fallback!
    if (matchedLedger && confidence >= 0.90) {
      return {
        ...tx,
        tallyLedger: matchedLedger,
        confidence,
        reasoning
      } as MappedTransaction;
    } else {
      if (hasSuspense) {
        return {
          ...tx,
          tallyLedger: 'Suspense Account',
          confidence: matchedLedger ? confidence : 0,
          reasoning: matchedLedger
            ? `Match found ("${matchedLedger}" with ${(confidence*100).toFixed(0)}% confidence) is below 90% threshold. Fallback to Suspense.`
            : 'No reliable local match. Fallback to Suspense.'
        } as MappedTransaction;
      } else {
        return {
          ...tx,
          tallyLedger: '',
          confidence: matchedLedger ? confidence : 0,
          reasoning: matchedLedger
            ? `Match found ("${matchedLedger}" with ${(confidence*100).toFixed(0)}% confidence) is below 90% threshold. "Suspense Account" not found.`
            : 'No reliable local match. "Suspense Account" not found.'
        } as MappedTransaction;
      }
    }
  });

  return results;
}
