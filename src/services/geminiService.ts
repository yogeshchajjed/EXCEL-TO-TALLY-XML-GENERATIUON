import { GoogleGenAI, Type, ThinkingLevel } from "@google/genai";

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

export interface ColumnMapping {
  excelColumn: string;
  tallyField: string;
  confidence: number;
  reasoning: string;
}

export async function getAIColumnMapping(headers: string[]): Promise<ColumnMapping[]> {
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

  const response = await ai.models.generateContent({
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
}

export interface MappedTransaction extends BankTransaction {
  tallyLedger: string;
  confidence: number;
  reasoning: string;
}

export async function parseBankStatementText(text: string): Promise<BankTransaction[]> {
  const prompt = `
    Extract bank transactions from the following text extracted from a PDF bank statement.
    Return a JSON array of objects with the following keys: date, description, amount.
    The "description" field should contain the full transaction narration/particulars.
    Ensure amounts are numbers. If a transaction has both debit and credit, use a single amount field (positive for credit/deposit, negative for debit/withdrawal if that's the convention, or just provide the absolute value and we will map it later).
    
    Text: ${text}
  `;

  const response = await ai.models.generateContent({
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
  const prompt = `
    You are an expert Indian Accountant. Your task is to map bank statement transactions to the correct Tally Ledger.
    
    CRITICAL RULES:
    1. PRIORITIZE HISTORICAL DATA: Look at the "Historical Mappings" provided below. If a new transaction's description (Narration) contains keywords or patterns similar to a historical narration, use the same ledger.
    2. AVOID SUSPENSE: Do not map to "Suspense Account" or "Unknown" unless there is absolutely no clue. Try to categorize based on common business patterns (e.g., "UPI" often relates to small expenses or sales, "Salary" to Salary Ledger, "Rent" to Rent Ledger).
    3. FUZZY MATCHING: Even if the narration is slightly different (e.g., "UPI/123/Zomato" vs "Zomato Payment"), identify the core keyword ("Zomato") and map it to the corresponding ledger.
    4. KEYWORD MATCHING: If a description (Narration) contains a company name (e.g., "Airtel", "HDFC Life", "Zomato"), look for a ledger that matches that name or its category (e.g., Telephone Expenses, Insurance, Staff Welfare).
    5. REASONING: For every mapping, provide a clear reason (e.g., "Matched keyword 'Zomato' from historical narration data" or "Categorized as 'Bank Charges' based on description pattern").
    
    Available Tally Ledgers: ${tallyLedgers.join(', ')}
    
    ${historicalMappings && historicalMappings.length > 0 ? `
    Historical Mappings (Reference these for patterns):
    ${historicalMappings.slice(0, 100).map(m => `- "${m.narration}" was mapped to "${m.ledger}"`).join('\n')}
    ` : ''}
    
    New Transactions to Map: ${JSON.stringify(transactions)}
    
    Return a JSON array of mapped transactions.
  `;

  const response = await ai.models.generateContent({
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
            date: { type: Type.STRING },
            description: { type: Type.STRING },
            reference: { type: Type.STRING },
            amount: { type: Type.NUMBER },
            tallyLedger: { type: Type.STRING },
            confidence: { type: Type.NUMBER },
            reasoning: { type: Type.STRING }
          },
          required: ["date", "description", "amount", "tallyLedger", "confidence", "reasoning"]
        }
      }
    }
  });

  return JSON.parse(response.text);
}
