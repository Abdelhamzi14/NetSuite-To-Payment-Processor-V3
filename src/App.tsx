import React, { useState, useMemo, useEffect } from 'react';
import { Upload, FileSpreadsheet, CheckCircle2, AlertCircle, Download, ArrowRight, Table as TableIcon, Settings2, RefreshCw, Layers, Filter, ShieldCheck, Lock, Trash2 } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as xlsx from 'xlsx';
import ExcelJS from 'exceljs';
import CryptoJS from 'crypto-js';

// Generate a random session key for in-memory encryption
const SESSION_KEY = CryptoJS.lib.WordArray.random(32).toString();

interface ColumnMapping {
  transaction_id: string;
  secondary_transaction_id?: string;
  amount: string;
  account_type?: string;
  card_type?: string;
  payment_processor?: string;
  account?: string;
  document_number?: string;
  date?: string;
}

interface ReconciliationResult {
  transaction_id: any;
  sigma_amount: number;
  netsuite_amount: number;
  status: string;
  variance?: number;
  explanation?: string;
  account_type?: string;
  card_type?: string;
  document_number?: string;
  account?: string;
  sigma_date?: string;
  netsuite_date?: string;
  sigma_row?: any;
  netsuite_row?: any;
  sigma_count?: number;
  netsuite_count?: number;
}

interface ReconciliationSummary {
  totalTransactionsUploaded: number;
  matchedCount: number;
  mismatchedCount: number;
  missingInNetsuiteCount: number;
  missingInSigmaCount: number;
  matchRate: string;
  totalVariance: string;
  sigmaTotalUpload: string;
  netsuiteTotalUpload: string;
  totalUnmatchedItems: number;
}

export default function App() {
  const getRowValue = (row: any, targetKey: string) => {
    if (!targetKey) return undefined;
    if (row[targetKey] !== undefined) return row[targetKey];
    const lowTarget = targetKey.toLowerCase().trim();
    const actualKey = Object.keys(row).find(k => k.toLowerCase().trim() === lowTarget);
    return actualKey ? row[actualKey] : undefined;
  };

  const parseAmount = (val: any): number => {
    if (val === undefined || val === null || val === '') return 0;
    if (typeof val === 'number') return val;
    if (val instanceof Date) return 0; // Dates should not be treated as amounts
    
    let str = String(val).trim();
    if (!str) return 0;

    // Handle accounting negative: (1,234.56) -> -1234.56
    const isAccountingNegative = str.startsWith('(') && str.endsWith(')');
    
    // Remove characters that might interfere with parsing, but keep minus sign and decimal
    // We already identified accounting negative, so we can strip ( and ) now
    let cleanStr = str.replace(/[^\d.-]/g, '');
    
    // If it was in parentheses and doesn't already have a minus, make it negative
    if (isAccountingNegative && !cleanStr.startsWith('-')) {
      cleanStr = '-' + cleanStr;
    }
    
    const num = parseFloat(cleanStr);
    if (isNaN(num)) return 0;
    // Round to 2 decimal places to avoid floating point precision issues commonly found in Excel exports
    return Math.round(num * 100) / 100;
  };

  const formatAmount = (val: any) => {
    if (val === undefined || val === null || val === '') return '-';
    if (val instanceof Date) return val.toLocaleDateString();
    const num = parseAmount(val);
    return num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  };

  const formatInteger = (val: any) => {
    if (val === undefined || val === null || val === '') return '-';
    if (val instanceof Date) return val.toLocaleDateString();
    const num = typeof val === 'string' ? parseInt(val.replace(/,/g, ''), 10) : val;
    if (isNaN(num) || typeof num !== 'number') return String(val);
    return num.toLocaleString('en-US');
  };

  const formatDate = (val: any) => {
    if (!val || val === "N/A" || val === "-") return "N/A";
    if (val instanceof Date) {
      return val.toLocaleDateString();
    }
    // Handle cases where it might be a date string already or something else
    return String(val);
  };

  const [sigmaDataEncrypted, setSigmaDataEncrypted] = useState<any>(null);
  const [netsuiteDataEncrypted, setNetsuiteDataEncrypted] = useState<any>(null);
  const [sigmaFiles, setSigmaFiles] = useState<File[]>([]);
  const [netsuiteFiles, setNetsuiteFiles] = useState<File[]>([]);
  const [sigmaLoading, setSigmaLoading] = useState(false);
  const [netsuiteLoading, setNetsuiteLoading] = useState(false);
  const [sigmaColumns, setSigmaColumns] = useState<string[]>([]);
  const [netsuiteColumns, setNetsuiteColumns] = useState<string[]>([]);
  const [loading, setLoading] = useState(false);
  const [step, setStep] = useState(1); // 1: Upload, 2: Mapping, 3: Results
  const [activeTab, setActiveTab] = useState<'all' | 'matched' | 'mismatched' | 'missing'>('all');
  
  // Decrypt data on the fly when needed
  const sigmaData = useMemo(() => {
    if (!sigmaDataEncrypted) return [];
    
    // Handle raw data fallback for very large files
    if (typeof sigmaDataEncrypted === 'object' && sigmaDataEncrypted._isRaw) {
      return sigmaDataEncrypted.data;
    }
    
    if (typeof sigmaDataEncrypted === 'string' && sigmaDataEncrypted.startsWith("PLAIN:")) {
      try {
        return JSON.parse(sigmaDataEncrypted.substring(6));
      } catch (e) {
        return [];
      }
    }
    
    try {
      const bytes = CryptoJS.AES.decrypt(sigmaDataEncrypted, SESSION_KEY);
      const decryptedText = bytes.toString(CryptoJS.enc.Utf8);
      if (!decryptedText) return [];
      return JSON.parse(decryptedText);
    } catch (e) {
      return [];
    }
  }, [sigmaDataEncrypted]);

  const netsuiteData = useMemo(() => {
    if (!netsuiteDataEncrypted) return [];
    
    // Handle raw data fallback for very large files
    if (typeof netsuiteDataEncrypted === 'object' && netsuiteDataEncrypted._isRaw) {
      return netsuiteDataEncrypted.data;
    }
    
    if (typeof netsuiteDataEncrypted === 'string' && netsuiteDataEncrypted.startsWith("PLAIN:")) {
      try {
        return JSON.parse(netsuiteDataEncrypted.substring(6));
      } catch (e) {
        return [];
      }
    }
    
    try {
      const bytes = CryptoJS.AES.decrypt(netsuiteDataEncrypted, SESSION_KEY);
      const decryptedText = bytes.toString(CryptoJS.enc.Utf8);
      if (!decryptedText) return [];
      return JSON.parse(decryptedText);
    } catch (e) {
      return [];
    }
  }, [netsuiteDataEncrypted]);
  
  const [sigmaMapping, setSigmaMapping] = useState<ColumnMapping>({
    transaction_id: '',
    secondary_transaction_id: '',
    amount: '',
    account_type: '',
    card_type: '',
    date: ''
  });
  
  const [netsuiteMapping, setNetsuiteMapping] = useState<ColumnMapping>({
    transaction_id: '',
    secondary_transaction_id: '',
    amount: '',
    account_type: '',
    card_type: '',
    payment_processor: '',
    account: '',
    document_number: '',
    date: ''
  });

  const sigmaTotal = useMemo(() => {
    if (!sigmaData.length || !sigmaMapping.amount) return 0;
    const rawSum = sigmaData.reduce((acc, row) => acc + parseAmount(getRowValue(row, sigmaMapping.amount)), 0);
    return Math.round(rawSum * 100) / 100;
  }, [sigmaData, sigmaMapping.amount]);

  const netsuiteTotal = useMemo(() => {
    if (!netsuiteData.length || !netsuiteMapping.amount) return 0;
    const rawSum = netsuiteData.reduce((acc, row) => acc + parseAmount(getRowValue(row, netsuiteMapping.amount)), 0);
    return Math.round(rawSum * 100) / 100;
  }, [netsuiteData, netsuiteMapping.amount]);

  const [results, setResults] = useState<ReconciliationResult[]>([]);
  const [summary, setSummary] = useState<ReconciliationSummary | null>(null);
  const [dateGenerated, setDateGenerated] = useState<string | null>(null);

  const handleFilesUpload = async (type: 'sigma' | 'netsuite', files: FileList | File[]) => {
    if (type === 'sigma') setSigmaLoading(true);
    else setNetsuiteLoading(true);
    setLoading(true);

    try {
      const fileList = Array.from(files);
      const allCombinedData: any[] = type === 'sigma' ? [...sigmaData] : [...netsuiteData];
      const addedFiles: File[] = [];

      for (const file of fileList) {
        // Skip if file already added (by name and size)
        const isDuplicate = (type === 'sigma' ? sigmaFiles : netsuiteFiles).some(f => f.name === file.name && f.size === file.size);
        if (isDuplicate) continue;

        const buffer = await file.arrayBuffer();
        const workbook = xlsx.read(buffer, { 
          type: "array",
          cellStyles: false,
          cellHTML: false,
          cellFormula: false,
          cellDates: true,
          dense: true
        });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        
        // Use raw: true to get numeric values directly when available, avoiding formatting issues
        const data = xlsx.utils.sheet_to_json(sheet, { defval: "", raw: true });
        
        const filteredData = data.filter((row: any) => {
          const values = Object.values(row).map(v => String(v).toLowerCase().trim());
          if (values.length === 0) return false;
          
          // More precise total row detection:
          // Often total rows have words like 'total' or 'summary' in a cell, but are NOT the primary data.
          // We also filter out rows that have fewer than 2 non-empty values as they are usually noise.
          const nonEmptyValues = Object.values(row).filter(v => v !== null && v !== undefined && v !== "");
          if (nonEmptyValues.length < 2) return false;

          return !values.some(v => 
            v === "total" || 
            v === "grand total" || 
            v === "overall total" || 
            v === "subtotal" ||
            v.startsWith("total ") ||
            v.startsWith("summary")
          );
        });

        for (const row of filteredData) {
          allCombinedData.push(row);
        }
        addedFiles.push(file);
      }

      if (addedFiles.length === 0 && fileList.length > 0) {
        alert("Files already uploaded.");
        return;
      }

      // Collect all unique columns from all rows to ensure dropdowns are complete
      const columnSet = new Set<string>();
      allCombinedData.forEach(row => {
        if (row && typeof row === 'object') {
          Object.keys(row).forEach(k => columnSet.add(k));
        }
      });
      const columns = Array.from(columnSet);

      // Encrypt combined data
      let storageValue: any = null;
      try {
        const jsonString = JSON.stringify(allCombinedData);
        if (jsonString.length < 25 * 1024 * 1024) { 
          storageValue = CryptoJS.AES.encrypt(jsonString, SESSION_KEY).toString();
        } else {
          storageValue = { _isRaw: true, data: allCombinedData };
        }
      } catch (encryptError) {
        storageValue = { _isRaw: true, data: allCombinedData };
      }

      if (type === 'sigma') {
        setSigmaDataEncrypted(storageValue);
        setSigmaColumns(columns);
        setSigmaFiles(prev => [...prev, ...addedFiles]);
        if (!sigmaMapping.transaction_id) {
          const transactionIdCol = columns.find((c: string) => c.toLowerCase().includes('transaction id') || c.toLowerCase().includes('id'));
          const amountCol = columns.find((c: string) => {
            const low = c.toLowerCase().trim();
            return ['amount', 'amt'].some(kw => low.includes(kw));
          });
          setSigmaMapping({
            transaction_id: transactionIdCol || '',
            secondary_transaction_id: '',
            amount: amountCol || '',
            account_type: columns.find((c: string) => c.toLowerCase().includes('type')) || '',
            card_type: columns.find((c: string) => c.toLowerCase().includes('card')) || '',
            date: columns.find((c: string) => c.toLowerCase().includes('date')) || ''
          });
        }
      } else {
        setNetsuiteDataEncrypted(storageValue);
        setNetsuiteColumns(columns);
        setNetsuiteFiles(prev => [...prev, ...addedFiles]);
        if (!netsuiteMapping.transaction_id) {
          const transactionIdCol = columns.find((c: string) => c.toLowerCase().includes('transaction id') || (c.toLowerCase().includes('id') && !c.toLowerCase().includes('internal')));
          const amountCol = columns.find((c: string) => {
            const low = c.toLowerCase().trim();
            return ['amount', 'amt'].some(kw => low.includes(kw));
          });
          setNetsuiteMapping({
            transaction_id: transactionIdCol || '',
            secondary_transaction_id: '',
            amount: amountCol || '',
            account_type: columns.find((c: string) => c.toLowerCase().includes('type')) || '',
            card_type: columns.find((c: string) => c.toLowerCase().includes('card')) || '',
            payment_processor: columns.find((c: string) => c.toLowerCase().includes('processor')) || '',
            account: columns.find((c: string) => c.toLowerCase().includes('account')) || '',
            document_number: columns.find((c: string) => c.toLowerCase().includes('doc') || c.toLowerCase().includes('number')) || '',
            date: columns.find((c: string) => c.toLowerCase().includes('date')) || ''
          });
        }
      }
    } catch (error) {
      console.error('Upload error:', error);
      alert(`Upload error:\n${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setLoading(false);
      if (type === 'sigma') setSigmaLoading(false);
      else setNetsuiteLoading(false);
    }
  };

  const removeFile = async (type: 'sigma' | 'netsuite', fileName: string) => {
    setLoading(true);
    try {
      if (type === 'sigma') {
        const newFiles = sigmaFiles.filter(f => f.name !== fileName);
        setSigmaFiles(newFiles);
        if (newFiles.length === 0) {
          setSigmaDataEncrypted(null);
          setSigmaColumns([]);
        } else {
          // Re-process remaining files to get combined data
          // For simplicity, we could also store data per file in a map, but re-processing is safer for memory consistency
          // However, for UX we might just want to clear everything or support individual removal.
          // Let's implement full re-processing for removal.
          const combinedData: any[] = [];
          for (const file of newFiles) {
            const buffer = await file.arrayBuffer();
            const workbook = xlsx.read(buffer, { type: "array", cellDates: true, dense: true });
            const data = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: "", raw: true });
            const filteredData = data.filter((row: any) => {
              const values = Object.values(row).map(v => String(v).toLowerCase().trim());
              if (values.length === 0) return false;
              
              const nonEmptyValues = Object.values(row).filter(v => v !== null && v !== undefined && v !== "");
              if (nonEmptyValues.length < 2) return false;

              return !values.some(v => 
                v === "total" || 
                v === "grand total" || 
                v === "overall total" || 
                v === "subtotal" ||
                v.startsWith("total ") ||
                v.startsWith("summary")
              );
            });
            for (const row of filteredData) {
              combinedData.push(row);
            }
          }
          
          const columnSet = new Set<string>();
          combinedData.forEach(row => {
            if (row && typeof row === 'object') {
              Object.keys(row).forEach(k => columnSet.add(k));
            }
          });
          setSigmaColumns(Array.from(columnSet));
          
          let storageValue: any = null;
          const jsonString = JSON.stringify(combinedData);
          if (jsonString.length < 25 * 1024 * 1024) { 
            storageValue = CryptoJS.AES.encrypt(jsonString, SESSION_KEY).toString();
          } else {
            storageValue = { _isRaw: true, data: combinedData };
          }
          setSigmaDataEncrypted(storageValue);
        }
      } else {
        const newFiles = netsuiteFiles.filter(f => f.name !== fileName);
        setNetsuiteFiles(newFiles);
        if (newFiles.length === 0) {
          setNetsuiteDataEncrypted(null);
          setNetsuiteColumns([]);
        } else {
          const combinedData: any[] = [];
          for (const file of newFiles) {
            const buffer = await file.arrayBuffer();
            const workbook = xlsx.read(buffer, { type: "array", cellDates: true, dense: true });
            const data = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: "", raw: true });
            const filteredData = data.filter((row: any) => {
              const values = Object.values(row).map(v => String(v).toLowerCase().trim());
              if (values.length === 0) return false;
              
              const nonEmptyValues = Object.values(row).filter(v => v !== null && v !== undefined && v !== "");
              if (nonEmptyValues.length < 2) return false;

              return !values.some(v => 
                v === "total" || 
                v === "grand total" || 
                v === "overall total" || 
                v === "subtotal" ||
                v.startsWith("total ") ||
                v.startsWith("summary")
              );
            });
            for (const row of filteredData) {
              combinedData.push(row);
            }
          }
          
          const columnSet = new Set<string>();
          combinedData.forEach(row => {
            if (row && typeof row === 'object') {
              Object.keys(row).forEach(k => columnSet.add(k));
            }
          });
          setNetsuiteColumns(Array.from(columnSet));
          
          let storageValue: any = null;
          const jsonString = JSON.stringify(combinedData);
          if (jsonString.length < 25 * 1024 * 1024) { 
            storageValue = CryptoJS.AES.encrypt(jsonString, SESSION_KEY).toString();
          } else {
            storageValue = { _isRaw: true, data: combinedData };
          }
          setNetsuiteDataEncrypted(storageValue);
        }
      }
    } catch (error) {
      console.error('Remove error:', error);
    } finally {
      setLoading(false);
    }
  };

  const aggregateRows = (data: any[], idCol: string, amountCol: string, joinCols: (string | undefined)[] = []) => {
    if (!idCol || !amountCol) return data;
    const map = new Map<string, { row: any, count: number, extraData: Record<string, Set<string>> }>();
    
    // Filter out undefined and actual ID/Amount columns from join list to avoid unnecessary overhead
    const columnsToJoin = joinCols.filter((c): c is string => !!c && c !== idCol && c !== amountCol);

    data.forEach(row => {
      const idVal = getRowValue(row, idCol);
      const id = idVal !== undefined && idVal !== null ? String(idVal).trim() : "";
      const amount = parseAmount(getRowValue(row, amountCol));
      
      if (!id || id === "" || id.toLowerCase() === "null" || id.toLowerCase() === "undefined") {
        // For rows without ID, we don't aggregate them to avoid losing data
        const uniqueKey = `EMPTY_ID_${Math.random().toString(36).substring(7)}`;
        map.set(uniqueKey, { row: { ...row }, count: 1, extraData: {} });
        return;
      }

      const item = map.get(id);
      if (!item) {
        const extraData: Record<string, Set<string>> = {};
        columnsToJoin.forEach(col => {
          const val = getRowValue(row, col);
          const set = new Set<string>();
          if (val !== undefined && val !== null && String(val).trim() !== "") {
            set.add(String(val).trim());
          }
          extraData[col] = set;
        });
        map.set(id, { row: { ...row, [amountCol]: amount }, count: 1, extraData });
      } else {
        // Sum total amount with precision handling (2 decimal places for currency)
        const currentAmount = item.row[amountCol] || 0;
        item.row[amountCol] = Math.round((currentAmount + amount) * 100) / 100;
        item.count += 1;
        
        // Collect extra joined data
        columnsToJoin.forEach(col => {
          const val = getRowValue(row, col);
          if (val !== undefined && val !== null && String(val).trim() !== "") {
            item.extraData[col].add(String(val).trim());
          }
        });
      }
    });
    
    return Array.from(map.values()).map(item => {
      const finalRow = { ...item.row, _aggregated_count: item.count };
      
      // Update the row with joined strings for columns with multiple distinct values
      Object.entries(item.extraData).forEach(([col, set]) => {
        if (set.size > 1) {
          // Join unique values
          const joinedValue = Array.from(set).join(", ");
          
          // Ensure we update the correct key in the final row object
          // We check if the col exists directly or find its case-insensitive match
          if (finalRow[col] !== undefined) {
            finalRow[col] = joinedValue;
          } else {
            const lowCol = col.toLowerCase().trim();
            const actualKey = Object.keys(item.row).find(k => k.toLowerCase().trim() === lowCol);
            if (actualKey) {
              finalRow[actualKey] = joinedValue;
            } else {
              // Fallback
              finalRow[col] = joinedValue;
            }
          }
        }
      });
      
      return finalRow;
    });
  };

  const startReconciliation = async () => {
    if (!sigmaData.length || !netsuiteData.length) return;
    setLoading(true);
    
    // Use setTimeout to allow UI to update to "Processing..." before heavy synchronous work blocks the main thread
    setTimeout(() => {
      try {
        // Aggregate duplicates first as requested
        const sigmaJoinCols = [sigmaMapping.secondary_transaction_id, sigmaMapping.account_type, sigmaMapping.card_type, sigmaMapping.date];
        const netsuiteJoinCols = [netsuiteMapping.secondary_transaction_id, netsuiteMapping.account_type, netsuiteMapping.card_type, netsuiteMapping.account, netsuiteMapping.document_number, netsuiteMapping.date];
        
        const aggregatedSigmaData = aggregateRows(sigmaData, sigmaMapping.transaction_id, sigmaMapping.amount, sigmaJoinCols);
        const aggregatedNetsuiteData = aggregateRows(netsuiteData, netsuiteMapping.transaction_id, netsuiteMapping.amount, netsuiteJoinCols);

        const newResults: ReconciliationResult[] = [];
        const usedNetsuiteIndices = new Set<number>();

        let sigmaTotalUpload = 0;
        let netsuiteTotalUpload = 0;

        // Calculate totals from original data (to ensure accuracy)
        sigmaData.forEach((row: any) => {
          sigmaTotalUpload += parseAmount(getRowValue(row, sigmaMapping.amount));
        });
        netsuiteData.forEach((row: any) => {
          netsuiteTotalUpload += parseAmount(getRowValue(row, netsuiteMapping.amount));
        });

        sigmaTotalUpload = Math.round(sigmaTotalUpload * 100) / 100;
        netsuiteTotalUpload = Math.round(netsuiteTotalUpload * 100) / 100;

        const primaryNetsuiteMap = new Map<string, any[]>();
        const secondaryNetsuiteMap = new Map<string, any[]>();

        // Build mapping from aggregated NetSuite data
        aggregatedNetsuiteData.forEach((row, index) => {
          const primaryIdVal = getRowValue(row, netsuiteMapping.transaction_id);
          const primaryId = primaryIdVal !== undefined && primaryIdVal !== null ? String(primaryIdVal).trim() : "";
          
          if (primaryId && !primaryId.startsWith("EMPTY_ID_")) {
            if (!primaryNetsuiteMap.has(primaryId)) {
              primaryNetsuiteMap.set(primaryId, []);
            }
            primaryNetsuiteMap.get(primaryId)!.push({ row, index });
          }

          if (netsuiteMapping.secondary_transaction_id) {
            const secondaryIdVal = getRowValue(row, netsuiteMapping.secondary_transaction_id);
            const secondaryId = secondaryIdVal !== undefined && secondaryIdVal !== null ? String(secondaryIdVal).trim() : "";
            
            if (secondaryId && secondaryId !== "undefined" && secondaryId !== "null" && secondaryId !== "" && !secondaryId.startsWith("EMPTY_ID_")) {
              if (!secondaryNetsuiteMap.has(secondaryId)) {
                secondaryNetsuiteMap.set(secondaryId, []);
              }
              secondaryNetsuiteMap.get(secondaryId)!.push({ row, index });
            }
          }
        });

        let matchedCount = 0;
        let mismatchedCount = 0;
        let missingInNetsuiteCount = 0;
        let missingInSigmaCount = 0;

        // Reconcile aggregated Sigma data against aggregated NetSuite data
        aggregatedSigmaData.forEach((sigmaRow: any) => {
          const sigmaPrimaryIdVal = getRowValue(sigmaRow, sigmaMapping.transaction_id);
          const sigmaPrimaryId = sigmaPrimaryIdVal !== undefined && sigmaPrimaryIdVal !== null ? String(sigmaPrimaryIdVal).trim() : "";
          
          const sigmaSecondaryIdVal = sigmaMapping.secondary_transaction_id ? getRowValue(sigmaRow, sigmaMapping.secondary_transaction_id) : '';
          const sigmaSecondaryId = sigmaSecondaryIdVal !== undefined && sigmaSecondaryIdVal !== null ? String(sigmaSecondaryIdVal).trim() : '';
          
          const sigmaAmount = parseAmount(getRowValue(sigmaRow, sigmaMapping.amount));
          const sigmaCount = sigmaRow._aggregated_count || 1;

          let nsMatch = null;
          let nsMatchIndex = -1;
          let matchMethod = "";

          // 1. Primary Match: Sigma[Primary] == NetSuite[Primary]
          if (sigmaPrimaryId && sigmaPrimaryId !== "" && sigmaPrimaryId !== "null" && sigmaPrimaryId !== "undefined" && !sigmaPrimaryId.startsWith("EMPTY_ID_")) {
            const nsPrimaryMatches = primaryNetsuiteMap.get(sigmaPrimaryId);
            if (nsPrimaryMatches) {
              const matchObj = nsPrimaryMatches.find(m => !usedNetsuiteIndices.has(m.index));
              if (matchObj) {
                nsMatch = matchObj.row;
                nsMatchIndex = matchObj.index;
                matchMethod = "Primary ID Match";
              }
            }
          }

          // 2. Secondary Match: Sigma[Secondary] == NetSuite[Secondary] (only if no primary match found)
          if (nsMatchIndex === -1 && sigmaSecondaryId && sigmaSecondaryId !== "" && sigmaSecondaryId !== "null" && sigmaSecondaryId !== "undefined" && netsuiteMapping.secondary_transaction_id && !sigmaSecondaryId.startsWith("EMPTY_ID_")) {
            const nsMatchesByBothSecondary = secondaryNetsuiteMap.get(sigmaSecondaryId);
            if (nsMatchesByBothSecondary) {
              const matchObj = nsMatchesByBothSecondary.find(m => !usedNetsuiteIndices.has(m.index));
              if (matchObj) {
                nsMatch = matchObj.row;
                nsMatchIndex = matchObj.index;
                matchMethod = "Secondary ID Match";
              }
            }
          }

          if (nsMatchIndex !== -1 && nsMatch) {
            const nsAmount = parseAmount(getRowValue(nsMatch, netsuiteMapping.amount));
            const nsCount = nsMatch._aggregated_count || 1;
            const variance = sigmaAmount - nsAmount;
            usedNetsuiteIndices.add(nsMatchIndex);

            const isDuplicate = sigmaCount > 1 || nsCount > 1;
            const duplicateInfo = isDuplicate ? ` (Grouped: ${sigmaCount} Sigma items vs ${nsCount} NS items)` : "";

            if (Math.abs(variance) < 0.01) {
                matchedCount++;
              newResults.push({
                transaction_id: sigmaPrimaryId.startsWith("EMPTY_ID_") ? "N/A (No ID)" : sigmaPrimaryId,
                sigma_amount: sigmaAmount,
                netsuite_amount: nsAmount,
                status: "Matched",
                variance: 0,
                explanation: `${matchMethod} and Amount match perfectly${duplicateInfo}`,
                account_type: getRowValue(nsMatch, netsuiteMapping.account_type!) || getRowValue(sigmaRow, sigmaMapping.account_type!) || "N/A",
                card_type: getRowValue(nsMatch, netsuiteMapping.card_type!) || getRowValue(sigmaRow, sigmaMapping.card_type!) || "N/A",
                document_number: getRowValue(nsMatch, netsuiteMapping.document_number!) || "N/A",
                account: getRowValue(nsMatch, netsuiteMapping.account!) || "N/A",
                sigma_date: sigmaMapping.date ? getRowValue(sigmaRow, sigmaMapping.date) : "N/A",
                netsuite_date: netsuiteMapping.date ? getRowValue(nsMatch, netsuiteMapping.date) : "N/A",
                sigma_row: sigmaRow,
                netsuite_row: nsMatch,
                sigma_count: sigmaCount,
                netsuite_count: nsCount
              });
            } else {
              mismatchedCount++;
              newResults.push({
                transaction_id: sigmaPrimaryId.startsWith("EMPTY_ID_") ? "N/A (No ID)" : sigmaPrimaryId,
                sigma_amount: sigmaAmount,
                netsuite_amount: nsAmount,
                status: "Mismatched",
                variance: variance,
                explanation: `${matchMethod} found, but amount discrepancy: Sigma ($${sigmaAmount.toFixed(2)}) vs NetSuite ($${nsAmount.toFixed(2)})${duplicateInfo}`,
                account_type: getRowValue(nsMatch, netsuiteMapping.account_type!) || getRowValue(sigmaRow, sigmaMapping.account_type!) || "N/A",
                card_type: getRowValue(nsMatch, netsuiteMapping.card_type!) || getRowValue(sigmaRow, sigmaMapping.card_type!) || "N/A",
                document_number: getRowValue(nsMatch, netsuiteMapping.document_number!) || "N/A",
                account: getRowValue(nsMatch, netsuiteMapping.account!) || "N/A",
                sigma_date: sigmaMapping.date ? getRowValue(sigmaRow, sigmaMapping.date) : "N/A",
                netsuite_date: netsuiteMapping.date ? getRowValue(nsMatch, netsuiteMapping.date) : "N/A",
                sigma_row: sigmaRow,
                netsuite_row: nsMatch,
                sigma_count: sigmaCount,
                netsuite_count: nsCount
              });
            }
          } else {
            missingInNetsuiteCount++;
            newResults.push({
              transaction_id: sigmaPrimaryId.startsWith("EMPTY_ID_") ? "N/A (No ID)" : sigmaPrimaryId,
              sigma_amount: sigmaAmount,
              netsuite_amount: 0,
              status: "Missing in NetSuite",
              variance: sigmaAmount,
              explanation: `Transaction ID found in Sigma but missing in NetSuite${sigmaCount > 1 ? ` (Grouped: ${sigmaCount} items)` : ""}`,
              account_type: getRowValue(sigmaRow, sigmaMapping.account_type!) || "N/A",
              card_type: getRowValue(sigmaRow, sigmaMapping.card_type!) || "N/A",
              document_number: "N/A",
              account: "N/A",
              sigma_date: sigmaMapping.date ? getRowValue(sigmaRow, sigmaMapping.date) : "N/A",
              netsuite_date: "N/A",
              sigma_row: sigmaRow,
              sigma_count: sigmaCount
            });
          }
        });

        aggregatedNetsuiteData.forEach((nsRow, index) => {
          if (!usedNetsuiteIndices.has(index)) {
            missingInSigmaCount++;
            const nsAmount = parseAmount(getRowValue(nsRow, netsuiteMapping.amount));
            const nsCount = nsRow._aggregated_count || 1;
            const nsIdVal = getRowValue(nsRow, netsuiteMapping.transaction_id);
            const nsId = nsIdVal !== undefined && nsIdVal !== null ? String(nsIdVal).trim() : "";
            
            newResults.push({
              transaction_id: nsId.startsWith("EMPTY_ID_") ? "N/A (No ID)" : nsId,
              sigma_amount: 0,
              netsuite_amount: nsAmount,
              status: "Missing in Sigma",
              variance: -nsAmount,
              explanation: `Transaction ID found in NetSuite but missing in Sigma${nsCount > 1 ? ` (Grouped: ${nsCount} items)` : ""}`,
              account_type: getRowValue(nsRow, netsuiteMapping.account_type!) || "N/A",
              card_type: getRowValue(nsRow, netsuiteMapping.card_type!) || "N/A",
              document_number: getRowValue(nsRow, netsuiteMapping.document_number!) || "N/A",
              account: getRowValue(nsRow, netsuiteMapping.account!) || "N/A",
              sigma_date: "N/A",
              netsuite_date: netsuiteMapping.date ? getRowValue(nsRow, netsuiteMapping.date) : "N/A",
              netsuite_row: nsRow,
              netsuite_count: nsCount
            });
          }
        });

        const totalTransactionsUploaded = aggregatedSigmaData.length + aggregatedNetsuiteData.length;
        const matchRate = (matchedCount / aggregatedSigmaData.length) * 100;

        setResults(newResults);
        setSummary({
          totalTransactionsUploaded,
          matchedCount,
          mismatchedCount,
          missingInNetsuiteCount,
          missingInSigmaCount,
          matchRate: matchRate.toFixed(2),
          totalVariance: (sigmaTotalUpload - netsuiteTotalUpload).toFixed(2),
          sigmaTotalUpload: sigmaTotalUpload.toFixed(2),
          netsuiteTotalUpload: netsuiteTotalUpload.toFixed(2),
          totalUnmatchedItems: mismatchedCount + missingInNetsuiteCount + missingInSigmaCount
        });
        setDateGenerated(new Date().toLocaleString());
        setStep(3);
      } catch (error) {
        console.error('Reconciliation error:', error);
        alert('Error during reconciliation');
      } finally {
        setLoading(false);
      }
    }, 100);
  };

  const downloadExport = async (onlyUnmatched: boolean = false) => {
    if (!summary || !results.length) return;
    
    try {
      // Sort results by status
      const sortedResults = [...results].sort((a, b) => a.status.localeCompare(b.status));
      
      const matched = sortedResults.filter((r: any) => r.status === "Matched");
      const unmatched = sortedResults.filter((r: any) => r.status !== "Matched");

      const workbook = new ExcelJS.Workbook();
      
      // 1. Summary Sheet
      const summarySheet = workbook.addWorksheet("Summary");
      
      // Title
      summarySheet.mergeCells('A1:B1');
      const titleCell = summarySheet.getCell('A1');
      titleCell.value = onlyUnmatched ? 'Reconciliation Summary (Unmatched Only)' : 'Reconciliation Summary';
      titleCell.font = { size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
      titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF141414' } };
      titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
      
      summarySheet.getRow(1).height = 30;

      // Data rows
      const summaryData = [
        ["Date Generated", new Date().toLocaleString()],
        ["Total Transactions Uploaded", summary.totalTransactionsUploaded],
        ["Sigma Total Upload", parseFloat(summary.sigmaTotalUpload)],
        ["NetSuite Total Upload", parseFloat(summary.netsuiteTotalUpload)],
        ["Match Rate", parseFloat(summary.matchRate) / 100],
        ["Matched Items Count", summary.matchedCount],
        ["Mismatched Items Count", summary.mismatchedCount],
        ["Missing in NetSuite Count", summary.missingInNetsuiteCount],
        ["Missing in Sigma Count", summary.missingInSigmaCount],
        ["Total Unmatched Items", summary.totalUnmatchedItems],
        ["Total Variance Amount", { formula: `SUM('Unmatched Items'!G2:G${unmatched.length + 1})` }]
      ];

      summaryData.forEach((row, index) => {
        const rowIndex = index + 3; // Start at row 3
        summarySheet.addRow(row);
        const labelCell = summarySheet.getCell(`A${rowIndex}`);
        const valueCell = summarySheet.getCell(`B${rowIndex}`);
        
        labelCell.font = { bold: true, size: 14 };
        valueCell.font = { size: 14 };
        labelCell.border = { bottom: { style: 'thin', color: { argb: 'FFDDDDDD' } } };
        valueCell.border = { bottom: { style: 'thin', color: { argb: 'FFDDDDDD' } } };
        
        // Formatting based on row type
        if (row[0] === "Match Rate") {
          valueCell.numFmt = '0.00%';
        } else if (row[0] === "Sigma Total Upload" || row[0] === "NetSuite Total Upload" || row[0] === "Total Variance Amount") {
          valueCell.numFmt = '"$"#,##0.00';
        } else if (typeof row[1] === 'number') {
          valueCell.numFmt = '#,##0';
        }
      });

      if (onlyUnmatched) {
        summarySheet.addRow(["Export Type", "UNMATCHED ONLY"]);
        summarySheet.getCell(`A${summaryData.length + 3}`).font = { bold: true, color: { argb: 'FFFF0000' } };
        summarySheet.getCell(`B${summaryData.length + 3}`).font = { color: { argb: 'FFFF0000' }, bold: true };
      }

      summarySheet.getColumn(1).width = 30;
      summarySheet.getColumn(2).width = 25;
      summarySheet.views = [{ showGridLines: false }];

      // 2. Matched Items Sheet (Optional)
      if (!onlyUnmatched) {
        const matchedSheet = workbook.addWorksheet("Matched Items");
        matchedSheet.views = [{ state: 'frozen', ySplit: 1 }];
        matchedSheet.columns = [
          { header: "Transaction ID", key: "id", width: 25 },
          { header: "Status", key: "status", width: 20 },
          { header: "Date", key: "date", width: 15 },
          { header: "Amount", key: "amount", width: 15 },
          { header: "Variance", key: "variance", width: 15 },
          { header: "Explanation", key: "explanation", width: 40 },
          { header: "Account Type", key: "account_type", width: 20 },
          { header: "Card Type", key: "card_type", width: 20 },
          { header: "Account", key: "account", width: 20 },
          { header: "Document Number", key: "doc_num", width: 20 },
          { header: "Sigma Count", key: "sigma_count", width: 15 },
          { header: "NS Count", key: "ns_count", width: 15 }
        ];

        const matchedHeaderRow = matchedSheet.getRow(1);
        
        // Style headers A-L (1-12) with Black background and White text
        for (let i = 1; i <= 12; i++) {
          const cell = matchedHeaderRow.getCell(i);
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } };
          cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        }

        matched.forEach((r: any) => {
          const row = matchedSheet.addRow({
            id: r.transaction_id,
            status: r.status,
            date: r.sigma_date !== "N/A" ? r.sigma_date : r.netsuite_date,
            amount: r.sigma_amount,
            variance: r.variance || 0,
            explanation: r.explanation,
            account_type: r.account_type,
            card_type: r.card_type,
            account: r.account,
            doc_num: r.document_number,
            sigma_count: r.sigma_count || 1,
            ns_count: r.netsuite_count || 1
          });
          
          // Highlight columns A to L (1 to 12) in light green
          for (let i = 1; i <= 12; i++) {
            const cell = row.getCell(i);
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCFFCC' } };
            
            // Make Variance Column Bold (Column E / 5)
            if (i === 5) {
              cell.font = { ...cell.font, bold: true };
            }
          }
        });
        
        matchedSheet.getColumn('amount').numFmt = '"$"#,##0.00';
        matchedSheet.getColumn('variance').numFmt = '"$"#,##0.00';
      }

      // 3. Unmatched Items Sheet
      const unmatchedSheet = workbook.addWorksheet("Unmatched Items");
      unmatchedSheet.views = [{ state: 'frozen', ySplit: 1 }];
      unmatchedSheet.columns = [
        { header: "Transaction ID", key: "id", width: 25 },
        { header: "Status", key: "status", width: 20 },
        { header: "Sigma Date", key: "sigma_date", width: 15 },
        { header: "NetSuite Date", key: "netsuite_date", width: 15 },
        { header: "Sigma Amount", key: "sigma_amount", width: 15 },
        { header: "NetSuite Amount", key: "netsuite_amount", width: 15 },
        { header: "Variance", key: "variance", width: 15 },
        { header: "Explanation", key: "explanation", width: 40 },
        { header: "Account Type", key: "account_type", width: 20 },
        { header: "Card Type", key: "card_type", width: 20 },
        { header: "Account", key: "account", width: 20 },
        { header: "Document Number", key: "doc_num", width: 20 },
        { header: "Sigma Count", key: "sigma_count", width: 15 },
        { header: "NS Count", key: "ns_count", width: 15 }
      ];

      // Style headers and columns A-N
      const unmatchedHeaderRow = unmatchedSheet.getRow(1);
      unmatchedHeaderRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      
      // Highlight headers A-N (1-14)
      for (let i = 1; i <= 14; i++) {
        const cell = unmatchedHeaderRow.getCell(i);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF141414' } };
      }
      
      unmatched.forEach((r: any) => {
        const row = unmatchedSheet.addRow({
          id: r.transaction_id,
          status: r.status,
          sigma_date: r.sigma_date,
          netsuite_date: r.netsuite_date,
          sigma_amount: r.sigma_amount,
          netsuite_amount: r.netsuite_amount,
          variance: r.variance,
          explanation: r.explanation,
          account_type: r.account_type,
          card_type: r.card_type,
          account: r.account,
          doc_num: r.document_number,
          sigma_count: r.sigma_count || (r.sigma_amount !== 0 ? 1 : 0),
          ns_count: r.netsuite_count || (r.netsuite_amount !== 0 ? 1 : 0)
        });

        // Highlight columns A to N (1 to 14)
        for (let i = 1; i <= 14; i++) {
          const cell = row.getCell(i);
          
          // Color coding based on status
          if (r.status === "Missing in NetSuite") {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE4E4' } }; // Light Red
          } else if (r.status === "Missing in Sigma") {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE4F0FF' } }; // Light Blue
          } else if (r.status === "Mismatched") {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; // Yellow as requested
          } else {
            // Default highlight if requested? User said "Highlight columns and headers only from A to K"
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F0F0' } }; 
          }
          
          // Make Variance Column Bold (Column G / 7)
          if (i === 7) {
            cell.font = { ...cell.font, bold: true };
          }
        }
      });
      
      // Format currency columns
      unmatchedSheet.getColumn('sigma_amount').numFmt = '"$"#,##0.00';
      unmatchedSheet.getColumn('netsuite_amount').numFmt = '"$"#,##0.00';
      unmatchedSheet.getColumn('variance').numFmt = '"$"#,##0.00';

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = onlyUnmatched ? `unmatched_report_${new Date().getTime()}.xlsx` : `full_reconciliation_${new Date().getTime()}.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error('Export error:', error);
      alert('Failed to generate export');
    }
  };

  const reset = () => {
    setSigmaDataEncrypted(null);
    setNetsuiteDataEncrypted(null);
    setSigmaFiles([]);
    setNetsuiteFiles([]);
    setSigmaColumns([]);
    setNetsuiteColumns([]);
    setResults([]);
    setSummary(null);
    setStep(1);
    setActiveTab('all');
    // Explicitly clear any sensitive data from memory if possible
    // (React state updates are asynchronous, but setting to null is the standard way)
  };

  const filteredResults = results.filter(r => {
    if (activeTab === 'all') return true;
    if (activeTab === 'matched') return r.status === 'Matched';
    if (activeTab === 'mismatched') return r.status === 'Mismatched';
    if (activeTab === 'missing') return r.status.startsWith('Missing');
    return true;
  });

  return (
    <div className="min-h-screen bg-[#E4E3E0] text-[#141414] font-sans selection:bg-[#141414] selection:text-[#E4E3E0]">
      {/* Header */}
      <header className="border-b border-[#141414] p-6 flex justify-between items-center">
        <div className="flex items-center gap-4">
          <div>
            <h1 className="text-xs uppercase tracking-widest opacity-50 font-mono mb-1">Data Match Tool v2.0</h1>
            <h2 className="text-2xl font-serif italic">Reconciliation Engine</h2>
          </div>
          <div className="flex items-center gap-2 bg-emerald-50 text-emerald-700 px-3 py-1 rounded-full border border-emerald-200 text-[10px] uppercase tracking-wider font-bold">
            <ShieldCheck size={12} /> Secure Mode Active
          </div>
        </div>
        <div className="flex gap-4">
          {step > 1 && (
            <button 
              onClick={reset}
              className="px-4 py-2 border border-[#141414] rounded-full text-xs uppercase tracking-widest hover:bg-rose-600 hover:text-white hover:border-rose-600 transition-all flex items-center gap-2"
            >
              <Trash2 size={14} /> Clear & Reset
            </button>
          )}
        </div>
      </header>

      <div className="bg-[#141414] text-[#E4E3E0] py-2 px-6 flex items-center justify-center gap-4 text-[10px] uppercase tracking-[0.2em]">
        <Lock size={10} /> 100% Client-Side Processing &bull; No Data Leaves Your Browser &bull; Files are Never Saved to Server
      </div>

      <main className="max-w-7xl mx-auto p-8">
        {/* Step Indicator */}
        <div className="flex gap-8 mb-12 border-b border-[#141414]/10 pb-4">
          {[
            { n: 1, label: 'Upload Files' },
            { n: 2, label: 'Map Columns' },
            { n: 3, label: 'Review Results' }
          ].map((s) => (
            <div key={s.n} className={`flex items-center gap-2 text-xs uppercase tracking-widest ${step === s.n ? 'opacity-100 font-bold' : 'opacity-30'}`}>
              <span className={`w-5 h-5 rounded-full border border-[#141414] flex items-center justify-center ${step === s.n ? 'bg-[#141414] text-[#E4E3E0]' : ''}`}>
                {s.n}
              </span>
              {s.label}
            </div>
          ))}
        </div>

        <AnimatePresence mode="wait">
          {step === 1 && (
            <motion.div 
              key="step1"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="grid md:grid-cols-2 gap-8"
            >
              {/* Sigma Upload */}
              <div className="border border-[#141414] p-8 bg-white/50 rounded-2xl flex flex-col items-center text-center">
                <div className="w-16 h-16 rounded-full bg-[#141414] text-[#E4E3E0] flex items-center justify-center mb-6">
                  <FileSpreadsheet size={32} />
                </div>
                <h3 className="font-serif italic text-xl mb-2">Sigma Import</h3>
                <p className="text-xs opacity-60 mb-6 max-w-xs">Upload your primary transaction data file from Sigma. Usually contains raw transaction logs.</p>
                
                {sigmaLoading ? (
                  <div className="flex items-center gap-2 text-[#141414] font-mono text-xs bg-white/80 px-4 py-2 rounded-full border border-[#141414]/20">
                    <RefreshCw size={14} className="animate-spin" /> Processing Sigma...
                  </div>
                ) : (
                  <div className="w-full space-y-4 flex flex-col items-center">
                    <label className="cursor-pointer px-8 py-3 bg-[#141414] text-[#E4E3E0] rounded-full text-xs uppercase tracking-widest hover:opacity-90 transition-opacity flex items-center gap-2">
                      <Upload size={14} /> Add Sigma Files
                      <input type="file" className="hidden" accept=".xlsx,.xls,.csv" multiple onChange={(e) => e.target.files && handleFilesUpload('sigma', e.target.files)} />
                    </label>

                    <AnimatePresence>
                      {sigmaFiles.length > 0 && (
                        <div className="w-full">
                          <div className="max-h-40 overflow-y-auto space-y-2 mt-4 px-4 w-full">
                            {sigmaFiles.map(file => (
                              <motion.div 
                                key={file.name}
                                initial={{ opacity: 0, x: -10 }}
                                animate={{ opacity: 1, x: 0 }}
                                exit={{ opacity: 0, scale: 0.9 }}
                                className="flex items-center justify-between gap-2 text-emerald-600 font-mono text-xs bg-emerald-50 px-3 py-2 rounded-lg border border-emerald-200"
                              >
                                <div className="flex items-center gap-2 truncate">
                                  <CheckCircle2 size={12} className="shrink-0" /> 
                                  <span className="truncate">{file.name}</span>
                                </div>
                                <button 
                                  onClick={() => removeFile('sigma', file.name)}
                                  className="text-emerald-800 hover:text-rose-600 p-1"
                                >
                                  <Trash2 size={12} />
                                </button>
                              </motion.div>
                            ))}
                          </div>

                          <div className="mt-6 pt-4 border-t border-[#141414]/10 w-full flex flex-col items-center">
                            <div className="flex justify-between w-full px-4 mb-1">
                              <span className="text-[10px] uppercase tracking-widest opacity-50 font-mono">Total Uploaded</span>
                              <span className="font-mono text-sm font-bold text-[#141414]">${formatAmount(sigmaTotal)}</span>
                            </div>
                            <p className="text-[9px] opacity-30 uppercase tracking-tighter">
                              Summing column: "{sigmaMapping.amount || 'Not detected'}"
                            </p>
                          </div>
                        </div>
                      )}
                    </AnimatePresence>
                  </div>
                )}
              </div>

              {/* NetSuite Upload */}
              <div className={`border border-[#141414] p-8 bg-white/50 rounded-2xl flex flex-col items-center text-center ${sigmaFiles.length === 0 ? 'opacity-30 pointer-events-none' : ''}`}>
                <div className="w-16 h-16 rounded-full bg-[#141414] text-[#E4E3E0] flex items-center justify-center mb-6">
                  <TableIcon size={32} />
                </div>
                <h3 className="font-serif italic text-xl mb-2">NetSuite Import</h3>
                <p className="text-xs opacity-60 mb-6 max-w-xs">Upload corresponding NetSuite files. We'll match this against your Sigma data.</p>
                
                {netsuiteLoading ? (
                  <div className="flex items-center gap-2 text-[#141414] font-mono text-xs bg-white/80 px-4 py-2 rounded-full border border-[#141414]/20">
                    <RefreshCw size={14} className="animate-spin" /> Processing NetSuite...
                  </div>
                ) : (
                  <div className="w-full space-y-4 flex flex-col items-center">
                    <label className="cursor-pointer px-8 py-3 bg-[#141414] text-[#E4E3E0] rounded-full text-xs uppercase tracking-widest hover:opacity-90 transition-opacity flex items-center gap-2">
                      <Upload size={14} /> Add NetSuite Files
                      <input type="file" className="hidden" accept=".xlsx,.xls,.csv" multiple onChange={(e) => e.target.files && handleFilesUpload('netsuite', e.target.files)} />
                    </label>

                    <AnimatePresence>
                      {netsuiteFiles.length > 0 && (
                        <div className="w-full">
                          <div className="max-h-40 overflow-y-auto space-y-2 mt-4 px-4 w-full">
                            {netsuiteFiles.map(file => (
                              <motion.div 
                                key={file.name}
                                initial={{ opacity: 0, x: -10 }}
                                animate={{ opacity: 1, x: 0 }}
                                exit={{ opacity: 0, scale: 0.9 }}
                                className="flex items-center justify-between gap-2 text-emerald-600 font-mono text-xs bg-emerald-50 px-3 py-2 rounded-lg border border-emerald-200"
                              >
                                <div className="flex items-center gap-2 truncate">
                                  <CheckCircle2 size={12} className="shrink-0" /> 
                                  <span className="truncate">{file.name}</span>
                                </div>
                                <button 
                                  onClick={() => removeFile('netsuite', file.name)}
                                  className="text-emerald-800 hover:text-rose-600 p-1"
                                >
                                  <Trash2 size={12} />
                                </button>
                              </motion.div>
                            ))}
                          </div>
                          
                          <div className="mt-6 pt-4 border-t border-[#141414]/10 w-full flex flex-col items-center">
                            <div className="flex justify-between w-full px-4 mb-1">
                              <span className="text-[10px] uppercase tracking-widest opacity-50 font-mono">Total Uploaded</span>
                              <span className="font-mono text-sm font-bold text-[#141414]">${formatAmount(netsuiteTotal)}</span>
                            </div>
                            <p className="text-[9px] opacity-30 uppercase tracking-tighter">
                              Summing column: "{netsuiteMapping.amount || 'Not detected'}"
                            </p>
                          </div>
                        </div>
                      )}
                    </AnimatePresence>
                  </div>
                )}
              </div>

              {sigmaFiles.length > 0 && netsuiteFiles.length > 0 && (
                <div className="md:col-span-2 flex justify-center mt-8">
                  <button 
                    onClick={() => setStep(2)}
                    className="px-12 py-4 bg-[#141414] text-[#E4E3E0] rounded-full text-xs uppercase tracking-widest hover:scale-105 transition-transform flex items-center gap-3 font-bold"
                  >
                    Continue to Mapping <ArrowRight size={16} />
                  </button>
                </div>
              )}
            </motion.div>
          )}

          {step === 2 && (
            <motion.div 
              key="step2"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="space-y-8"
            >
              <div className="grid md:grid-cols-2 gap-8">
                {/* Sigma Mapping */}
                <div className="border border-[#141414] p-6 rounded-2xl bg-white/30">
                  <div className="flex items-center gap-2 mb-6">
                    <Settings2 size={18} />
                    <h3 className="font-serif italic text-lg">Sigma Column Mapping</h3>
                  </div>
                  <div className="space-y-4">
                    <div>
                      <label className="block text-[10px] uppercase tracking-widest opacity-50 mb-1">Transaction ID</label>
                      <select 
                        value={sigmaMapping.transaction_id}
                        onChange={(e) => setSigmaMapping({...sigmaMapping, transaction_id: e.target.value})}
                        className="w-full bg-transparent border-b border-[#141414] py-2 text-sm focus:outline-none"
                      >
                        <option value="">Select Column...</option>
                        {sigmaColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>
                    <div>
                      <label className="block text-[10px] uppercase tracking-widest opacity-50 mb-1">Secondary ID (Optional)</label>
                      <select 
                        value={sigmaMapping.secondary_transaction_id}
                        onChange={(e) => setSigmaMapping({...sigmaMapping, secondary_transaction_id: e.target.value})}
                        className="w-full bg-transparent border-b border-[#141414] py-2 text-sm focus:outline-none"
                      >
                        <option value="">select secondary ID column...</option>
                        {sigmaColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>
                    <div>
                      <label className="block text-[10px] uppercase tracking-widest opacity-50 mb-1">Amount</label>
                      <select 
                        value={sigmaMapping.amount}
                        onChange={(e) => setSigmaMapping({...sigmaMapping, amount: e.target.value})}
                        className="w-full bg-transparent border-b border-[#141414] py-2 text-sm focus:outline-none"
                      >
                        <option value="">Select Column...</option>
                        {sigmaColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>
                    <div>
                      <label className="block text-[10px] uppercase tracking-widest opacity-50 mb-1">Account Type (Optional)</label>
                      <select 
                        value={sigmaMapping.account_type}
                        onChange={(e) => setSigmaMapping({...sigmaMapping, account_type: e.target.value})}
                        className="w-full bg-transparent border-b border-[#141414] py-2 text-sm focus:outline-none"
                      >
                        <option value="">Select Column...</option>
                        {sigmaColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>
                    <div>
                      <label className="block text-[10px] uppercase tracking-widest opacity-50 mb-1">Card Type (Optional)</label>
                      <select 
                        value={sigmaMapping.card_type}
                        onChange={(e) => setSigmaMapping({...sigmaMapping, card_type: e.target.value})}
                        className="w-full bg-transparent border-b border-[#141414] py-2 text-sm focus:outline-none"
                      >
                        <option value="">Select Column...</option>
                        {sigmaColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>
                    <div>
                      <label className="block text-[10px] uppercase tracking-widest opacity-50 mb-1">Date (Optional)</label>
                      <select 
                        value={sigmaMapping.date}
                        onChange={(e) => setSigmaMapping({...sigmaMapping, date: e.target.value})}
                        className="w-full bg-transparent border-b border-[#141414] py-2 text-sm focus:outline-none"
                      >
                        <option value="">Select Column...</option>
                        {sigmaColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>
                  </div>
                </div>

                {/* NetSuite Mapping */}
                <div className="border border-[#141414] p-6 rounded-2xl bg-white/30">
                  <div className="flex items-center gap-2 mb-6">
                    <Settings2 size={18} />
                    <h3 className="font-serif italic text-lg">NetSuite Column Mapping</h3>
                  </div>
                  <div className="space-y-4">
                    <div>
                      <label className="block text-[10px] uppercase tracking-widest opacity-50 mb-1">Transaction ID</label>
                      <select 
                        value={netsuiteMapping.transaction_id}
                        onChange={(e) => setNetsuiteMapping({...netsuiteMapping, transaction_id: e.target.value})}
                        className="w-full bg-transparent border-b border-[#141414] py-2 text-sm focus:outline-none"
                      >
                        <option value="">Select Column...</option>
                        {netsuiteColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>
                    <div>
                      <label className="block text-[10px] uppercase tracking-widest opacity-50 mb-1">Secondary ID (Optional)</label>
                      <select 
                        value={netsuiteMapping.secondary_transaction_id}
                        onChange={(e) => setNetsuiteMapping({...netsuiteMapping, secondary_transaction_id: e.target.value})}
                        className="w-full bg-transparent border-b border-[#141414] py-2 text-sm focus:outline-none"
                      >
                        <option value="">select secondary ID column...</option>
                        {netsuiteColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>
                    <div>
                      <label className="block text-[10px] uppercase tracking-widest opacity-50 mb-1">Amount / Total</label>
                      <select 
                        value={netsuiteMapping.amount}
                        onChange={(e) => setNetsuiteMapping({...netsuiteMapping, amount: e.target.value})}
                        className="w-full bg-transparent border-b border-[#141414] py-2 text-sm focus:outline-none"
                      >
                        <option value="">Select Column...</option>
                        {netsuiteColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>
                    <div>
                      <label className="block text-[10px] uppercase tracking-widest opacity-50 mb-1">Document Number</label>
                      <select 
                        value={netsuiteMapping.document_number}
                        onChange={(e) => setNetsuiteMapping({...netsuiteMapping, document_number: e.target.value})}
                        className="w-full bg-transparent border-b border-[#141414] py-2 text-sm focus:outline-none"
                      >
                        <option value="">Select Column...</option>
                        {netsuiteColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>
                    <div>
                      <label className="block text-[10px] uppercase tracking-widest opacity-50 mb-1">Account</label>
                      <select 
                        value={netsuiteMapping.account}
                        onChange={(e) => setNetsuiteMapping({...netsuiteMapping, account: e.target.value})}
                        className="w-full bg-transparent border-b border-[#141414] py-2 text-sm focus:outline-none"
                      >
                        <option value="">Select Column...</option>
                        {netsuiteColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>
                    <div>
                      <label className="block text-[10px] uppercase tracking-widest opacity-50 mb-1">Card Type (Optional)</label>
                      <select 
                        value={netsuiteMapping.card_type}
                        onChange={(e) => setNetsuiteMapping({...netsuiteMapping, card_type: e.target.value})}
                        className="w-full bg-transparent border-b border-[#141414] py-2 text-sm focus:outline-none"
                      >
                        <option value="">Select Column...</option>
                        {netsuiteColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>
                    <div>
                      <label className="block text-[10px] uppercase tracking-widest opacity-50 mb-1">Date (Optional)</label>
                      <select 
                        value={netsuiteMapping.date}
                        onChange={(e) => setNetsuiteMapping({...netsuiteMapping, date: e.target.value})}
                        className="w-full bg-transparent border-b border-[#141414] py-2 text-sm focus:outline-none"
                      >
                        <option value="">Select Column...</option>
                        {netsuiteColumns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>
                  </div>
                </div>
              </div>

              <div className="flex justify-center pt-8">
                <button 
                  disabled={!sigmaMapping.transaction_id || !sigmaMapping.amount || !netsuiteMapping.transaction_id || !netsuiteMapping.amount}
                  onClick={startReconciliation}
                  className="px-12 py-4 bg-[#141414] text-[#E4E3E0] rounded-full text-xs uppercase tracking-widest hover:scale-105 transition-transform disabled:opacity-30 disabled:pointer-events-none flex items-center gap-3 font-bold"
                >
                  {loading ? 'Processing...' : 'Run Reconciliation'} <RefreshCw size={16} className={loading ? 'animate-spin' : ''} />
                </button>
              </div>
            </motion.div>
          )}

          {step === 3 && (
            <motion.div 
              key="step3"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="space-y-6"
            >
              {/* Report Header */}
              <div className="border border-[#141414] p-8 rounded-2xl bg-white/80 shadow-sm mb-8">
                <h2 className="text-3xl font-serif italic mb-4 text-center border-b border-[#141414]/10 pb-4">
                  NetSuite vs Payment Processor Validation Report
                </h2>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-6 text-[11px] uppercase tracking-widest">
                  <div className="flex flex-col gap-1">
                    <span className="opacity-50">Sigma File</span>
                    <span className="font-bold">{sigmaFiles.map(f => f.name).join(', ')}</span>
                    {summary && <span className="mt-1 font-mono text-[10px] text-[#141414]/70">Total: ${formatAmount(summary.sigmaTotalUpload)}</span>}
                  </div>
                  <div className="flex flex-col gap-1">
                    <span className="opacity-50">NetSuite File</span>
                    <span className="font-bold">{netsuiteFiles.map(f => f.name).join(', ')}</span>
                    {summary && <span className="mt-1 font-mono text-[10px] text-[#141414]/70">Total: ${formatAmount(summary.netsuiteTotalUpload)}</span>}
                  </div>
                  <div className="flex flex-col gap-1">
                    <span className="opacity-50">Date Generated</span>
                    <span className="font-bold">{dateGenerated}</span>
                  </div>
                </div>
              </div>

              {/* Summary Header v2 */}
              {summary && (
                <div className="space-y-4">
                  {/* Row 1 */}
                  <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
                    <div className="border border-[#141414] p-4 rounded-2xl bg-white/50">
                      <p className="text-[10px] uppercase tracking-widest opacity-50 mb-1">Total Uploaded</p>
                      <p className="text-xl font-serif italic">{formatInteger(summary.totalTransactionsUploaded)}</p>
                    </div>
                    <div className="border border-[#141414] p-4 rounded-2xl bg-emerald-50 border-emerald-200">
                      <p className="text-[10px] uppercase tracking-widest text-emerald-700 mb-1">Matched</p>
                      <p className="text-xl font-serif italic text-emerald-900">{formatInteger(summary.matchedCount)}</p>
                    </div>
                    <div className="border border-[#141414] p-4 rounded-2xl bg-amber-50 border-amber-200">
                      <p className="text-[10px] uppercase tracking-widest text-amber-700 mb-1">Mismatched</p>
                      <p className="text-xl font-serif italic text-amber-900">{formatInteger(summary.mismatchedCount)}</p>
                    </div>
                    <div className="border border-[#141414] p-4 rounded-2xl bg-rose-50 border-rose-200">
                      <p className="text-[10px] uppercase tracking-widest text-rose-700 mb-1">MISSING IN NETSUITE</p>
                      <p className="text-xl font-serif italic text-rose-900">{formatInteger(summary.missingInNetsuiteCount)}</p>
                    </div>
                    <div className="border border-[#141414] p-4 rounded-2xl bg-blue-50 border-blue-200">
                      <p className="text-[10px] uppercase tracking-widest text-blue-700 mb-1">Missing in Sigma</p>
                      <p className="text-xl font-serif italic text-blue-900">{formatInteger(summary.missingInSigmaCount)}</p>
                    </div>
                  </div>

                  {/* Row 2 */}
                  <div className="grid grid-cols-2 gap-4">
                    <div className="border border-[#141414] p-4 rounded-2xl bg-white/50">
                      <p className="text-[10px] uppercase tracking-widest opacity-50 mb-1">Match Rate</p>
                      <p className="text-3xl font-serif italic">{summary.matchRate}%</p>
                    </div>
                    <div className="border border-[#141414] p-4 rounded-2xl bg-rose-50 border-rose-200">
                      <p className="text-[10px] uppercase tracking-widest text-rose-700 mb-1">Total Variance</p>
                      <p className="text-3xl font-serif italic text-rose-600">
                        {parseFloat(summary.totalVariance) < 0 ? '-' : ''}${formatAmount(Math.abs(parseFloat(summary.totalVariance)))}
                      </p>
                    </div>
                  </div>

                  {/* Row 3 */}
                  <div className="grid grid-cols-2 gap-4">
                    <div className="border border-[#141414] p-4 rounded-2xl bg-white/50 flex justify-between items-center">
                      <div>
                        <p className="text-[10px] uppercase tracking-widest opacity-50 mb-1">Sigma Total Upload</p>
                        <p className="text-xl font-serif italic">${formatAmount(summary.sigmaTotalUpload)}</p>
                      </div>
                      <div className="w-px h-10 bg-[#141414]/10 mx-4" />
                      <div>
                        <p className="text-[10px] uppercase tracking-widest opacity-50 mb-1">NetSuite Total Upload</p>
                        <p className="text-xl font-serif italic">${formatAmount(summary.netsuiteTotalUpload)}</p>
                      </div>
                    </div>
                    <div className="border border-[#141414] p-4 rounded-2xl bg-amber-50 border-amber-200">
                      <p className="text-[10px] uppercase tracking-widest text-amber-700 mb-1">Total Unmatched Items</p>
                      <p className="text-3xl font-serif italic text-amber-900">{formatInteger(summary.totalUnmatchedItems)}</p>
                    </div>
                  </div>
                </div>
              )}

              {/* Tabs */}
              <div className="flex gap-4 border-b border-[#141414]/10 pb-2 mt-8">
                {[
                  { id: 'all', label: 'All Transactions', icon: <Layers size={14} /> },
                  { id: 'matched', label: 'Matched Only', icon: <CheckCircle2 size={14} /> },
                  { id: 'mismatched', label: 'Mismatched Only', icon: <AlertCircle size={14} /> },
                  { id: 'missing', label: 'Missing Only', icon: <Filter size={14} /> }
                ].map((tab) => (
                  <button
                    key={tab.id}
                    onClick={() => setActiveTab(tab.id as any)}
                    className={`flex items-center gap-2 px-4 py-2 text-xs uppercase tracking-widest transition-all rounded-t-lg ${activeTab === tab.id ? 'bg-[#141414] text-[#E4E3E0] font-bold' : 'opacity-50 hover:opacity-100'}`}
                  >
                    {tab.icon} {tab.label}
                  </button>
                ))}
              </div>

              <div className="flex justify-between items-end">
                <div>
                  <h3 className="font-serif italic text-2xl">Reconciliation Results</h3>
                  <p className="text-xs opacity-50 font-mono">
                    Showing {Math.min(filteredResults.length, 1000)} of {filteredResults.length} filtered rows 
                    {filteredResults.length > 1000 && " (Export to Excel to see all)"}
                  </p>
                </div>
                <div className="flex gap-4">
                  <button 
                    onClick={() => downloadExport(false)}
                    className="px-6 py-3 bg-[#141414] text-[#E4E3E0] rounded-full text-xs uppercase tracking-widest hover:bg-slate-700 transition-colors flex items-center gap-2 font-bold"
                  >
                    <Download size={14} /> Full Export
                  </button>
                  <button 
                    onClick={() => downloadExport(true)}
                    className="px-6 py-3 bg-[#141414] text-[#E4E3E0] rounded-full text-xs uppercase tracking-widest hover:bg-emerald-600 transition-colors flex items-center gap-2 font-bold"
                  >
                    <Download size={14} /> Unmatched Only
                  </button>
                </div>
              </div>

              {/* Results Table v2 */}
              <div className="border border-[#141414] rounded-2xl overflow-hidden bg-white/50">
                <div className="grid grid-cols-12 p-4 border-b border-[#141414] bg-[#141414] text-[#E4E3E0] text-[9px] uppercase tracking-widest font-bold">
                  <div>Transaction ID</div>
                  <div>Status</div>
                  <div>Sigma Date</div>
                  <div>NS Date</div>
                  <div>Sigma Amt</div>
                  <div>NS Amt</div>
                  <div>Variance</div>
                  <div>Type</div>
                  <div>Card</div>
                  <div>Doc #</div>
                  <div>Account</div>
                  <div>Explanation</div>
                </div>
                <div className="max-h-[500px] overflow-y-auto">
                  {filteredResults.slice(0, 1000).map((row, i) => (
                    <div key={i} className="grid grid-cols-12 p-4 border-b border-[#141414]/10 hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors group cursor-default text-[10px]">
                      <div className="font-mono text-[9px] truncate flex flex-col">
                        <span className="truncate">{row.transaction_id}</span>
                        {(row.sigma_count && row.sigma_count > 1) || (row.netsuite_count && row.netsuite_count > 1) ? (
                          <span className="text-[8px] text-amber-600 group-hover:text-amber-400 font-bold uppercase">
                            Grouped ({row.sigma_count || 1}S / {row.netsuite_count || 1}N)
                          </span>
                        ) : null}
                      </div>
                      <div className="flex items-center">
                        {row.status === 'Matched' ? (
                          <span className="bg-emerald-100 text-emerald-700 px-1 rounded group-hover:bg-emerald-500 group-hover:text-white">Matched</span>
                        ) : row.status === 'Mismatched' ? (
                          <span className="bg-amber-100 text-amber-700 px-1 rounded group-hover:bg-amber-500 group-hover:text-white">Mismatched</span>
                        ) : row.status === 'Missing in Sigma' ? (
                          <span className="bg-blue-100 text-blue-700 px-1 rounded group-hover:bg-blue-500 group-hover:text-white">Missing</span>
                        ) : (
                          <span className="bg-rose-100 text-rose-700 px-1 rounded group-hover:bg-rose-500 group-hover:text-white">Missing</span>
                        )}
                      </div>
                      <div className="font-mono">{formatDate(row.sigma_date)}</div>
                      <div className="font-mono">{formatDate(row.netsuite_date)}</div>
                      <div className="font-mono">${formatAmount(row.sigma_amount)}</div>
                      <div className="font-mono">${formatAmount(row.netsuite_amount)}</div>
                      <div className={`font-mono ${row.variance && Math.abs(row.variance) > 0.01 ? 'text-rose-600 group-hover:text-rose-300' : ''}`}>
                        {row.variance !== undefined ? `$${formatAmount(row.variance)}` : '-'}
                      </div>
                      <div className="opacity-60 group-hover:opacity-100 truncate">{String(row.account_type || '')}</div>
                      <div className="opacity-60 group-hover:opacity-100 truncate">{String(row.card_type || '')}</div>
                      <div className="opacity-60 group-hover:opacity-100 truncate">{String(row.document_number || '')}</div>
                      <div className="opacity-60 group-hover:opacity-100 truncate">{String(row.account || '')}</div>
                      <div className="opacity-60 group-hover:opacity-100 italic line-clamp-2 leading-tight">{row.explanation}</div>
                    </div>
                  ))}
                </div>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* Footer */}
      <footer className="mt-20 border-t border-[#141414]/10 p-8 text-center opacity-30 text-[10px] uppercase tracking-[0.2em]">
        Data Match Tool &copy; 2026 &bull; Secure Reconciliation Engine
      </footer>
    </div>
  );
}
