import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';

const EIRIncomeProjectionTool = () => {
  const [file, setFile] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [message, setMessage] = useState({ type: '', text: '' });
  const [progress, setProgress] = useState({ current: 0, total: 0 });
  const [dragActive, setDragActive] = useState(false);
  const [results, setResults] = useState(null);
  const [productType, setProductType] = useState('Other Products'); // 'Other Products' or 'Tractor'

  const logoUrl = "./tvscredit-logo.jfif";

  // ==================== CALCULATION FUNCTIONS ====================
  const PMT = (rate, nper, pv) => {
    if (rate === 0) return pv / nper;
    const pvif = Math.pow(1 + rate, nper);
    return (rate * pv * pvif) / (pvif - 1);
  };

  const daysBetween = (date1, date2) => {
    const d1 = new Date(date1);
    const d2 = new Date(date2);
    return Math.round(Math.abs((d2.getTime() - d1.getTime()) / (1000 * 60 * 60 * 24)));
  };

  const addMonths = (date, months) => {
    const result = new Date(date);
    result.setMonth(result.getMonth() + months);
    return result;
  };

  const calculateFirstEMIDate = (disbDate, productType = 'Other Products') => {
    const day = disbDate.getDate();
    const emiDay = productType === 'Tractor' ? 5 : 3;
    
    if (day < 20) {
      // Pre-20: First EMI is next month
      return new Date(disbDate.getFullYear(), disbDate.getMonth() + 1, emiDay);
    } else {
      // Post-20: First EMI is month after next
      return new Date(disbDate.getFullYear(), disbDate.getMonth() + 2, emiDay);
    }
  };

  const formatMonthYear = (date) => {
    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return `${monthNames[date.getMonth()]}-${date.getFullYear().toString().slice(-2)}`;
  };

  const formatDate = (date) => {
    const day = date.getDate().toString().padStart(2, '0');
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    return `${day}/${month}/${date.getFullYear()}`;
  };

  // Get months per installment based on frequency
  const getMonthsPerInstallment = (frequency) => {
    switch (frequency) {
      case 'Bimonthly': return 2;
      case 'Quarterly': return 3;
      case 'Halfyearly': return 6;
      default: return 1; // Monthly
    }
  };

  const calculateXIRR = (cashflows, dates) => {
    let rate = 0.1;
    for (let iter = 0; iter < 100; iter++) {
      let npv = 0, dnpv = 0;
      const baseDate = dates[0];
      for (let j = 0; j < cashflows.length; j++) {
        const days = daysBetween(baseDate, dates[j]);
        const years = days / 365.0;
        const factor = Math.pow(1 + rate, years);
        npv += cashflows[j] / factor;
        dnpv -= years * cashflows[j] / (factor * (1 + rate));
      }
      if (Math.abs(npv) < 1e-6) break;
      if (Math.abs(dnpv) < 1e-12) break;
      const newRate = rate - npv / dnpv;
      const change = Math.max(-0.5, Math.min(0.5, newRate - rate));
      rate = rate + change;
      if (Math.abs(change) < 1e-6) break;
    }
    return rate;
  };

  const parseExcelDate = (value) => {
    if (value instanceof Date) return value;
    if (typeof value === 'number') {
      return new Date(new Date(1899, 11, 30).getTime() + value * 86400000);
    }
    if (typeof value === 'string') {
      const parsed = new Date(value);
      if (!isNaN(parsed.getTime())) return parsed;
      const parts = value.split(/[\/\-]/);
      if (parts.length === 3) {
        return new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
      }
    }
    return null;
  };

  const parseAmortIRR = (value) => {
    if (value === null || value === undefined) return 0;
    const numValue = parseFloat(value);
    return numValue < 1 ? numValue * 100 : numValue;
  };

  const calculateEIR = (loanCase) => {
    const agreementNumber = loanCase.agreementNumber;
    const amountFinanced = parseFloat(loanCase.amountFinanced);
    const tenureMonths = parseInt(loanCase.tenure); // Tenure in months
    const productType = loanCase.productType || 'Other Products';
    const repaymentFrequency = loanCase.repaymentFrequency || 'Monthly';
    const disbursementDate = parseExcelDate(loanCase.disbursementDate);
    const amortIRR = parseAmortIRR(loanCase.amortIRR);
    const upfrontIncome = parseFloat(loanCase.upfrontIncome) || 0;
    const upfrontExpense = parseFloat(loanCase.upfrontExpense) || 0;
    const advanceEMI = parseFloat(loanCase.advanceEMI) || 0;

    if (!disbursementDate || isNaN(disbursementDate.getTime())) {
      throw new Error(`Invalid disbursement date for agreement ${agreementNumber}`);
    }

    // Calculate installment parameters based on frequency
    const monthsPerInstallment = getMonthsPerInstallment(repaymentFrequency);
    const numberOfInstallments = Math.round(tenureMonths / monthsPerInstallment);

    const amortRate = amortIRR / 100;
    const periodicRate = amortRate * monthsPerInstallment / 12; // Rate per installment period
    const installmentAmount = PMT(periodicRate, numberOfInstallments, amountFinanced);
    const firstEMIDate = calculateFirstEMIDate(disbursementDate, productType);

    const igaapTable = [];
    igaapTable.push({ month: 0, emiDate: disbursementDate, emi: -(amountFinanced - advanceEMI), closingBalance: 0, days: null });

    const disbursementMonthEnd = new Date(disbursementDate.getFullYear(), disbursementDate.getMonth() + 1, 0);
    const daysToMonthEnd = daysBetween(disbursementDate, disbursementMonthEnd);
    let openingBalance = amountFinanced - advanceEMI;
    const interestPartial = openingBalance * amortRate * (daysToMonthEnd / 365);

    igaapTable.push({ month: 1, emiDate: disbursementMonthEnd, emi: 0, closingBalance: openingBalance + interestPartial, days: daysToMonthEnd });

    let currentDate = disbursementDate;
    openingBalance = amountFinanced - advanceEMI;
    let installmentCount = 0;
    let currentEMIDate = firstEMIDate;

    while (installmentCount < numberOfInstallments && installmentCount < 100) {
      const days = daysBetween(currentDate, currentEMIDate);
      const interest = openingBalance * amortRate * (days / 365);
      const cashflow = (installmentCount + 1 < numberOfInstallments) ? installmentAmount : openingBalance + interest;
      const principal = cashflow - interest;
      let closingBalance = openingBalance - principal;
      
      // Ensure closing balance doesn't go negative due to rounding
      if (Math.abs(closingBalance) < 0.0001) closingBalance = 0;

      igaapTable.push({ month: installmentCount + 2, emiDate: currentEMIDate, openingBalance, emi: cashflow, closingBalance, days });

      currentDate = currentEMIDate;
      openingBalance = closingBalance;
      installmentCount++;
      currentEMIDate = addMonths(currentEMIDate, monthsPerInstallment); // Add months based on frequency
      if (closingBalance <= 0) break;
    }

    const lastEMIDate = igaapTable[igaapTable.length - 1].emiDate;

    // Calculate RLV (Revised Loan Value)
    const revisedLoanValue = amountFinanced - upfrontIncome + upfrontExpense - advanceEMI;
    
    // Get EMI schedule from Step A (IGAAP) for Goal Seek
    const stepAEMIs = igaapTable.filter(row => row.month >= 2).map(row => ({
      emi: row.emi,
      days: row.days,
      date: row.emiDate
    }));
    
    // Goal Seek function: calculates Step B final closing balance for a given rate
    const getStepBFinalClosing = (xirrNominalRate) => {
      let openingB = revisedLoanValue;
      
      for (let i = 0; i < stepAEMIs.length; i++) {
        const row = stepAEMIs[i];
        const interestB = openingB * xirrNominalRate * (row.days / 365);
        const closingB = openingB + interestB - row.emi;
        
        if (i + 1 === stepAEMIs.length) {
          return closingB;
        }
        
        openingB = closingB;
      }
      return 0;
    };
    
    // Binary search Goal Seek to find rate where closing = 0
    let goalSeekLow = 0.10;
    let goalSeekHigh = 0.60;
    let xirrNominal = 0;
    
    for (let iter = 0; iter < 100; iter++) {
      const mid = (goalSeekLow + goalSeekHigh) / 2;
      const finalClosing = getStepBFinalClosing(mid);
      
      if (Math.abs(finalClosing) < 1e-8) {
        xirrNominal = mid;
        break;
      }
      
      if (finalClosing > 0) {
        goalSeekHigh = mid;
      } else {
        goalSeekLow = mid;
      }
      
      if (iter === 99) {
        xirrNominal = mid;
      }
    }

    const stepBTable = [];
    stepBTable.push({ month: 0, emiDate: disbursementDate, emi: -revisedLoanValue, closingBalance: 0, days: null });

    let openingBalanceB = revisedLoanValue;
    const interestPartialB = openingBalanceB * xirrNominal * (daysToMonthEnd / 365);
    stepBTable.push({ month: 1, emiDate: disbursementMonthEnd, emi: 0, closingBalance: openingBalanceB + interestPartialB, days: daysToMonthEnd });

    openingBalanceB = revisedLoanValue;
    for (let i = 2; i < igaapTable.length; i++) {
      const row = igaapTable[i];
      const interestB = openingBalanceB * xirrNominal * (row.days / 365);
      const closingBalanceB = openingBalanceB + interestB - row.emi;
      stepBTable.push({ month: row.month, emiDate: row.emiDate, openingBalance: openingBalanceB, emi: row.emi, closingBalance: closingBalanceB, days: row.days });
      openingBalanceB = closingBalanceB;
    }

    const stepCTable = [];
    let disbursementUnamortized = null;
    for (let i = 0; i < Math.min(igaapTable.length, stepBTable.length); i++) {
      const stepARow = igaapTable[i];
      const stepBRow = stepBTable[i];
      let unamortized, eirIncome = null;

      if (i === 0) {
        unamortized = -(stepARow.emi - stepBRow.emi);
        disbursementUnamortized = unamortized;
      } else {
        unamortized = stepARow.closingBalance - stepBRow.closingBalance;
      }

      if (i === 0 || i === 1) eirIncome = null;
      else if (i === 2) eirIncome = disbursementUnamortized - unamortized;
      else eirIncome = stepCTable[i - 1].unamortized - unamortized;

      stepCTable.push({ month: stepARow.month, emiDate: stepARow.emiDate, unamortized, eirIncome, days: stepARow.days });
    }

    const stepDTable = [];
    const is_pre_20 = disbursementDate.getDate() < 20;
    const emiDay = productType === 'Tractor' ? 5 : 3;
    
    const emiRows = stepCTable
      .map((row, idx) => ({ ...row, idx }))
      .filter(row => row.eirIncome !== null)
      .map(row => ({
        idx: row.idx,
        emi_date: row.emiDate,
        eir_income: row.eirIncome,
        prev_date: row.idx === 2 ? disbursementDate : stepCTable[row.idx - 1].emiDate,
        days: row.days
      }));

    if (emiRows.length > 0) {
      if (repaymentFrequency === 'Monthly') {
        // Monthly: Show all calendar months from disbursement with col_a, col_b, col_c
        let monthIdx = 0;
        let currentMonth = new Date(disbursementDate.getFullYear(), disbursementDate.getMonth(), 1);
        const endMonth = new Date(lastEMIDate.getFullYear(), lastEMIDate.getMonth() + 1, 1);
        
        while (currentMonth < endMonth && monthIdx < 100) {
          const monthEnd = new Date(currentMonth.getFullYear(), currentMonth.getMonth() + 1, 0);
          const daysInCurrentMonth = monthEnd.getDate();
          let col_b = 0;
          
          // Check if there's an EMI in this month (for Step C EIR Income)
          const emiInThisMonth = emiRows.find(emi => 
            emi.emi_date.getFullYear() === currentMonth.getFullYear() && 
            emi.emi_date.getMonth() === currentMonth.getMonth()
          );
          const stepC_eirIncome = emiInThisMonth ? emiInThisMonth.eir_income : null;
          
          if (monthIdx === 0) {
            const firstEmi = emiRows[0];
            const daysFromDisbToMonthEnd = daysBetween(disbursementDate, monthEnd);
            col_b = (firstEmi.eir_income / firstEmi.days) * daysFromDisbToMonthEnd;
          } else if (monthIdx === 1 && !is_pre_20) {
            const firstEmi = emiRows[0];
            col_b = (firstEmi.eir_income / firstEmi.days) * daysInCurrentMonth;
          } else {
            const nextMonth = new Date(currentMonth.getFullYear(), currentMonth.getMonth() + 1, 1);
            const foundEmi = emiRows.find(emi => 
              emi.emi_date.getFullYear() === nextMonth.getFullYear() && 
              emi.emi_date.getMonth() === nextMonth.getMonth()
            );
            if (foundEmi) {
              col_b = (foundEmi.eir_income / foundEmi.days) * (foundEmi.days - emiDay);
            }
          }
          
          stepDTable.push({
            monthIdx,
            month: formatMonthYear(currentMonth),
            stepC_eirIncome,
            col_a: 0,
            col_b,
            col_c: 0
          });
          
          monthIdx++;
          currentMonth = new Date(currentMonth.getFullYear(), currentMonth.getMonth() + 1, 1);
        }
        
        // Calculate Col A values
        for (let monthIdx = 0; monthIdx < stepDTable.length; monthIdx++) {
          if (monthIdx > 0) {
            const currentMonthDate = new Date(disbursementDate.getFullYear(), disbursementDate.getMonth() + monthIdx, 1);
            const emiOnDay = emiRows.find(emi => 
              emi.emi_date.getFullYear() === currentMonthDate.getFullYear() && 
              emi.emi_date.getMonth() === currentMonthDate.getMonth()
            );
            
            if (emiOnDay) {
              const isFirstEmi = (emiOnDay.idx === 2);
              if (!is_pre_20 && isFirstEmi) {
                let sum = 0;
                for (let i = 0; i < monthIdx; i++) sum += stepDTable[i].col_b;
                stepDTable[monthIdx].col_a = emiOnDay.eir_income - sum;
              } else {
                stepDTable[monthIdx].col_a = emiOnDay.eir_income - stepDTable[monthIdx - 1].col_b;
              }
            }
          }
        }
        
        // Calculate Col C
        for (let i = 0; i < stepDTable.length; i++) {
          stepDTable[i].col_c = stepDTable[i].col_a + stepDTable[i].col_b;
        }
      } else {
        // Non-Monthly (Tractor): Show disbursement month + intermediate months (for post-20) + installment months
        let monthIdx = 0;
        let currentMonth = new Date(disbursementDate.getFullYear(), disbursementDate.getMonth(), 1);
        
        const firstEmi = emiRows[0];
        const firstEmiMonth = new Date(firstEmi.emi_date.getFullYear(), firstEmi.emi_date.getMonth(), 1);
        
        // First month (disbursement month) - partial month calculation, no EMI so stepC_eirIncome = null
        const firstMonthEnd = new Date(currentMonth.getFullYear(), currentMonth.getMonth() + 1, 0);
        const daysFromDisbToMonthEnd = daysBetween(disbursementDate, firstMonthEnd);
        let dailyRate = firstEmi.eir_income / firstEmi.days;
        let col_b = dailyRate * daysFromDisbToMonthEnd;
        
        stepDTable.push({
          monthIdx: monthIdx++,
          month: formatMonthYear(currentMonth),
          stepC_eirIncome: null,
          col_a: 0,
          col_b: col_b,
          col_c: col_b
        });
        
        let prevColB = col_b;
        
        // For Post-20: Add intermediate months between disbursement and first EMI
        if (!is_pre_20) {
          currentMonth = new Date(currentMonth.getFullYear(), currentMonth.getMonth() + 1, 1);
          
          // Add all months between disbursement month and first EMI month
          while (currentMonth < firstEmiMonth) {
            const monthEnd = new Date(currentMonth.getFullYear(), currentMonth.getMonth() + 1, 0);
            const daysInMonth = monthEnd.getDate();
            col_b = dailyRate * daysInMonth;
            
            stepDTable.push({
              monthIdx: monthIdx++,
              month: formatMonthYear(currentMonth),
              stepC_eirIncome: null,
              col_a: 0,
              col_b: col_b,
              col_c: col_b
            });
            
            prevColB += col_b;
            currentMonth = new Date(currentMonth.getFullYear(), currentMonth.getMonth() + 1, 1);
          }
        }
        
        // For each installment - these have Step C EIR Income
        for (let i = 0; i < emiRows.length; i++) {
          const emiRow = emiRows[i];
          const emiMonth = new Date(emiRow.emi_date.getFullYear(), emiRow.emi_date.getMonth(), 1);
          
          // Col A = EIR Income from Step C - previous Col B (accumulated for post-20 first EMI)
          const col_a = emiRow.eir_income - prevColB;
          
          // Col B = days after EMI in this period prorated
          col_b = 0;
          if (i < emiRows.length - 1) {
            const nextEmi = emiRows[i + 1];
            col_b = (nextEmi.eir_income / nextEmi.days) * (nextEmi.days - emiDay);
          }
          
          const col_c = col_a + col_b;
          
          stepDTable.push({
            monthIdx: monthIdx++,
            month: formatMonthYear(emiMonth),
            stepC_eirIncome: emiRow.eir_income,
            col_a,
            col_b,
            col_c
          });
          
          prevColB = col_b;
        }
      }
    }

    const finalOutputTable = stepDTable.map(row => ({
      agreementNumber,
      productType,
      amountFinanced,
      tenure: tenureMonths,
      repaymentFrequency,
      numberOfInstallments,
      disbursement: formatDate(disbursementDate),
      firstEMI: formatDate(firstEMIDate),
      lastEMI: formatDate(lastEMIDate),
      amortIRR,
      advanceEMI,
      upfrontIncome,
      upfrontExpense,
      month: row.month,
      stepC_eirIncome: row.stepC_eirIncome !== null ? parseFloat(row.stepC_eirIncome.toFixed(4)) : '',
      eirIncome: parseFloat(row.col_c.toFixed(4))
    }));

    return finalOutputTable;
  };

  const handleDrag = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  }, []);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const droppedFile = e.dataTransfer.files[0];
      if (droppedFile.name.match(/\.(xlsx|xls)$/i)) {
        setFile(droppedFile);
        setMessage({ type: '', text: '' });
        setResults(null);
      } else {
        setMessage({ type: 'error', text: 'Please upload only Excel files (.xlsx or .xls)' });
      }
    }
  }, []);

  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0];
    if (selectedFile) {
      setFile(selectedFile);
      setMessage({ type: '', text: '' });
      setResults(null);
    }
  };

  const removeFile = () => {
    setFile(null);
    setResults(null);
    setMessage({ type: '', text: '' });
  };

  const processFile = async () => {
    if (!file) {
      setMessage({ type: 'error', text: 'Please select an Excel file first' });
      return;
    }
    setIsProcessing(true);
    setMessage({ type: '', text: '' });

    try {
      const arrayBuffer = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = () => reject(new Error('Failed to read file'));
        reader.readAsArrayBuffer(file);
      });

      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);

      if (jsonData.length === 0) throw new Error('No data found in the Excel file');
      setProgress({ current: 0, total: jsonData.length });

      const loanCases = jsonData.map((row, index) => ({
        agreementNumber: row['Agreement Number'] || row['agreement_number'] || row['AgreementNumber'] || `LOAN_${index + 1}`,
        productType: row['Product Type'] || row['product_type'] || row['ProductType'] || productType, // Use dropdown value as default
        repaymentFrequency: row['Repayment Frequency'] || row['repayment_frequency'] || row['RepaymentFrequency'] || 'Monthly',
        amountFinanced: row['Amount Financed'] || row['amount_financed'] || row['AmountFinanced'],
        tenure: row['Tenure'] || row['tenure'],
        disbursementDate: row['Disbursement Date'] || row['disbursement_date'] || row['DisbursementDate'] || row['Disbursement'],
        amortIRR: row['Amort IRR'] || row['amort_irr'] || row['AmortIRR'] || row['Amort IRR (%)'],
        advanceEMI: row['Advance EMI'] || row['advance_emi'] || row['AdvanceEMI'] || 0,
        upfrontIncome: row['Upfront Income'] || row['upfront_income'] || row['UpfrontIncome'] || 0,
        upfrontExpense: row['Upfront Expense'] || row['upfront_expense'] || row['UpfrontExpense'] || 0
      }));

      const outputWorkbook = XLSX.utils.book_new();
      const headers = ['Agreement Number', 'Product Type', 'Amount Financed', 'Tenure (Months)', 'Repayment Frequency', 'No. of Installments', 'Disbursement', 'First EMI', 'Last EMI', 'Amort IRR (%)', 'Advance EMI', 'Upfront Income', 'Upfront Expense', 'Month', 'Step C EIR Income', 'EIR Income'];
      const errors = [];
      const allResults = [];

      // Helper function to sanitize cell values
      const sanitizeValue = (val) => {
        if (val === null || val === undefined) return '';
        if (typeof val === 'number' && !isFinite(val)) return 0;
        return val;
      };

      // Helper function to create safe sheet name
      const createSafeSheetName = (name, existingNames) => {
        // Remove invalid characters and limit length
        let safeName = String(name || 'Sheet')
          .replace(/[\\\/\?\*\[\]:\'\"]/g, '_')
          .replace(/[\x00-\x1F\x7F]/g, '') // Remove control characters
          .trim()
          .substring(0, 31);
        
        if (!safeName) safeName = 'Sheet';
        
        // Ensure uniqueness
        let uniqueName = safeName;
        let counter = 1;
        while (existingNames.includes(uniqueName)) {
          const suffix = `_${counter}`;
          uniqueName = safeName.substring(0, 31 - suffix.length) + suffix;
          counter++;
        }
        return uniqueName;
      };

      for (let i = 0; i < loanCases.length; i++) {
        const loanCase = loanCases[i];
        setProgress({ current: i + 1, total: loanCases.length });

        try {
          const finalOutputTable = calculateEIR(loanCase);
          allResults.push({ agreementNumber: loanCase.agreementNumber, data: finalOutputTable });

          const wsData = [headers, ...finalOutputTable.map(row => [
            sanitizeValue(row.agreementNumber),
            sanitizeValue(row.productType),
            sanitizeValue(row.amountFinanced),
            sanitizeValue(row.tenure),
            sanitizeValue(row.repaymentFrequency),
            sanitizeValue(row.numberOfInstallments),
            sanitizeValue(row.disbursement),
            sanitizeValue(row.firstEMI),
            sanitizeValue(row.lastEMI),
            sanitizeValue(row.amortIRR),
            sanitizeValue(row.advanceEMI),
            sanitizeValue(row.upfrontIncome),
            sanitizeValue(row.upfrontExpense),
            sanitizeValue(row.month),
            sanitizeValue(row.stepC_eirIncome),
            sanitizeValue(row.eirIncome)
          ])];
          const ws = XLSX.utils.aoa_to_sheet(wsData);
          ws['!cols'] = [{ wch: 18 }, { wch: 16 }, { wch: 16 }, { wch: 14 }, { wch: 20 }, { wch: 16 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 10 }, { wch: 16 }, { wch: 12 }];

          const sheetName = createSafeSheetName(loanCase.agreementNumber, outputWorkbook.SheetNames);
          XLSX.utils.book_append_sheet(outputWorkbook, ws, sheetName);
        } catch (error) {
          errors.push(`${loanCase.agreementNumber}: ${error.message}`);
        }
      }

      if (allResults.length > 0) {
        const consolidatedData = [];
        for (let i = 0; i < allResults.length; i++) {
          consolidatedData.push(headers);
          allResults[i].data.forEach(row => consolidatedData.push([
            sanitizeValue(row.agreementNumber),
            sanitizeValue(row.productType),
            sanitizeValue(row.amountFinanced),
            sanitizeValue(row.tenure),
            sanitizeValue(row.repaymentFrequency),
            sanitizeValue(row.numberOfInstallments),
            sanitizeValue(row.disbursement),
            sanitizeValue(row.firstEMI),
            sanitizeValue(row.lastEMI),
            sanitizeValue(row.amortIRR),
            sanitizeValue(row.advanceEMI),
            sanitizeValue(row.upfrontIncome),
            sanitizeValue(row.upfrontExpense),
            sanitizeValue(row.month),
            sanitizeValue(row.stepC_eirIncome),
            sanitizeValue(row.eirIncome)
          ]));
          if (i < allResults.length - 1) consolidatedData.push([]);
        }
        const consolidatedWs = XLSX.utils.aoa_to_sheet(consolidatedData);
        consolidatedWs['!cols'] = [{ wch: 18 }, { wch: 16 }, { wch: 16 }, { wch: 14 }, { wch: 20 }, { wch: 16 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 10 }, { wch: 16 }, { wch: 12 }];

        // Properly construct the workbook with All Cases first
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, consolidatedWs, 'All Cases');
        
        // Add individual sheets
        outputWorkbook.SheetNames.forEach(name => {
          XLSX.utils.book_append_sheet(newWorkbook, outputWorkbook.Sheets[name], name);
        });
        
        const fileName = `EIR_Projection_${new Date().toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(newWorkbook, fileName);

        setResults({
          total: loanCases.length,
          success: loanCases.length - errors.length,
          failed: errors.length,
          fileName
        });
      } else {
        const fileName = `EIR_Projection_${new Date().toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(outputWorkbook, fileName);

        setResults({
          total: loanCases.length,
          success: loanCases.length - errors.length,
          failed: errors.length,
          fileName
        });
      }

      if (errors.length > 0) {
        setMessage({ type: 'warning', text: `Processed ${loanCases.length - errors.length}/${loanCases.length} loans.\n\nErrors:\n${errors.join('\n')}` });
      } else {
        setMessage({ type: 'success', text: `Successfully processed ${loanCases.length} loan cases!` });
      }
    } catch (error) {
      console.error('Error:', error);
      setMessage({ type: 'error', text: `Error: ${error.message}` });
    } finally {
      setIsProcessing(false);
      setProgress({ current: 0, total: 0 });
    }
  };

  const downloadTemplate = () => {
    const templateData = [
      ['Agreement Number', 'Product Type', 'Repayment Frequency', 'Amount Financed', 'Tenure', 'Disbursement Date', 'Amort IRR', 'Advance EMI', 'Upfront Income', 'Upfront Expense'],
      ['AGR001', 'Other Products', 'Monthly', 100000, 36, '15/01/2025', 12.5, 0, 2500, 500],
      ['AGR002', 'Tractor', 'Quarterly', 150000, 36, '20/02/2025', 11.75, 0, 3200, 750],
      ['AGR003', 'Tractor', 'Halfyearly', 200000, 36, '10/03/2025', 12.0, 0, 4000, 800]
    ];
    const ws = XLSX.utils.aoa_to_sheet(templateData);
    ws['!cols'] = [{ wch: 18 }, { wch: 16 }, { wch: 20 }, { wch: 16 }, { wch: 10 }, { wch: 18 }, { wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 16 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Template');
    XLSX.writeFile(wb, 'EIR_Input_Template.xlsx');
  };

  const columns = [
    { name: 'Agreement Number', type: 'Text', required: true },
    { name: 'Product Type', type: 'Text', required: false, note: 'Other Products / Tractor' },
    { name: 'Repayment Frequency', type: 'Text', required: false, note: 'Monthly / Bimonthly / Quarterly / Halfyearly' },
    { name: 'Amount Financed', type: 'Number', required: true },
    { name: 'Tenure', type: 'Number', required: true, note: 'In months' },
    { name: 'Disbursement Date', type: 'Date', required: true },
    { name: 'Amort IRR', type: '%', required: true },
    { name: 'Advance EMI', type: 'Number', required: false },
    { name: 'Upfront Income', type: 'Number', required: false },
    { name: 'Upfront Expense', type: 'Number', required: false }
  ];

  return (
    <div className="min-h-screen bg-slate-50">
      {/* Modern Header */}
      <header className="bg-white border-b border-slate-200">
        <div className="max-w-7xl mx-auto">
          <div className="flex items-center justify-between px-6 py-4">
            <div className="flex items-center">
              <img src={logoUrl} alt="TVS Credit" className="h-12 w-auto" />
            </div>
            <div className="hidden md:flex items-center gap-2 text-sm text-slate-500">
              <span className="w-2 h-2 bg-green-500 rounded-full animate-pulse"></span>
              Finance Portal
            </div>
          </div>
        </div>
      </header>

      {/* Hero Section */}
      <div style={{ background: 'linear-gradient(135deg, #00703c 0%, #009950 50%, #00703c 100%)' }}>
        <div className="max-w-7xl mx-auto px-6 py-16">
          <div className="text-center">
            <h1 className="text-4xl md:text-5xl font-light text-white mb-4 tracking-tight">
              EIR Income Projection
            </h1>
            <p className="text-green-100 text-lg font-light max-w-2xl mx-auto">
              Generate accurate Effective Interest Rate projections for your loan portfolio
            </p>
          </div>
        </div>
        {/* Curved bottom */}
        <div className="h-16 bg-slate-50" style={{ borderRadius: '100% 100% 0 0', marginTop: '-1px' }}></div>
      </div>

      {/* Main Content */}
      <main className="max-w-6xl mx-auto px-6 -mt-8">
        <div className="grid grid-cols-1 lg:grid-cols-5 gap-8">
          
          {/* Upload Card - Main */}
          <div className="lg:col-span-3">
            <div className="bg-white rounded-2xl shadow-xl shadow-slate-200/50 overflow-hidden">
              {/* Card Header */}
              <div className="px-8 py-6 border-b border-slate-100">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 rounded-xl flex items-center justify-center" style={{ backgroundColor: '#e8f5e9' }}>
                    <svg className="w-5 h-5" style={{ color: '#00703c' }} fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                    </svg>
                  </div>
                  <div>
                    <h2 className="text-lg font-semibold text-slate-800">Upload Loan Data</h2>
                    <p className="text-sm text-slate-500">Excel files (.xlsx, .xls)</p>
                  </div>
                </div>
              </div>

              <div className="p-8">
                {/* Drop Zone */}
                <div
                  onDragEnter={handleDrag}
                  onDragLeave={handleDrag}
                  onDragOver={handleDrag}
                  onDrop={handleDrop}
                  className={`relative rounded-2xl border-2 border-dashed transition-all duration-300 ${
                    dragActive 
                      ? 'border-green-500 bg-green-50' 
                      : file 
                        ? 'border-green-400 bg-green-50/50' 
                        : 'border-slate-200 hover:border-slate-300 bg-slate-50/50'
                  }`}
                >
                  {file ? (
                    <div className="p-10 text-center">
                      <div className="w-16 h-16 mx-auto mb-4 rounded-2xl flex items-center justify-center" style={{ backgroundColor: '#dcfce7' }}>
                        <svg className="w-8 h-8 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                        </svg>
                      </div>
                      <p className="text-lg font-medium text-slate-800 mb-1">{file.name}</p>
                      <p className="text-sm text-slate-500 mb-4">{(file.size / 1024).toFixed(1)} KB</p>
                      <button
                        onClick={removeFile}
                        className="text-sm text-red-500 hover:text-red-600 font-medium inline-flex items-center gap-1.5 px-4 py-2 rounded-lg hover:bg-red-50 transition-colors"
                      >
                        <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                        </svg>
                        Remove
                      </button>
                    </div>
                  ) : (
                    <label className="block p-10 cursor-pointer">
                      <div className="text-center">
                        <div className="w-16 h-16 mx-auto mb-4 rounded-2xl bg-slate-100 flex items-center justify-center">
                          <svg className="w-8 h-8 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                          </svg>
                        </div>
                        <p className="text-lg font-medium text-slate-700 mb-1">Drop your file here</p>
                        <p className="text-sm text-slate-500 mb-4">or click to browse</p>
                        <span className="inline-flex items-center gap-2 px-5 py-2.5 rounded-xl text-white text-sm font-medium transition-all hover:opacity-90" style={{ backgroundColor: '#00703c' }}>
                          <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
                          </svg>
                          Select File
                        </span>
                      </div>
                      <input type="file" className="hidden" accept=".xlsx,.xls" onChange={handleFileChange} />
                    </label>
                  )}
                </div>

                {/* Progress */}
                {isProcessing && progress.total > 0 && (
                  <div className="mt-6 p-4 bg-slate-50 rounded-xl">
                    <div className="flex justify-between text-sm mb-2">
                      <span className="text-slate-600">Processing loans...</span>
                      <span className="font-semibold" style={{ color: '#00703c' }}>{progress.current}/{progress.total}</span>
                    </div>
                    <div className="h-2 bg-slate-200 rounded-full overflow-hidden">
                      <div
                        className="h-full rounded-full transition-all duration-300"
                        style={{ 
                          width: `${(progress.current / progress.total) * 100}%`,
                          backgroundColor: '#00703c'
                        }}
                      />
                    </div>
                  </div>
                )}

                {/* Product Type Selection */}
                <div className="mt-6 p-4 bg-blue-50 rounded-xl border border-blue-200">
                  <div className="flex items-center justify-between">
                    <div>
                      <label className="block text-sm font-semibold text-blue-800 mb-1">Product Type</label>
                      <p className="text-xs text-blue-600">Select default product type (can be overridden by Excel column)</p>
                    </div>
                    <select
                      value={productType}
                      onChange={(e) => setProductType(e.target.value)}
                      className="px-4 py-2 border border-blue-300 rounded-lg bg-white text-blue-800 font-medium focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                    >
                      <option value="Other Products">Other Products (EMI on 3rd)</option>
                      <option value="Tractor">Tractor (EMI on 5th)</option>
                    </select>
                  </div>
                  <div className="mt-2 text-xs text-blue-500">
                    <strong>Note:</strong> For Tractor, Repayment Frequency (Monthly/Bimonthly/Quarterly/Halfyearly) is read from Excel's "Repayment Frequency" column.
                  </div>
                </div>

                {/* Actions */}
                <div className="mt-6 flex gap-3">
                  <button
                    onClick={processFile}
                    disabled={isProcessing || !file}
                    className="flex-1 py-4 px-6 rounded-xl text-white font-semibold transition-all duration-200 flex items-center justify-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed hover:shadow-lg"
                    style={{ backgroundColor: isProcessing || !file ? '#94a3b8' : '#00703c' }}
                  >
                    {isProcessing ? (
                      <>
                        <svg className="animate-spin w-5 h-5" viewBox="0 0 24 24">
                          <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" fill="none"/>
                          <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"/>
                        </svg>
                        Processing...
                      </>
                    ) : (
                      <>
                        <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 10V3L4 14h7v7l9-11h-7z" />
                        </svg>
                        Generate Projection
                      </>
                    )}
                  </button>
                  
                  <button
                    onClick={downloadTemplate}
                    className="py-4 px-5 rounded-xl font-medium border-2 transition-all hover:bg-slate-50 flex items-center gap-2"
                    style={{ borderColor: '#00703c', color: '#00703c' }}
                  >
                    <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                    </svg>
                    <span className="hidden sm:inline">Template</span>
                  </button>
                </div>

                {/* Messages */}
                {message.text && (
                  <div className={`mt-6 p-4 rounded-xl flex items-start gap-3 ${
                    message.type === 'success' ? 'bg-green-50 text-green-800' :
                    message.type === 'warning' ? 'bg-amber-50 text-amber-800' :
                    'bg-red-50 text-red-800'
                  }`}>
                    {message.type === 'success' && (
                      <svg className="w-5 h-5 mt-0.5 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                      </svg>
                    )}
                    {message.type === 'warning' && (
                      <svg className="w-5 h-5 mt-0.5 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
                      </svg>
                    )}
                    {message.type === 'error' && (
                      <svg className="w-5 h-5 mt-0.5 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M10 14l2-2m0 0l2-2m-2 2l-2-2m2 2l2 2m7-2a9 9 0 11-18 0 9 9 0 0118 0z" />
                      </svg>
                    )}
                    <p className="text-sm whitespace-pre-wrap">{message.text}</p>
                  </div>
                )}

                {/* Results */}
                {results && (
                  <div className="mt-6 p-6 rounded-xl bg-gradient-to-br from-green-50 to-emerald-50 border border-green-100">
                    <div className="flex items-center gap-3 mb-4">
                      <div className="w-10 h-10 rounded-xl bg-green-500 flex items-center justify-center">
                        <svg className="w-5 h-5 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M5 13l4 4L19 7" />
                        </svg>
                      </div>
                      <div>
                        <h3 className="font-semibold text-slate-800">Processing Complete</h3>
                        <p className="text-sm text-slate-500">{results.fileName}</p>
                      </div>
                    </div>
                    <div className="grid grid-cols-3 gap-3">
                      <div className="bg-white rounded-xl p-4 text-center shadow-sm">
                        <p className="text-2xl font-bold text-slate-800">{results.total}</p>
                        <p className="text-xs text-slate-500 mt-1">Total</p>
                      </div>
                      <div className="bg-white rounded-xl p-4 text-center shadow-sm">
                        <p className="text-2xl font-bold text-green-600">{results.success}</p>
                        <p className="text-xs text-slate-500 mt-1">Success</p>
                      </div>
                      <div className="bg-white rounded-xl p-4 text-center shadow-sm">
                        <p className="text-2xl font-bold text-red-500">{results.failed}</p>
                        <p className="text-xs text-slate-500 mt-1">Failed</p>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* Sidebar */}
          <div className="lg:col-span-2 space-y-6">
            {/* Required Columns */}
            <div className="bg-white rounded-2xl shadow-xl shadow-slate-200/50 overflow-hidden">
              <div className="px-6 py-5 border-b border-slate-100">
                <h3 className="font-semibold text-slate-800 flex items-center gap-2">
                  <svg className="w-5 h-5 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2" />
                  </svg>
                  Input Columns
                </h3>
              </div>
              <div className="p-4">
                <div className="space-y-1">
                  {columns.map((col, idx) => (
                    <div key={idx} className="flex items-center justify-between py-2.5 px-3 rounded-lg hover:bg-slate-50 transition-colors">
                      <div className="flex items-center gap-3">
                        <span className={`w-1.5 h-1.5 rounded-full ${col.required ? 'bg-red-400' : 'bg-slate-300'}`}></span>
                        <div>
                          <span className="text-sm text-slate-700">{col.name}</span>
                          {col.note && <span className="text-xs text-slate-400 ml-2">({col.note})</span>}
                        </div>
                      </div>
                      <span className="text-xs px-2 py-1 rounded-md bg-slate-100 text-slate-500 font-medium">{col.type}</span>
                    </div>
                  ))}
                </div>
                <div className="mt-4 pt-4 border-t border-slate-100 flex items-center gap-4 text-xs text-slate-500">
                  <span className="flex items-center gap-1.5">
                    <span className="w-1.5 h-1.5 rounded-full bg-red-400"></span>
                    Required
                  </span>
                  <span className="flex items-center gap-1.5">
                    <span className="w-1.5 h-1.5 rounded-full bg-slate-300"></span>
                    Optional
                  </span>
                </div>
              </div>
            </div>

            {/* Output Info */}
            <div className="bg-gradient-to-br from-blue-50 to-indigo-50 rounded-2xl p-6 border border-blue-100">
              <div className="flex items-center gap-3 mb-3">
                <div className="w-8 h-8 rounded-lg bg-blue-500 flex items-center justify-center">
                  <svg className="w-4 h-4 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                  </svg>
                </div>
                <h3 className="font-semibold text-slate-800">Output Structure</h3>
              </div>
              <ul className="space-y-2 text-sm text-slate-600">
                <li className="flex items-center gap-2">
                  <svg className="w-4 h-4 text-blue-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M5 13l4 4L19 7" />
                  </svg>
                  "All Cases" consolidated sheet
                </li>
                <li className="flex items-center gap-2">
                  <svg className="w-4 h-4 text-blue-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M5 13l4 4L19 7" />
                  </svg>
                  Individual agreement sheets
                </li>
                <li className="flex items-center gap-2">
                  <svg className="w-4 h-4 text-blue-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M5 13l4 4L19 7" />
                  </svg>
                  Monthly EIR projections
                </li>
              </ul>
            </div>
          </div>
        </div>
      </main>

      {/* Footer */}
      <footer className="mt-16 py-8 border-t border-slate-200 bg-white">
        <div className="max-w-6xl mx-auto px-6">
          <div className="flex flex-col md:flex-row items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <img src={logoUrl} alt="TVS Credit" className="h-6 w-auto opacity-60" />
              <span className="text-sm text-slate-400">© {new Date().getFullYear()} TVS Credit Services Limited</span>
            </div>
            <div className="flex items-center gap-4">
              <span className="text-xs text-slate-400">Internal Use Only</span>
              <span className="text-xs px-3 py-1.5 rounded-full bg-slate-100 text-slate-500 font-medium">
                Finance Team
              </span>
            </div>
          </div>
        </div>
      </footer>
    </div>
  );
};

export default EIRIncomeProjectionTool;
