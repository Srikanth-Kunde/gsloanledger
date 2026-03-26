// ═══════════════════════════════════════════════════════════════
// LOAN MANAGEMENT SYSTEM - GOOGLE SHEETS v4.2
// - Performance Optimized All Members Ledger (Batch Processing)
// - Interest starts from NEXT month (not same month as loan)
// - Tally Audit Format
// - Strict January 31, 2026 Cut-off
// ═══════════════════════════════════════════════════════════════

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Loan Manager')
    .addItem('🔄 Generate Interest for All', 'generateInterestAll')
    .addSeparator()
    .addItem('📋 Single Member Ledger', 'showMemberLedger')
    .addItem('📋 ALL Members Ledger', 'generateAllMembersLedger')
    .addSeparator()
    .addItem('📊 Summary Report', 'showSummaryReport')
    .addItem('🗑️ Clear Auto-Generated Interest', 'clearAutoInterest')
    .addSeparator()
    .addItem('⚙️ Settings & Help', 'showSettings')
    .addItem('🐛 Debug Check', 'debugCheckData')
    .addToUi();
}

// ═══════════════════════════════════════════════════════════════
// CONFIGURATION
// ═══════════════════════════════════════════════════════════════

const SHEETS = {
  MEMBERS: 'Members',
  TRANSACTIONS: 'Transactions',
  RATES: 'Interest Rate Config',
  INTEREST: 'Interest Details',
  LEDGER: 'Ledger',
  ALL_LEDGERS: 'All Members Ledger',
  SUMMARY: 'Summary'
};

const CUTOFF_DATE = new Date(2026, 0, 31); // January 31, 2026

// ═══════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function parseDate(dateStr) {
  if (!dateStr) return null;
  if (dateStr instanceof Date) {
    return isNaN(dateStr.getTime()) ? null : dateStr;
  }
  dateStr = dateStr.toString().trim();
  if (dateStr.includes('T')) {
    const date = new Date(dateStr);
    return isNaN(date.getTime()) ? null : date;
  }
  const mmyyyyMatch = dateStr.match(/^(\d{1,2})[-\/](\d{4})$/);
  if (mmyyyyMatch) {
    return new Date(parseInt(mmyyyyMatch[2]), parseInt(mmyyyyMatch[1]) - 1, 1);
  }
  const ddmmyyyyMatch = dateStr.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (ddmmyyyyMatch) {
    return new Date(parseInt(ddmmyyyyMatch[3]), parseInt(ddmmyyyyMatch[2]) - 1, parseInt(ddmmyyyyMatch[1]));
  }
  const date = new Date(dateStr);
  return isNaN(date.getTime()) ? null : date;
}

function formatDate(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}-${month}-${year}`;
}

function formatMonthYear(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${months[date.getMonth()]}-${date.getFullYear()}`;
}

function getInterestRate(forDate, ratesData) {
  if (!forDate || !(forDate instanceof Date)) return 1.5;
  for (let i = 0; i < ratesData.length; i++) {
    const fromDate = parseDate(ratesData[i][1]);
    const toDate = parseDate(ratesData[i][2]);
    if (fromDate && toDate && forDate >= fromDate && forDate <= toDate) {
      return parseFloat(ratesData[i][0]) || 1.5;
    }
  }
  return 1.5;
}

function addMonths(date, months) {
  const newDate = new Date(date);
  newDate.setMonth(newDate.getMonth() + months);
  return newDate;
}

function getMonthStart(date) {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function getMonthEnd(date) {
  return new Date(date.getFullYear(), date.getMonth() + 1, 0);
}

function getDaysInMonth(date) {
  return new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
}

function formatCurrency(amount) {
  if (amount === 0 || amount === '' || amount === null || isNaN(amount)) return '';
  return '₹' + Math.round(amount).toLocaleString('en-IN');
}

function formatCurrencyNumber(amount) {
  if (amount === 0 || amount === '' || amount === null || isNaN(amount)) return 0;
  return Math.round(amount * 100) / 100;
}

function getVoucherColumn(headers) {
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] && headers[i].toString().replace(/\s+/g, ' ').trim().toLowerCase() === 'voucher type') {
      return i;
    }
  }
  return 4; // Default
}

// ═══════════════════════════════════════════════════════════════
// CORE: INTEREST GENERATION
// Interest starts from NEXT MONTH after loan is taken
// ═══════════════════════════════════════════════════════════════

function generateInterestForMember(memberId, memberName, upToDate, ratesData, txnData, headers) {
  const colMap = {};
  headers.forEach((h, i) => {
    if (h) colMap[h.toString().trim()] = i;
  });
  const voucherCol = getVoucherColumn(headers);
  
  // Get member transactions (excluding interest entries)
  const memberTxns = [];
  for (let i = 1; i < txnData.length; i++) {
    const row = txnData[i];
    const txnMemberId = row[colMap['Member ID']];
    const voucherType = row[voucherCol];
    
    if (voucherType && voucherType.toString().includes('Interest')) continue;
    
    if (txnMemberId == memberId) {
      const txnDate = parseDate(row[colMap['Date']]);
      if (!txnDate) continue;
      
      memberTxns.push({
        date: txnDate,
        voucherType: voucherType ? voucherType.toString().trim() : '',
        debit: parseFloat(row[colMap['Debit']] || 0),
        credit: parseFloat(row[colMap['Credit']] || 0),
        interestDays: row[colMap['Interest Days']] ? parseInt(row[colMap['Interest Days']]) : null,
        narration: row[colMap['Narration']] || ''
      });
    }
  }
  
  if (memberTxns.length === 0) {
    return { interest: [], principal: 0, totalInterest: 0, transactions: [] };
  }
  
  memberTxns.sort((a, b) => a.date - b.date);
  
  const firstDate = memberTxns[0].date;
  const interestEntries = [];
  
  // Start from the NEXT month after first loan
  let currentMonth = addMonths(getMonthStart(firstDate), 1);
  let principalBalance = 0;
  let totalInterest = 0;
  
  // Process first month's transactions to get initial principal (no interest for first month)
  const firstMonthTxns = memberTxns.filter(t => 
    t.date.getFullYear() === firstDate.getFullYear() &&
    t.date.getMonth() === firstDate.getMonth()
  );
  
  for (const txn of firstMonthTxns) {
    if (txn.voucherType === 'Loan' || txn.voucherType === 'Top-up') {
      principalBalance += txn.debit;
    } else if (txn.voucherType === 'Payment') {
      principalBalance -= txn.credit;
      if (principalBalance < 0) principalBalance = 0;
    }
  }
  
  // Now calculate interest from second month onwards
  while (currentMonth <= upToDate) {
    const monthEnd = getMonthEnd(currentMonth);
    const daysInMonth = getDaysInMonth(currentMonth);
    
    // Opening balance is the principal at start of this month
    const openingBalance = principalBalance;
    let monthInterest = 0;
    
    // Calculate interest on opening balance (loans from previous months)
    if (openingBalance > 0) {
      const rate = getInterestRate(currentMonth, ratesData);
      monthInterest = openingBalance * (rate / 100);
    }
    
    // Process transactions in this month
    const monthTxns = memberTxns.filter(t => 
      t.date.getFullYear() === currentMonth.getFullYear() &&
      t.date.getMonth() === currentMonth.getMonth()
    );
    
    for (const txn of monthTxns) {
      if (txn.voucherType === 'Loan' || txn.voucherType === 'Top-up') {
        // New loan - add to principal but NO interest this month
        // Interest on this new loan will start from NEXT month
        principalBalance += txn.debit;
        
        // If interest days specified, calculate pro-rata for THIS loan only
        if (txn.interestDays && txn.interestDays > 0) {
          const rate = getInterestRate(txn.date, ratesData);
          monthInterest += txn.debit * (rate / 100) * (txn.interestDays / 30);
        }
        // If no interest days specified, no interest this month for new loan
        
      } else if (txn.voucherType === 'Payment') {
        principalBalance -= txn.credit;
        if (principalBalance < 0) principalBalance = 0;
      }
    }
    
    // Record interest entry
    if (monthInterest > 0) {
      const rate = getInterestRate(currentMonth, ratesData);
      interestEntries.push({
        date: monthEnd,
        month: formatMonthYear(monthEnd),
        memberId: memberId,
        memberName: memberName,
        amount: Math.round(monthInterest * 100) / 100,
        rate: rate,
        openingBalance: Math.round(openingBalance * 100) / 100,
        closingBalance: Math.round(principalBalance * 100) / 100
      });
      totalInterest += monthInterest;
    }
    
    currentMonth = addMonths(currentMonth, 1);
  }
  
  return {
    interest: interestEntries,
    principal: Math.round(principalBalance * 100) / 100,
    totalInterest: Math.round(totalInterest * 100) / 100,
    transactions: memberTxns
  };
}

// ═══════════════════════════════════════════════════════════════
// CLEAR AUTO INTEREST
// ═══════════════════════════════════════════════════════════════

function clearAutoInterest() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clear Interest Details sheet
  const interestSheet = ss.getSheetByName(SHEETS.INTEREST);
  if (interestSheet) {
    interestSheet.clear();
  }
  
  // Remove Interest entries from Transactions (cleanup old format)
  const txnSheet = ss.getSheetByName(SHEETS.TRANSACTIONS);
  if (txnSheet) {
    const data = txnSheet.getDataRange().getValues();
    const voucherCol = getVoucherColumn(data[0]);
    
    for (let i = data.length - 1; i >= 1; i--) {
      const voucherType = data[i][voucherCol];
      if (voucherType && voucherType.toString().includes('Interest')) {
        txnSheet.deleteRow(i + 1);
      }
    }
  }
  
  ui.alert('✅ Auto-generated interest cleared!');
}

// ═══════════════════════════════════════════════════════════════
// GENERATE INTEREST FOR ALL MEMBERS
// ═══════════════════════════════════════════════════════════════

function generateInterestAll() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const response = ui.alert(
    'Generate Interest for All Members',
    'This will:\n' +
    '• Calculate interest starting from the MONTH AFTER loan\n' +
    '• Store interest in "Interest Details" tab\n' +
    '• Keep Transactions tab clean\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  // Clear old interest
  const interestSheet = getSheet(SHEETS.INTEREST);
  interestSheet.clear();
  
  // Clean Transactions of any old Interest entries
  const txnSheet = ss.getSheetByName(SHEETS.TRANSACTIONS);
  if (txnSheet) {
    const data = txnSheet.getDataRange().getValues();
    const voucherCol = getVoucherColumn(data[0]);
    for (let i = data.length - 1; i >= 1; i--) {
      const voucherType = data[i][voucherCol];
      if (voucherType && voucherType.toString().includes('Interest')) {
        txnSheet.deleteRow(i + 1);
      }
    }
  }
  
  // Setup Interest Details header
  interestSheet.appendRow([
    'Member ID', 'Member Name', 'Month', 'Opening Balance', 'Interest Amount', 'Rate %', 'Closing Balance', 'Date'
  ]);
  interestSheet.getRange('A1:H1').setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
  
  // Get data
  const rateSheet = ss.getSheetByName(SHEETS.RATES);
  const ratesData = rateSheet ? rateSheet.getDataRange().getValues().slice(1) : [];
  
  const memberSheet = ss.getSheetByName(SHEETS.MEMBERS);
  const memberData = memberSheet.getDataRange().getValues().slice(1);
  
  const txnData = txnSheet.getDataRange().getValues();
  const headers = txnData[0];
  
  const upToDate = CUTOFF_DATE;
  let processedCount = 0;
  let allInterestRows = [];
  
  for (const member of memberData) {
    const memberId = member[0];
    const memberName = member[1];
    
    if (!memberId) continue;
    
    const result = generateInterestForMember(memberId, memberName, upToDate, ratesData, txnData, headers);
    
    if (result.interest.length > 0) {
      for (const entry of result.interest) {
        allInterestRows.push([
          entry.memberId,
          entry.memberName,
          entry.month,
          entry.openingBalance,
          entry.amount,
          entry.rate,
          entry.closingBalance,
          formatDate(entry.date)
        ]);
      }
      processedCount++;
    }
  }
  
  // Write all interest entries
  if (allInterestRows.length > 0) {
    interestSheet.getRange(2, 1, allInterestRows.length, 8).setValues(allInterestRows);
    
    // Format currency columns
    interestSheet.getRange(2, 4, allInterestRows.length, 1).setNumberFormat('₹#,##0');
    interestSheet.getRange(2, 5, allInterestRows.length, 1).setNumberFormat('₹#,##0.00');
    interestSheet.getRange(2, 7, allInterestRows.length, 1).setNumberFormat('₹#,##0');
  }
  
  for (let i = 1; i <= 8; i++) {
    interestSheet.autoResizeColumn(i);
  }
  
  ui.alert(
    `✅ Interest Generated!\n\n` +
    `• Processed: ${processedCount} members\n` +
    `• Total entries: ${allInterestRows.length}\n` +
    `• Interest starts from NEXT month after loan\n\n` +
    `View "Interest Details" tab.`
  );
}

// ═══════════════════════════════════════════════════════════════
// SINGLE MEMBER LEDGER - TALLY FORMAT
// S.No | Date | Type | Debit | Credit | Interest | Balance | Narration
// ═══════════════════════════════════════════════════════════════

function showMemberLedger() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Generate Member Ledger',
    'Enter Member ID:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() !== ui.Button.OK) return;
  
  const memberId = parseInt(result.getResponseText());
  if (!memberId && memberId !== 0) {
    ui.alert('Invalid Member ID');
    return;
  }
  
  generateLedger(memberId);
}

function generateLedger(memberId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txnSheet = ss.getSheetByName(SHEETS.TRANSACTIONS);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBERS);
  const interestSheet = ss.getSheetByName(SHEETS.INTEREST);
  const ledgerSheet = getSheet(SHEETS.LEDGER);
  
  ledgerSheet.clear();
  ledgerSheet.clearConditionalFormatRules();
  if (ledgerSheet.getMaxRows() > 1) {
    ledgerSheet.getRange(1, 1, ledgerSheet.getMaxRows(), 9).clearFormat();
  }
  
  // Get member details
  const memberData = memberSheet.getDataRange().getValues();
  let memberName = '';
  let memberAddress = '';
  for (let i = 1; i < memberData.length; i++) {
    if (memberData[i][0].toString().trim() == memberId.toString().trim()) {
      memberName = memberData[i][1];
      memberAddress = memberData[i][3] || '';
      break;
    }
  }
  
  // Get transactions
  const txnData = txnSheet.getDataRange().getValues();
  const headers = txnData[0];
  const colMap = {};
  headers.forEach((h, i) => {
    if (h) colMap[h.toString().trim()] = i;
  });
  const voucherCol = getVoucherColumn(headers);
  
  // Build combined entries (transactions + interest)
  const allEntries = [];
  
  // Add transactions
  for (let i = 1; i < txnData.length; i++) {
    if (txnData[i][colMap['Member ID']] == memberId) {
      const vType = txnData[i][voucherCol];
      if (vType && !vType.toString().includes('Interest')) {
        const txnDate = parseDate(txnData[i][colMap['Date']]);
        if (txnDate && txnDate <= CUTOFF_DATE) {
          allEntries.push({
            date: txnDate,
            type: vType.toString().trim(),
            debit: parseFloat(txnData[i][colMap['Debit']] || 0),
            credit: parseFloat(txnData[i][colMap['Credit']] || 0),
            interest: 0,
            narration: txnData[i][colMap['Narration']] || '',
            isInterest: false
          });
        }
      }
    }
  }
  
  // Add interest entries
  if (interestSheet) {
    const intData = interestSheet.getDataRange().getValues();
    for (let i = 1; i < intData.length; i++) {
      if (intData[i][0] == memberId) {
        const intDate = parseDate(intData[i][7]); // Date column
        // Strict Cut-off: Skip entries after Jan 2026
        if (intDate > CUTOFF_DATE) continue;
        
        const intAmount = parseFloat(intData[i][4] || 0); // Interest Amount
        const rate = intData[i][5]; // Rate
        
        allEntries.push({
          date: intDate,
          type: 'Interest',
          debit: 0,
          credit: 0,
          interest: intAmount,
          narration: `Interest @${rate}%`,
          isInterest: true
        });
      }
    }
  }
  
  // Sort by date
  allEntries.sort((a, b) => {
    if (a.date.getTime() === b.date.getTime()) {
      // Put transactions before interest on same date
      return a.isInterest ? 1 : -1;
    }
    return a.date - b.date;
  });
  
  // Header Section - Set Rows correctly
  ledgerSheet.getRange(1, 1).setValue('LEDGER ACCOUNT');
  ledgerSheet.getRange(2, 1).setValue('');
  ledgerSheet.getRange(3, 1).setValue(`Name: ${memberName}      |      Member ID: ${memberId}`);
  ledgerSheet.getRange(4, 1).setValue(`Address: ${memberAddress || 'N/A'}`);
  ledgerSheet.getRange(5, 1).setValue('');
  ledgerSheet.getRange(6, 1).setValue('');
  
  // Column Headers - Tally Format
  const headerRow = 7;
  ledgerSheet.getRange(headerRow, 1, 1, 9).setValues([[
    "'Sl No", 'Date', 'Particulars', 'Vch Type', 'Debit (Dr)', 'Credit (Cr)', 'Interest', 'Balance', 'Narration'
  ]]);
  
  let runningBalance = 0;
  let totalDebit = 0;      // Total Loans (Principal + Top-up)
  let totalCredit = 0;     // Total Repayments
  let totalInterest = 0;   // Total Interest
  let sno = 0;
  let firstLoanDate = null;
  let lastPaymentDate = null;
  
  // Add opening balance row
  ledgerSheet.appendRow([
    '', '', 'Opening Balance', '', '', '', '', formatCurrency(0), ''
  ]);
  
  // Transaction rows
  for (const entry of allEntries) {
    sno++;
    
    if (entry.isInterest) {
      totalInterest += entry.interest;
      ledgerSheet.appendRow([
        sno, formatDate(entry.date), 'Interest Charged', 'Interest',
        '', '', formatCurrency(entry.interest), formatCurrency(runningBalance), entry.narration
      ]);
    } else if (entry.type === 'Loan' || entry.type === 'Top-up') {
      totalDebit += entry.debit;
      runningBalance += entry.debit;
      if (!firstLoanDate) firstLoanDate = entry.date;
      ledgerSheet.appendRow([
        sno, formatDate(entry.date), entry.narration || 'Loan Given', entry.type,
        formatCurrency(entry.debit), '', '', formatCurrency(runningBalance), entry.narration
      ]);
    } else if (entry.type === 'Payment') {
      totalCredit += entry.credit;
      runningBalance -= entry.credit;
      lastPaymentDate = entry.date;
      if (runningBalance < 0) runningBalance = 0;
      ledgerSheet.appendRow([
        sno, formatDate(entry.date), 'Repayment Received', 'Payment',
        '', formatCurrency(entry.credit), '', formatCurrency(runningBalance), entry.narration
      ]);
    }
  }
  
  // Totals Section
  ledgerSheet.appendRow(['']);
  ledgerSheet.appendRow([
    '', '', '', 'TOTALS',
    formatCurrency(totalDebit),
    formatCurrency(totalCredit),
    formatCurrency(totalInterest),
    '',
    ''
  ]);
  
  // Total Due Calculation (Principal Only as per user request)
  const totalDue = Math.max(0, totalDebit - totalCredit);

  // Account Summary - Unified rendering
  const summaryData = {
    memberName: memberName,
    memberId: memberId,
    firstLoanDate: firstLoanDate,
    lastPaymentDate: lastPaymentDate,
    totalDebit: totalDebit,
    totalCredit: totalCredit,
    totalInterest: totalInterest,
    totalDue: totalDue,
    asOnDate: formatDate(CUTOFF_DATE) // Cut-off date
  };
  
  renderAccountSummary(ledgerSheet, summaryData);
  
  // Footer spacing
  ledgerSheet.appendRow(['']);
  ledgerSheet.appendRow(['']);
  
  const topRow = ledgerSheet.getLastRow();
  const indexUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl() + '#gid=' + ledgerSheet.getSheetId() + '&range=A1';
  ledgerSheet.getRange(topRow, 4).setFormula(`=HYPERLINK("${indexUrl}", "⬆️ Back to Top")`).setFontColor('#1a73e8').setFontLine('underline');
  
  // Formatting
  // Header Formatting
  ledgerSheet.getRange('A1:I1').merge().setFontWeight('bold').setFontSize(18).setBackground('#1a73e8').setFontColor('white');
  ledgerSheet.getRange('A2:I2').merge();
  ledgerSheet.getRange('A3:I3').merge().setFontWeight('bold');
  ledgerSheet.getRange('A4:I4').merge();
  ledgerSheet.getRange(headerRow, 1, 1, 9).setFontWeight('bold').setBackground('#4285F4').setFontColor('white').setHorizontalAlignment('center');
  
  const lastRow = ledgerSheet.getLastRow();
  for (let r = 1; r <= lastRow; r++) {
    const cellValue = ledgerSheet.getRange(r, 4).getValue();
    if (cellValue === 'TOTALS') {
      ledgerSheet.getRange(r, 1, 1, 9).setFontWeight('bold').setBackground('#e8f0fe');
    }
  }
  
  for (let i = 1; i <= 9; i++) {
    ledgerSheet.autoResizeColumn(i);
  }
  ledgerSheet.setColumnWidth(1, 50); // Explicit Sl No width
}

// ═══════════════════════════════════════════════════════════════
// GENERATE ALL MEMBERS LEDGER - TALLY FORMAT (OPTIMIZED)
// ═══════════════════════════════════════════════════════════════

function generateAllMembersLedger() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const response = ui.alert(
    'Generate All Members Ledger',
    'This will create ledgers for ALL members using high-performance batch processing.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const allLedgerSheet = getSheet(SHEETS.ALL_LEDGERS);
  allLedgerSheet.clear();
  allLedgerSheet.clearConditionalFormatRules();
  
  const txnSheet = ss.getSheetByName(SHEETS.TRANSACTIONS);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBERS);
  const interestSheet = ss.getSheetByName(SHEETS.INTEREST);
  
  const memberData = memberSheet.getDataRange().getValues().slice(1);
  const txnData = txnSheet.getDataRange().getValues();
  const headers = txnData[0];
  
  const colMap = {};
  headers.forEach((h, i) => { if (h) colMap[h.toString().trim()] = i; });
  const voucherCol = getVoucherColumn(headers);
  
  // 1. Group Data Upfront
  const txnsByMember = {};
  for (let i = 1; i < txnData.length; i++) {
    const mid = txnData[i][colMap['Member ID']];
    if (!mid) continue;
    const vType = txnData[i][voucherCol];
    if (vType && !vType.toString().includes('Interest')) {
      const txnDate = parseDate(txnData[i][colMap['Date']]);
      if (txnDate && txnDate <= CUTOFF_DATE) {
        if (!txnsByMember[mid]) txnsByMember[mid] = [];
        txnsByMember[mid].push({
          date: txnDate,
          type: vType.toString().trim(),
          debit: parseFloat(txnData[i][colMap['Debit']] || 0),
          credit: parseFloat(txnData[i][colMap['Credit']] || 0),
          narration: txnData[i][colMap['Narration']] || '',
          isInterest: false
        });
      }
    }
  }
  
  let interestEntriesByMember = {};
  if (interestSheet) {
    const intData = interestSheet.getDataRange().getValues();
    for (let i = 1; i < intData.length; i++) {
      const mid = intData[i][0];
      const intDate = parseDate(intData[i][7]);
      if (intDate && intDate <= CUTOFF_DATE) {
        if (!interestEntriesByMember[mid]) interestEntriesByMember[mid] = [];
        interestEntriesByMember[mid].push({
          date: intDate, amount: parseFloat(intData[i][4] || 0), rate: intData[i][5], isInterest: true
        });
      }
    }
  }
  
  const allRows = [];
  const memberIndices = [];
  let memberCount = 0;
  let grandDebit = 0, grandCredit = 0, grandInterest = 0, grandDue = 0;

  // Placeholder rows for Index
  const PLACEHOLDER_COUNT = memberData.length + 10; 
  for (let i = 0; i < PLACEHOLDER_COUNT; i++) {
    allRows.push(['', '', '', '', '', '', '', '', '']);
  }

  // 2. Build Memory Buffer
  for (const member of memberData) {
    const memberId = member[0];
    const memberName = member[1];
    if (!memberId) continue;
    
    const mTxns = (txnsByMember[memberId] || []).concat(
      (interestEntriesByMember[memberId] || []).map(ie => ({
        date: ie.date, type: 'Interest', debit: 0, credit: 0, 
        interest: ie.amount, narration: `Interest @${ie.rate}%`, isInterest: true
      }))
    );
    
    if (mTxns.length === 0) continue;
    mTxns.sort((a, b) => {
      if (a.date.getTime() === b.date.getTime()) return a.isInterest ? 1 : -1;
      return a.date - b.date;
    });
    
    const startRowIndex = allRows.length;
    allRows.push([`${memberName} (ID: ${memberId})`, '', '', '', '', '', '', '', '']);
    allRows.push(["'Sl No", 'Date', 'Particulars', 'Vch Type', 'Debit (Dr)', 'Credit (Cr)', 'Interest', 'Balance', 'Narration']);
    
    let runningBalance = 0, totalDebit = 0, totalCredit = 0, totalInterest = 0;
    let sno = 0, firstLoanDate = null, lastPaymentDate = null;
    
    for (const entry of mTxns) {
      sno++;
      if (entry.isInterest) {
        totalInterest += entry.interest;
        allRows.push([sno, formatDate(entry.date), 'Interest Charged', 'Interest', '', '', formatCurrency(entry.interest), formatCurrency(runningBalance), entry.narration]);
      } else if (entry.type === 'Loan' || entry.type === 'Top-up') {
        totalDebit += entry.debit; runningBalance += entry.debit;
        if (!firstLoanDate) firstLoanDate = entry.date;
        allRows.push([sno, formatDate(entry.date), entry.narration || 'Loan', entry.type, formatCurrency(entry.debit), '', '', formatCurrency(runningBalance), entry.narration]);
      } else if (entry.type === 'Payment') {
        totalCredit += entry.credit; runningBalance -= entry.credit;
        lastPaymentDate = entry.date;
        if (runningBalance < 0) runningBalance = 0;
        allRows.push([sno, formatDate(entry.date), 'Repayment', 'Payment', '', formatCurrency(entry.credit), '', formatCurrency(runningBalance), entry.narration]);
      }
    }
    
    const principalDue = Math.max(0, totalDebit - totalCredit);
    const mTotalDue = principalDue;
    const summaryRows = getAccountSummaryRows({ 
      memberName: memberName,
      memberId: memberId,
      firstLoanDate, lastPaymentDate, totalDebit, totalCredit, totalInterest, 
      totalDue: mTotalDue, asOnDate: '31-01-2026' 
    });
    
    allRows.push(['']);
    summaryRows.forEach(row => allRows.push(row));
    allRows.push(['', '', '', 'BACK_TO_INDEX_LINK', '', '', '', '', '']); // Marker for hyperlinks
    allRows.push(['']);
    allRows.push(['─'.repeat(100)]);
    allRows.push(['']);
    
    memberIndices.push({
      sno: memberCount + 1, id: memberId, name: memberName,
      principal: principalDue, interest: totalInterest, total: mTotalDue,
      row: startRowIndex + 1
    });
    
    memberCount++;
    grandDebit += totalDebit; grandCredit += totalCredit; grandInterest += totalInterest; grandDue += mTotalDue;
  }
  
  // 3. Render Index in Buffer
  const indexHeader = [
    ['ALL MEMBERS LEDGER INDEX', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', ''],
    ['📋 TABLE OF CONTENTS - Click on Member Name to Jump', '', '', '', '', '', '', '', ''],
    ["'Sl No", 'ID', 'Member Name', 'Principal Due', 'Interest Accrued', 'Total Due', 'Navigation', '', '']
  ];
  for (let i = 0; i < indexHeader.length; i++) { allRows[i] = indexHeader[i]; }
  
  memberIndices.forEach((item, i) => {
    allRows[4 + i] = [item.sno, item.id, item.name, formatCurrency(item.principal), formatCurrency(item.interest), formatCurrency(item.total), 'JUMP_LINK', '', ''];
  });
  
  // 4. Batch Write
  if (allRows.length > 0) { allLedgerSheet.getRange(1, 1, allRows.length, 9).setValues(allRows); }
  
  // 5. Batch Formatting & Hyperlinks
  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const gid = allLedgerSheet.getSheetId();
  const indexUrl = `${sheetUrl}#gid=${gid}&range=A1`;

  allLedgerSheet.getRange(1, 1, 1, 9).merge().setFontWeight('bold').setFontSize(20).setBackground('#1a73e8').setFontColor('white').setHorizontalAlignment('center');
  allLedgerSheet.getRange(3, 1, 1, 9).merge().setFontWeight('bold').setBackground('#f8f9fa');
  allLedgerSheet.getRange(4, 1, 1, 9).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
  
  memberIndices.forEach((item, i) => {
    const r = 5 + i;
    const jumpUrl = `${sheetUrl}#gid=${gid}&range=A${item.row}`;
    allLedgerSheet.getRange(r, 3).setFormula(`=HYPERLINK("${jumpUrl}", "${item.name}")`).setFontColor('#1a73e8').setFontLine('underline');
  });
  
  const values = allLedgerSheet.getRange(1, 1, allRows.length, 9).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (row[3] === 'BACK_TO_INDEX_LINK') {
      allLedgerSheet.getRange(i + 1, 4).setFormula(`=HYPERLINK("${indexUrl}", "⬆️ Back to Index")`).setFontColor('#1a73e8').setFontLine('underline');
    }
    if (row[0] === 'ACCOUNT SUMMARY') {
      const summaryHeight = 15;
      allLedgerSheet.getRange(i + 1, 1, summaryHeight, 9).setBackground('#C6EFCE');
      allLedgerSheet.getRange(i + 1, 1, 1, 9).merge().setFontWeight('bold').setHorizontalAlignment('center');
      allLedgerSheet.getRange(i + 2, 1, 1, 9).merge(); // Dashed
      allLedgerSheet.getRange(i + 3, 4, 7, 2).setFontWeight('bold'); // Labels/Values (Name, ID, etc)
      allLedgerSheet.getRange(i + 10, 1, 1, 9).merge(); // Dashed
      allLedgerSheet.getRange(i + 11, 1, 1, 9).merge(); // Double
      allLedgerSheet.getRange(i + 12, 4, 1, 2).setFontWeight('bold'); // Total Due
      allLedgerSheet.getRange(i + 13, 1, 1, 9).merge(); // Double
    }
    if (typeof row[0] === 'string' && row[0].includes('(ID: ')) {
      allLedgerSheet.getRange(i + 1, 1, 1, 9).merge().setFontWeight('bold').setFontSize(12).setBackground('#e8f0fe');
      allLedgerSheet.getRange(i + 2, 1, 1, 9).setFontWeight('bold').setBackground('#d0e0f0').setFontSize(10);
    }
  }

  allLedgerSheet.setFrozenRows(4);
  
  const grandTotalRow = allLedgerSheet.getLastRow() + 2;
  const gtRows = [
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════'],
    ['', '', '', 'GRAND TOTALS (' + memberCount + ' Members)', formatCurrency(grandDebit), formatCurrency(grandCredit), formatCurrency(grandInterest), '', ''],
    ['', '', '', 'TOTAL AMOUNT DUE (Principal)', formatCurrency(grandDue), '', '', '', ''],
    ['═══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════']
  ];
  allLedgerSheet.getRange(grandTotalRow, 1, 5, 9).setValues(gtRows);
  allLedgerSheet.getRange(grandTotalRow + 2, 4, 1, 2).setFontWeight('bold').setFontSize(14).setBackground('#c6efce');
  
  for (let i = 1; i <= 9; i++) { allLedgerSheet.autoResizeColumn(i); }
  allLedgerSheet.setColumnWidth(1, 50); // Explicitly narrow Sl No column
  ui.alert(`✅ All Members Ledger Generated!\n\n• ${memberCount} members processed.`);
}

// ═══════════════════════════════════════════════════════════════
// SUMMARY REPORT - TALLY FORMAT
// ═══════════════════════════════════════════════════════════════

function showSummaryReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBERS);
  const txnSheet = ss.getSheetByName(SHEETS.TRANSACTIONS);
  const interestSheet = ss.getSheetByName(SHEETS.INTEREST);
  const summarySheet = getSheet(SHEETS.SUMMARY);
  
  summarySheet.clear();
  
  const cutOffDate = CUTOFF_DATE;
  summarySheet.appendRow(['LOAN SUMMARY REPORT']);
  summarySheet.appendRow(['Report Date:', formatDate(new Date())]);
  summarySheet.appendRow(["'Sl no", 'ID', 'Member Name', 'First Loan Date', 'Total Loan(Dr)', 'Total Paid (Cr)', 'Principal Due', 'Interest Charged', 'Last Repayment Date', 'Total Due']);
  const headerRow = summarySheet.getLastRow();
  
  const memberData = memberSheet.getDataRange().getValues().slice(1);
  const txnData = txnSheet.getDataRange().getValues();
  const headers = txnData[0];
  const colMap = {};
  headers.forEach((h, i) => { if (h) colMap[h.toString().trim()] = i; });
  const voucherCol = getVoucherColumn(headers);
  
  let interestByMember = {};
  if (interestSheet) {
    const intData = interestSheet.getDataRange().getValues();
    for (let i = 1; i < intData.length; i++) {
      const mid = intData[i][0];
      const intDate = parseDate(intData[i][7]);
      if (intDate && intDate <= cutOffDate) {
        if (!interestByMember[mid]) interestByMember[mid] = 0;
        interestByMember[mid] += parseFloat(intData[i][4] || 0);
      }
    }
  }
  
  let grandLoans = 0, grandPaid = 0, grandPrincipal = 0, grandInterest = 0, grandTotal = 0;
  let memberCount = 0, sno = 0;
  
  for (const member of memberData) {
    const memberId = member[0];
    const memberName = member[1];
    if (!memberId) continue;
    
    let totalLoans = 0, totalPaid = 0;
    let firstLoanDate = null, lastRepaymentDate = null;
    
    for (let i = 1; i < txnData.length; i++) {
      const row = txnData[i];
      if (row[colMap['Member ID']] == memberId) {
        const txnDate = parseDate(row[colMap['Date']]);
        if (!txnDate || txnDate > cutOffDate) continue;
        const voucherType = row[voucherCol];
        if (!voucherType || voucherType.toString().includes('Interest')) continue;
        
        const debit = parseFloat(row[colMap['Debit']] || 0);
        const credit = parseFloat(row[colMap['Credit']] || 0);
        const vType = voucherType.toString().trim();
        
        if (vType === 'Loan' || vType === 'Top-up') {
          totalLoans += debit;
          if (!firstLoanDate || txnDate < firstLoanDate) firstLoanDate = txnDate;
        } else if (vType === 'Payment') {
          totalPaid += credit;
          if (!lastRepaymentDate || txnDate > lastRepaymentDate) lastRepaymentDate = txnDate;
        }
      }
    }
    
    const principal = totalLoans - totalPaid;
    const totalInterest = interestByMember[memberId] || 0;
    const totalDue = principal; // Principal only as per user clarification
    
    if (totalLoans > 0 || totalPaid > 0) {
      sno++;
      summarySheet.appendRow([sno, memberId, memberName, formatDate(firstLoanDate), formatCurrency(totalLoans), formatCurrency(totalPaid), formatCurrency(principal), formatCurrency(totalInterest), formatDate(lastRepaymentDate) || 'No Payments', formatCurrency(totalDue)]);
      grandLoans += totalLoans; grandPaid += totalPaid; grandPrincipal += principal; grandInterest += totalInterest; grandTotal += totalDue;
      memberCount++;
    }
  }
  
  const lastDataRow = summarySheet.getLastRow();
  summarySheet.appendRow(['']);
  const separator1Row = lastDataRow + 2;
  summarySheet.appendRow(['═══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════']);
  summarySheet.getRange(separator1Row, 1, 1, 10).merge();

  summarySheet.appendRow(['', '', 'GRAND TOTAL (' + memberCount + ' Members)', '', formatCurrency(grandLoans), formatCurrency(grandPaid), formatCurrency(grandPrincipal), formatCurrency(grandInterest), '', formatCurrency(grandTotal)]);
  const grandTotalRow = lastDataRow + 3;
  summarySheet.getRange(grandTotalRow, 1, 1, 10).setFontWeight('bold').setBackground('#c6efce');

  const separator2Row = lastDataRow + 4;
  summarySheet.appendRow(['═══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════']);
  summarySheet.getRange(separator2Row, 1, 1, 10).merge();
  
  // Final Formatting
  summarySheet.getRange('A1:J1').merge().setFontWeight('bold').setFontSize(18).setBackground('#1a73e8').setFontColor('white');
  summarySheet.getRange(headerRow, 1, 1, 10).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
  
  for (let i = 1; i <= 10; i++) summarySheet.autoResizeColumn(i);
  summarySheet.setColumnWidth(1, 50); // Explicit Sl No width
  
  SpreadsheetApp.getUi().alert(`✅ Summary Report generated!`);
}

// ═══════════════════════════════════════════════════════════════
// UNIFIED SUMMARY RENDERERS
// ═══════════════════════════════════════════════════════════════

function renderAccountSummary(sheet, data) {
  const rows = getAccountSummaryRows(data);
  const startRow = sheet.getLastRow() + 2;
  sheet.getRange(startRow, 1, rows.length, 9).setValues(rows);
  
  // Formatting
  sheet.getRange(startRow, 1, rows.length, 9).setBackground('#C6EFCE');
  sheet.getRange(startRow, 1, 1, 9).merge().setFontWeight('bold').setHorizontalAlignment('center'); // Title
  sheet.getRange(startRow + 1, 1, 1, 9).merge(); // Dashed line
  sheet.getRange(startRow + 2, 4, 7, 2).setFontWeight('bold'); // Labels and Values
  sheet.getRange(startRow + 9, 1, 1, 9).merge(); // Dashed line
  sheet.getRange(startRow + 10, 1, 1, 9).merge(); // Double line
  sheet.getRange(startRow + 11, 4, 1, 2).setFontWeight('bold'); // Total Due
  sheet.getRange(startRow + 12, 1, 1, 9).merge(); // Double line
  
  sheet.getRange(sheet.getLastRow(), 2, 1, 7).setBorder(true, null, true, null, null, null);
}

function getAccountSummaryRows(data) {
  return [
    ['ACCOUNT SUMMARY', '', '', '', '', '', '', '', ''],
    ['───────────────────────────────────────────────────────────────────────────────────────', '', '', '', '', '', '', '', ''],
    ['', '', '', 'Member Name:', '', data.memberName, '', '', ''],
    ['', '', '', 'Member ID:', '', data.memberId, '', '', ''],
    ['', '', '', 'First Loan Date:', '', formatDate(data.firstLoanDate) || 'No Loans', '', '', ''],
    ['', '', '', 'Last Repayment Date:', '', formatDate(data.lastPaymentDate) || 'No Payments', '', '', ''],
    ['', '', '', 'Total Principal Given (Dr)', '', formatCurrency(data.totalDebit), '', '', ''],
    ['', '', '', 'Total Repayment Received (Cr)', '', formatCurrency(data.totalCredit), '', '', ''],
    ['', '', '', 'Interest paid:', '', formatCurrency(data.totalInterest), '', '', ''],
    ['───────────────────────────────────────────────────────────────────────────────────────', '', '', '', '', '', '', '', ''],
    ['═══════════════════════════════════════════════════════════════════════════════════════', '', '', '', '', '', '', '', ''],
    ['', '', '', data.totalDue <= 0 ? 'STATUS:' : 'TOTAL DUE (Principal):', '', data.totalDue <= 0 ? 'BALANCE: ZERO' : formatCurrency(data.totalDue), '', '', ''],
    ['═══════════════════════════════════════════════════════════════════════════════════════', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', ''],
    ['', 'Closing Balance', '', '', '', '', formatCurrency(data.totalDue), 'As on ' + data.asOnDate, '']
  ];
}

function showSettings() {
  const ui = SpreadsheetApp.getUi();
  const helpText = `
╔═══════════════════════════════════════════════════════════╗
║               LOAN MANAGER v4.2 - HELP                    ║
╠═══════════════════════════════════════════════════════════╣
║                                                           ║
║  📊 INTEREST CALCULATION RULE:                            ║
║  • Loan taken in Jan → Interest starts from Feb           ║
║  • Interest calculated on opening balance each month      ║
║  • Strict Cut-off Date: 31-Jan-2026                       ║
║                                                           ║
║  📁 TAB STRUCTURE:                                        ║
║  • Members - Master list (ID, Name, Address)              ║
║  • Transactions - "Loan", "Top-up", "Payment" entries     ║
║  • Interest Rate Config - Rate periods                    ║
║  • Interest Details - Auto-generated logs                 ║
║  • Ledger - Single member statement                       ║
║  • All Members Ledger - Batch statements (FAST)           ║
║  • Summary - Business overview dashboard                   ║
║                                                           ║
╚═══════════════════════════════════════════════════════════╝
  `;
  ui.alert('Settings & Help', helpText, ui.ButtonSet.OK);
}

function debugCheckData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let report = "=== DEBUG REPORT ===\n\n";
  const memberSheet = ss.getSheetByName('Members');
  report += memberSheet ? "✅ Members: " + (memberSheet.getLastRow() - 1) + " members\n" : "❌ Members sheet NOT FOUND!\n";
  const txnSheet = ss.getSheetByName('Transactions');
  report += txnSheet ? "✅ Transactions: " + (txnSheet.getLastRow() - 1) + " entries\n" : "❌ Transactions sheet NOT FOUND!\n";
  const intSheet = ss.getSheetByName('Interest Details');
  report += intSheet ? "✅ Interest Details: " + (intSheet.getLastRow() - 1) + " entries\n" : "⚠️ Interest Details Empty\n";
  const htmlOutput = HtmlService.createHtmlOutput('<pre style="font-family: monospace; font-size: 14px; padding: 20px;">' + report + '</pre>').setWidth(600).setHeight(400);
  ui.showModalDialog(htmlOutput, 'Debug Report');
}