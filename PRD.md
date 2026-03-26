# Product Requirements Document (PRD) - Loan Management System v4.8

## 1. Overview
A Google Apps Script-based loan tracking and interest calculation engine designed for efficient member management, automated interest accrual, and professional Tally-style reporting.

## 2. Core Features
- **Member Management**: Track member IDs, names, addresses, and loan history.
- **Automated Interest Engine**: Generate monthly interest entries based on configurable rates.
- **Tally-style Ledgers**: Professional single-member and batch-member reports.
- **Summary Dashboard**: 10-column high-level overview including `First Loan Date`, `Last Repayment Date`, and `Interest Charged` for total audit visibility.
- **Navigation Index**: Clickable Table of Contents for batch ledgers to quickly jump to any member.

### 📋 Single Member Ledger
Detailed individual statements in Tally format.
- Includes **Principal Balance** tracking.
- Green highlighted **Account Summary** with key loan dates.
- **Total Due**: Accurately reflects **Principal balance** only (as interest is settled separately).

## 3. Financial Calculation Rules
- **Interest Cycle**: Interest starts accruing from the month *following* the loan or top-up (e.g., Jan → Feb interest).
- **Principal-Only Balance**: The "Balance" column in ledgers tracks only the principal amount (Loans + Top-ups - Repayments).
- **Interst Pro-rata**: Supports "Interest Days" for partial-month interest on new loans or top-ups.
- **Strict Cut-off**: All calculations (Loans, Payments, and Interest) stop at a hardcoded cut-off date of **January 31, 2026**. Data after this date is ignored in all reports.
- **Interest Accrual**: Calculated as simple interest on the outstanding principal balance each month.

## 4. UI & Formatting Specifications
- **Tally Layout**: Multi-column ledger with Date, Particulars, Vch Type, Debit, Credit, Interest, and Balance.
- **Account Summary Block**: A green-highlighted summary at the end of each ledger showing:
    - First Loan Date & Last Repayment Date.
    - Total Principal Given vs. Total Repayment Received.
    - Total Interest Paid/Accrued.
    - **TOTAL DUE (Principal)** in bold.
- **Dynamic Headers**: White text on dark blue background (`#4285F4`).
- **Summary Dashboard Columns**: 10 columns: `Sl no`, `ID`, `Member Name`, `First Loan Date`, `Total Loan(Dr)`, `Total Paid (Cr)`, `Principal Due`, `Interest Charged`, `Last Repayment Date`, and `Total Due`.
- **Index Navigation**: Clickable member names and "Back to Index" links for batch reports.
