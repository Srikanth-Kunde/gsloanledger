# Product Requirements Document (PRD) - Loan Management System

## 1. Overview
A Google Apps Script-based loan tracking and interest calculation engine designed for efficient member management, automated interest accrual, and professional Tally-style reporting.

## 2. Core Features
- **Member Management**: Track member IDs, names, addresses, and loan history.
- **Automated Interest Engine**: Generate monthly interest entries based on configurable rates.
- **Tally-style Ledgers**: Professional single-member and batch-member reports.
- **Summary Reports**: High-level overview of total principal, interest, and dues across all members.
- **Navigation Index**: Clickable Table of Contents for batch ledgers to quickly jump to any member.

## 3. Financial Calculation Rules
- **Interest Cycle**: Interest starts accruing from the month *following* the loan (e.g., Jan loan → Feb interest).
- **Principal-Only Balance**: The "Balance" column in ledgers tracks only the principal amount (Loans - Repayments).
- **Interst Pro-rata**: Supports "Interest Days" for partial-month interest on new loans or top-ups.
- **Strict Cut-off**: All calculations stop at a hardcoded cut-off date of **January 31, 2026**.
- **Interest Accrual**: Calculated as simple interest on the outstanding principal balance each month.

## 4. UI & Formatting Specifications
- **Tally Layout**: Multi-column ledger with Date, Particulars, Vch Type, Debit, Credit, Interest, and Balance.
- **Account Summary Block**: A green-highlighted summary at the end of each ledger showing:
    - First Loan Date & Last Repayment Date.
    - Total Principal Given vs. Total Repayment Received.
    - Total Interest Paid/Accrued.
    - **TOTAL DUE (Principal)** in bold.
- **Dynamic Headers**: White text on dark blue background (`#4285F4`).
- **Index Navigation**: Clickable member names and "Back to Index" links for batch reports.
