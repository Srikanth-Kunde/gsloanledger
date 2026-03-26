# Loan Management System (Google Apps Script) v4.8

A comprehensive loan tracking and interest calculation system built on Google Sheets and Apps Script.

## 🚀 Navigation & Features

### 📋 All Members Ledger (With Index)
Professional batch reporting with a built-in **Navigation Index** and high-speed **Batch Processing**.
- **Execution Speed**: Uses memory buffering to generate statements for all members instantly, avoiding script timeouts.
- **Quick Jump**: Click any member name in the index at the top to jump directly to their ledger.
- **Back to Index**: Use the "⬆️ Back to Index" link in each member's summary to return to the top.
- **Account Summary**: Synchronized green summary blocks for every member.

### 📋 Single Member Ledger
Detailed individual statements in Tally format.
- **Total Due**: Reflects the **Principal Outstanding** as of the cut-off date.

### 🔄 Interest Engine
- Automated monthly interest generation with configurable rates.
- **Strict Cut-off**: Calculations stop at **31-01-2026**.
- **Pro-rata Support**: Supports partial months via "Interest Days".

### 📊 Summary Dashboard
A high-level overview of dues across all members.
- **Detailed Tracking**: Includes `First Loan Date` and `Last Repayment Date` for each member.
- **Audit Ready**: Shows `Principal Due` vs `Interest Charged` separately in a 10-column layout.
- **Fixed Calculations**: All totals strictly honor the **31-01-2026** cut-off.

## 📁 Sheet Structure
- **Members**: Master list of member IDs, names, and contact details.
- **Transactions**: Log of all "Loan", "Top-up", and "Payment" voucher types.
- **Interest Rate Config**: Time-based interest rates.
- **Interest Details**: Auto-generated monthly interest logs.
- **Ledger / All Members Ledger**: The Tally-style reporting outputs.
- **Summary**: High-level cross-member dues dashboard.

## 🛠️ Setup & Usage
1. Manage your members in the **Members** tab.
2. Record loans and payments in the **Transactions** tab.
3. Access all functions through the **📊 Loan Manager** custom menu.
4. Run **"🔄 Generate Interest for All"** periodically to update dues.
5. Generate ledgers to view individual or batch statements.

---
*For technical details, see [PRD.md](./PRD.md).*  
*For a log of recent improvements, see [ChangeTracker.md](./ChangeTracker.md).*
