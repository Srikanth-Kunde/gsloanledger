# Change Tracker - Loan Management System

A chronological record of project enhancements, refactorings, and bug fixes.

## [v4.1] - 2026-03-26
### Added
- **Expanded Summary Report**: Now features 10 comprehensive columns including `First Loan Date`, `Last Repayment Date`, and `Interest Charged`.
- **Global Date Filtering**: Strict January 31, 2026 cut-off now applied to all transaction and interest processing in Summary and Ledger reports.
- **Top-up Voucher Support**: Expanded logic to treat both 'Loan' and 'Top-up' as debit/loan voucher types across all calculation and reporting engines.
- **Transactions**: Log of all "Loan", "Top-up", and "Payment" voucher types.

### Fixed
- **Summary Layout**: Redesigned to follow a 3-row header structure (Title, Date, Columns) to match user requirements.
- **Highlighting Logic**: Fixed a bug where data rows were being highlighted as headers by implementing dynamic row indexing.
- **Header Cleanliness**: Removed auto-hyperlinking on the "Sl no" column in the Summary report.

## [v4.0] - 2026-03-26

## [v3.0] - 2026-03-25
### Added
- **Date Tracking**: Added "First Loan Date" and "Last Repayment Date" to ledger summaries.
- **Interest paid label**: Updated label for total accrued interest.
- **Green Highlighting**: Applied `#C6EFCE` background to ledger summary blocks.

### Changed
- **Cut-off Enforcement**: Implemented system-wide cut-off date of January 31, 2026.

## [v2.0] - Earlier
### Added
- **Tally-Style Ledger Layout**: Professional multi-column reporting.
- **All Members Ledger**: Batch generation of member statements.
- **Interest Cycle**: Interest starts accruing from the month *following* the loan (e.g., Jan loan or top-up → Feb interest).
- **Interest Details Sheet**: Automated logging of monthly interest entries.
- **Summary Report**: High-level cross-member dashboard.

## [v1.0] - Initial
- Core implementation of Member, Transaction, and Interest Rate management.
- Basic interest calculation engine.
