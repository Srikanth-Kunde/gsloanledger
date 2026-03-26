# Change Tracker - Loan Management System

A chronological record of project enhancements, refactorings, and bug fixes.

## [v4.7] - 2026-03-26
### Fixed
- **Architectural Hardening**: Fixed a column mismatch exception in `showSummaryReport()` and `generateAllMembersLedger()`. Padded all rows in the 2D summation blocks to a consistent column count (10 and 9 respectively), ensuring 100% compatibility with the Google Sheets `setValues` API.

## [v4.6] - 2026-03-26
### Fixed
- **Summary Report Stability**: Refactored the "Grand Totals" section in `showSummaryReport` to use batch rendering (`setValues`). This resolves a synchronization issue where the green total row would occasionally appear blank.

## [v4.5] - 2026-03-26
### Fixed
- **Header Data Integrity**: Restructured the Single Member Ledger header to merge ranges *before* setting combined value strings, preventing Google Sheets from discarding Member Name/ID data.
- **Universal Column Width Enforcement**: Implemented explicit `setColumnWidth(1, 50)` across all reporting modules (Summary, Single Ledger, All Ledger) to ensure a consistent, compact "Sl No" column.
- **Separator Row Merging**: Synchronized the merging of decorative separator rows in the Summary Report to mirror the fix in the member ledgers, resolving unintended layout expansion.
- **Enhanced Account Summary**: Finalized the inclusion of **Member Name** and **Member ID** in all green summary blocks for 10x traceability.

## [v4.4] - 2026-03-26
### Added
- **Improved Layouts**: Merged decorative separator rows across all columns in **Summary Report**, **Single Ledger**, and **Batch Ledger** to prevent unintended column expansion.
- **Enhanced Account Summary**: Added **Member Name** and **Member ID** to the green summary block for better traceability.
- **Column Width Management**: Explicitly set the "Sl No" column width to prevent unintended expansion during auto-resizing.

### Fixed
- **Single Ledger Header**: Fixed a data comparison bug that was preventing Name and ID from printing in the single member ledger header.

## [v4.3] - 2026-03-26
### Fixed
- **Summary Report Bug**: Resolved `ReferenceError: totalInterest is not defined` in the `showSummaryReport` function.
- **Rounding Strategy**: Standardized currency rounding across reports (Summary, Single Ledger, All Members Ledger) for financial consistency.
- **Column Alignment**: Ensured "Total Due" always reflects "Principal Only" as per latest financial rules, while maintaining "Interest Charged" as a separate audit column.

## [v4.2] - 2026-03-26
### Added
- **Performance Optimization**: Refactored "All Members Ledger" to use high-performance batch processing (memory buffering). This resolves execution timeout issues for large datasets.
- **Financial Alignment**: Aligned "Total Due" calculation across all reports (Summary, Single Ledger, All Ledger) to consistently reflect **Principal Outstanding Only** (as per user clarification that interest is considered paid/settled separately).

### Fixed
- **Grand Totals**: Updated the All Members Ledger grand total footer to explicitly show the **Principal** total due.
- **Reporting Consistency**: Synchronized the green "Account Summary" block logic with the 10-column Summary Report.

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
