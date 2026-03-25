# Change Tracker - Loan Management System

A chronological record of project enhancements, refactorings, and bug fixes.

## [v4.0] - 2026-03-26
### Added
- **Navigation Index**: Clickable Table of Contents at the top of the "All Members Ledger" sheet.
- **Back to Index**: Jump links added to every member's summary for quick navigation.
- **Synchronized Summaries**: Uniform "ACCOUNT SUMMARY" block across both single and batch reports.
- **Fixed "As on" Date**: Reports now explicitly show "As on 31-01-2026" to reflect the cut-off.

### Fixed
- **Principal-Only Balance**: Resolved a bug where interest was inflating the running balance column. Balance now tracks only Principal (Debit - Credit).
- **Blue Highlight Bug**: Fixed Row 7 styling issue by making header detection dynamic (targeting Row 4).
- **Hyperlink Fix**: Changed "Sl.No" to "Sl No" (handled as text) to prevent annoying auto-hyperlinks.

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
- **Interest Details Sheet**: Automated logging of monthly interest entries.
- **Summary Report**: High-level cross-member dashboard.

## [v1.0] - Initial
- Core implementation of Member, Transaction, and Interest Rate management.
- Basic interest calculation engine.
