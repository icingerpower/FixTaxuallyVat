# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build

This is a C++17 Qt application built with CMake. The build requires Qt 5 or Qt 6 and the QXlsx library.

```bash
# Configure (from project root)
cmake -B build -S .

# Build
cmake --build build

# Run
./build/FixAmazonVat
```

QXlsx is expected at `${Qt_DIR}/../../../../lib/cmake/QXlsxQt6`. The shared CsvReader utility is pulled from `../common/utils/` via `utils.cmake`.

No automated tests exist (the `tests/` directory is gitignored).

## Architecture

**Purpose:** Desktop tool for Amazon sellers using Taxually. Compares Amazon VAT transaction CSVs against Taxually Excel reports to find discrepancies.

**Data flow:**
1. User loads Amazon VAT CSV files (one per country: DE, ES, IT, PL, CZ, UK)
2. User loads a Taxually Excel file (`.xlsx`, sheet: "Tax return detail")
3. `VatAnalyser` reads both sources, builds a transaction index, and finds mismatches
4. Results shown in a `QDialog` with a `QTableView` backed by `DifferenceTableModel`
5. Error rows can be saved as a new `-ERRORS.xlsx` with yellow highlights

**Key classes:**
- `VatAnalyser` — all core logic: CSV parsing, Excel reading (via QXlsx), transaction matching, report generation
- `DifferenceTableModel` — `QAbstractTableModel` with 9 columns (Order ID, File name, Shipment ID, Untaxed amount, Taxes, Amazon, Country from, Country to, Date tax)
- `MainWindow` — thin UI controller; manages file list widget, button connections, and launches the results dialog

**Transaction identity:** A transaction is uniquely identified by concatenating order ID + transaction date + untaxed amount (decimals normalized). This composite key is used to match Amazon rows against Taxually rows.

**Filtering rules:**
- Transactions older than 770 days are excluded (covers old-style refunds)
- Tax scheme is classified as `OSS` (Union-OSS) or `REGULAR` based on Amazon's "tax reporting scheme" column
- Country-specific column offsets in the Taxually Excel are hardcoded per country in `VatAnalyser`
