# HD Asset Balance

A browser-based report for calculating HD-wise asset balances from three Excel or CSV files. The app does not upload data to a server; parsing and calculations happen in the browser.

## Run

Open `index.html` in a modern browser. An internet connection is needed to load the SheetJS Excel reader from its CDN; CSV files can also be loaded once that library is available.

## Source file columns

Each workbook's first worksheet is read. Header matching ignores spaces, punctuation and capitalization, and supports the common misspellings shown here.

- **Asset raw data:** `Asset ID` (also `Asse ID`), `HD Code`, `Asset Description` (also `Asset Descipion`), `Capacity` (also `Capaciy`), and `Delivery Date` (also `Delie Date`). The first two columns are required.
- **Asset values:** `Asset ID` and `Asset Value`. Both are required.
- **HD master:** `HD Code`, `HD Name`, `Vendor Code`, `Location`, `SDM Name`, `RSM Name`, `ZSM Name`, `Region` (or `State`), `Security Deposite`, and `Total Deposite`. HD code and the two deposit columns are required; other HD master columns may be omitted.

Asset IDs are joined between the raw and values sheets. Each matching raw row counts as one asset. If the values sheet has repeated asset IDs, their values are added together and used for each matching raw row. An asset without a matching value is shown with a zero book value and reported in the data-quality notes.

**Balance = Total Deposit − (Total Asset Value + Security Deposit)**

The report includes HD search, asset detail views, aggregate totals, and CSV export.
