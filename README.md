# MyClaudeApps
Small tools that help me in day to day tasks I would like to share.

## Loyverse Data Export

A single PowerShell script that pulls your **current item inventory** and a trailing window of **receipts** straight from the Loyverse API into a two-sheet Excel workbook. 

### What you get

**Sheet 1 — Current Inventory** (one row per active item)
| ISBN | SKU | Category | Title | Item Price | Cost Per Item | Current Stock Qty |
|---|---|---|---|---|---|---|

**Sheet 2 — Receipts (Last N days)** (one row per line item per receipt)
| Receipt Number | Receipt Type | Note | Order | Book Title | Category | SKU | ISBN | Qty Sold | Price Per Item | Cost per Item | Payment Method | Transaction Date | Cancellation Date | Total Money |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|

If you run multiple stores under one Loyverse account, Current Stock Qty is the sum across all of them. The ISBN column is formatted as a plain whole number (no decimals) on both sheets.

### Requirements

- Windows PowerShell 5.1+ or PowerShell 7+
- A Loyverse account with API access enabled
- Internet access to `api.loyverse.com` (and, on first run only, the PowerShell Gallery to install a module)

No Excel installation is required — the workbook is written directly via the [ImportExcel](https://github.com/dfinke/ImportExcel) module, which the script installs automatically the first time if it isn't already present.

### Setup

1. **Download** `Export-LoyverseData.ps1`.
2. **Create an access token**: Loyverse Back Office → Settings → Access Tokens → **+ Add access token**.
3. **Save the token** into a plain text file named `loyverse_token.txt`, in the same folder as the script — nothing else in the file, just the token:

   ```powershell
   "your-token-here" | Out-File -FilePath .\loyverse_token.txt -Encoding utf8 -NoNewline
   ```

4. If this folder is a git repo, add the token file to `.gitignore` so it never gets committed:

   ```
   loyverse_token.txt
   *.xlsx
   ```

> **Security note:** a Loyverse access token grants full read/write access to your account. Keep `loyverse_token.txt` private — don't commit it, email it, or put it anywhere shared.

### Usage

```powershell
# Confirm the token and connection work before running a full export
.\Export-LoyverseData.ps1 -TestOnly

# Default run: last 30 days of receipts, output saved next to the script
.\Export-LoyverseData.ps1

# Custom window and output location
.\Export-LoyverseData.ps1 -DaysBack 60 -OutputPath "C:\Reports\loyverse_export.xlsx"
```

### Parameters

| Parameter | Default | Description |
|---|---|---|
| `-TokenPath` | `loyverse_token.txt` next to the script | Path to the plain-text token file |
| `-DaysBack` | `30` | Number of days of receipts to pull, ending now |
| `-OutputPath` | `loyverse_export_<date>.xlsx` next to the script | Where to write the workbook |
| `-TestOnly` | *(switch)* | Only checks the connection (fetches your stores) and exits |

Full parameter help, including examples, is available via:
```powershell
Get-Help .\Export-LoyverseData.ps1 -Full
```

### How it works

The script calls Loyverse's REST API (`/categories`, `/stores`, `/items`, `/inventory`, `/payment_types`, `/receipts`), paginating through each endpoint's cursor until everything in range is collected, then joins items/categories back onto receipt line items to fill in Category, SKU, and ISBN (which aren't included on the receipt itself). Because this runs as a server-to-server HTTPS call rather than a browser request, it isn't subject to the CORS restrictions that block a browser-based version of the same tool.

### A note on accuracy

A few details of Loyverse's API response shape (e.g. whether receipt line items include an `sku` field directly, the exact `/inventory` response structure) were taken from Loyverse's public documentation rather than verified against a live account for every field. If a run errors out partway through, the script prints the raw API response body in the error message — please open an issue with that text and the endpoint it failed on, and the field mapping can be corrected.
No license file is included yet — add one (e.g. [MIT](https://choosealicense.com/licenses/mit/)) before publishing if you want others to be able to reuse this freely.
