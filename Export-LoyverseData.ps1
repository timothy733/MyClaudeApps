<#
.SYNOPSIS
    Pulls current item inventory and a trailing window of receipts from the
    Loyverse API and writes them to a two-sheet Excel workbook.

.DESCRIPTION
    Sheet 1 "Current Inventory": one row per active item/variant — ISBN, SKU,
    Category, Title, Item Price, Cost Per Item, Current Stock Qty (summed
    across all stores if you have more than one).

    Sheet 2 "Receipts (Last Nd)": one row per line item per receipt — Receipt
    Number, Receipt Type, Note, Order, Book Title, Category, SKU, ISBN, Qty
    Sold, Price Per Item, Cost per Item, Payment Method, Transaction Date,
    Cancellation Date, Total Money.

    Requires the ImportExcel module (installed automatically the first time
    if it isn't already present — no Excel installation needed).

    The access token is never hardcoded in this script. It's read from a
    plain text file containing nothing but the token itself. Create that
    file once:

        "your-token-here" | Out-File -FilePath .\loyverse_token.txt -Encoding utf8 -NoNewline

    Get the token from Loyverse Back Office -> Settings -> Access Tokens ->
    "+ Add access token". Keep that file private (don't commit it, don't
    email it, don't put it anywhere shared).

.PARAMETER TokenPath
    Path to the plain-text token file. Default: loyverse_token.txt next to
    this script.

.PARAMETER DaysBack
    How many days of receipts to pull, ending now. Default: 30.

.PARAMETER OutputPath
    Where to write the .xlsx file. Default: loyverse_export_<date>.xlsx next
    to this script.

.PARAMETER TestOnly
    If set, only checks the connection (fetches your stores) and exits —
    useful to confirm the token works before running a full export.

.EXAMPLE
    .\Export-LoyverseData.ps1 -TestOnly

.EXAMPLE
    .\Export-LoyverseData.ps1 -DaysBack 60 -OutputPath "C:\Reports\export.xlsx"
#>

[CmdletBinding()]
param(
    [string]$TokenPath = (Join-Path $PSScriptRoot "loyverse_token.txt"),
    [int]$DaysBack = 30,
    [string]$OutputPath,
    [switch]$TestOnly
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'   # Invoke-RestMethod's default progress bar is very slow over many calls
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ---------------------------------------------------------------------------
# 0. Setup
# ---------------------------------------------------------------------------
if (-not $OutputPath) {
    $OutputPath = Join-Path $PSScriptRoot "loyverse_export_$(Get-Date -Format 'yyyy-MM-dd').xlsx"
}

if (-not (Test-Path -LiteralPath $TokenPath)) {
    Write-Error "Token file not found at '$TokenPath'. Create a plain text file containing only your Loyverse access token, then re-run (see -TokenPath, or the .DESCRIPTION help via Get-Help .\Export-LoyverseData.ps1 -Full)."
    exit 1
}
$Token = (Get-Content -LiteralPath $TokenPath -Raw).Trim()
if (-not $Token) {
    Write-Error "Token file '$TokenPath' is empty."
    exit 1
}

$BaseUrl = "https://api.loyverse.com/v1.0"
$Headers = @{
    Authorization = "Bearer $Token"
    Accept        = "application/json"
}

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found - installing (one-time setup)..." -ForegroundColor Yellow
    try {
        $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if ($repo -and $repo.InstallationPolicy -ne 'Trusted') {
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        }
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber
    } catch {
        Write-Error "Could not install the ImportExcel module automatically. Run this once yourself in PowerShell, then re-run this script: Install-Module -Name ImportExcel -Scope CurrentUser -Force"
        exit 1
    }
}
Import-Module ImportExcel -Force

# ---------------------------------------------------------------------------
# Helper: cursor-paginated GET against the Loyverse API
# ---------------------------------------------------------------------------
function Get-LoyversePages {
    param(
        [Parameter(Mandatory)][string]$Path,
        [hashtable]$QueryParams = @{},
        [Parameter(Mandatory)][string]$ArrayKey
    )
    $all = New-Object System.Collections.Generic.List[object]
    $cursor = $null
    do {
        $params = @{}
        foreach ($k in $QueryParams.Keys) { $params[$k] = $QueryParams[$k] }
        $params['limit'] = 250
        if ($cursor) { $params['cursor'] = $cursor }

        $pairs = foreach ($k in $params.Keys) {
            "$k=$([System.Uri]::EscapeDataString([string]$params[$k]))"
        }
        $uri = "$BaseUrl$Path`?$($pairs -join '&')"

        try {
            $resp = Invoke-RestMethod -Uri $uri -Headers $Headers -Method Get
        } catch {
            $statusCode = $null
            $body = $null
            if ($_.Exception.Response) {
                try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
                try {
                    $stream = $_.Exception.Response.GetResponseStream()
                    $reader = New-Object System.IO.StreamReader($stream)
                    $body = $reader.ReadToEnd()
                } catch {}
            }
            if (-not $body -and $_.ErrorDetails) { $body = $_.ErrorDetails.Message }
            throw "Loyverse API error on $Path (status $statusCode): $body"
        }

        $arr = $resp.$ArrayKey
        if ($arr) { foreach ($item in @($arr)) { $all.Add($item) } }
        $cursor = $resp.cursor
        Write-Host "  ...$($all.Count) records from $Path so far" -ForegroundColor DarkGray
    } while ($cursor)
    return $all
}

function ConvertTo-IsbnNumber {
    # Excel's number format only visually applies to cells holding an actual
    # number - a text string just sits there unformatted. ISBN-13 barcodes
    # fit safely inside an int64 with no precision loss, so convert when we
    # can and leave blanks blank rather than turning them into 0.
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $trimmed = $Value.Trim()
    $parsed = 0L
    if ([int64]::TryParse($trimmed, [ref]$parsed)) { return $parsed }
    return $Value   # non-numeric barcode (rare) - fall back to the raw string
}

function Test-LoyverseConnection {
    Write-Host "Testing connection..." -ForegroundColor Cyan
    $stores = Get-LoyversePages -Path "/stores" -ArrayKey "stores"
    $names = ($stores | ForEach-Object { $_.name }) -join ", "
    Write-Host "Connected. Stores visible to this token: $names" -ForegroundColor Green
    return $stores
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
try {
    if ($TestOnly) {
        Test-LoyverseConnection | Out-Null
        exit 0
    }

    Write-Host "Fetching categories..." -ForegroundColor Cyan
    $categories = Get-LoyversePages -Path "/categories" -ArrayKey "categories"
    $categoryNameById = @{}
    foreach ($c in $categories) { $categoryNameById[$c.id] = $c.name }
    Write-Host "Loaded $($categories.Count) categories." -ForegroundColor Green

    Write-Host "Fetching stores..." -ForegroundColor Cyan
    $stores = Get-LoyversePages -Path "/stores" -ArrayKey "stores"
    $storeNames = ($stores | ForEach-Object { $_.name }) -join ", "
    Write-Host "Loaded $($stores.Count) store(s): $storeNames" -ForegroundColor Green

    Write-Host "Fetching items..." -ForegroundColor Cyan
    $items = Get-LoyversePages -Path "/items" -QueryParams @{ show_deleted = "true" } -ArrayKey "items"
    Write-Host "Loaded $($items.Count) items." -ForegroundColor Green

    # variant_id -> info (kept for deleted items too, so old receipts still resolve)
    $variantInfo = @{}
    foreach ($item in $items) {
        $deleted = [bool]$item.deleted_at
        foreach ($v in $item.variants) {
            $titleParts = @($v.option1_value, $v.option2_value, $v.option3_value) | Where-Object { $_ }
            $title = $item.item_name
            if ($titleParts.Count -gt 0) { $title = "$($item.item_name) ($($titleParts -join ' / '))" }

            $price = $v.default_price
            if ($null -eq $price -and $v.stores -and $v.stores.Count -gt 0) { $price = $v.stores[0].price }

            $variantInfo[$v.variant_id] = [PSCustomObject]@{
                SKU        = $v.sku
                Barcode    = $v.barcode
                Cost       = $v.cost
                Price      = $price
                Title      = $title
                CategoryId = $item.category_id
                Deleted    = $deleted
            }
        }
    }

    Write-Host "Fetching inventory levels..." -ForegroundColor Cyan
    $stockByVariant = @{}
    foreach ($store in $stores) {
        $levels = Get-LoyversePages -Path "/inventory" -QueryParams @{ store_id = $store.id } -ArrayKey "inventory_levels"
        foreach ($lvl in $levels) {
            $current = 0
            if ($stockByVariant.ContainsKey($lvl.variant_id)) { $current = $stockByVariant[$lvl.variant_id] }
            $inStock = 0
            if ($null -ne $lvl.in_stock) { $inStock = $lvl.in_stock }
            $stockByVariant[$lvl.variant_id] = $current + $inStock
        }
    }
    Write-Host "Inventory levels loaded." -ForegroundColor Green

    Write-Host "Fetching payment types..." -ForegroundColor Cyan
    $paymentTypes = Get-LoyversePages -Path "/payment_types" -ArrayKey "payment_types"
    $paymentNameById = @{}
    foreach ($p in $paymentTypes) { $paymentNameById[$p.id] = $p.name }
    Write-Host "Loaded $($paymentTypes.Count) payment types." -ForegroundColor Green

    # ---- Sheet 1: Current Inventory (active items only) ----
    $inventoryRows = foreach ($vid in $variantInfo.Keys) {
        $info = $variantInfo[$vid]
        if ($info.Deleted) { continue }
        $stock = 0
        if ($stockByVariant.ContainsKey($vid)) { $stock = $stockByVariant[$vid] }
        $category = ""
        if ($categoryNameById.ContainsKey($info.CategoryId)) { $category = $categoryNameById[$info.CategoryId] }
        [PSCustomObject]@{
            "ISBN"              = (ConvertTo-IsbnNumber $info.Barcode)
            "SKU"               = $info.SKU
            "Category"          = $category
            "Title"             = $info.Title
            "Item Price"        = $info.Price
            "Cost Per Item"     = $info.Cost
            "Current Stock Qty" = $stock
        }
    }
    Write-Host "Built inventory sheet: $($inventoryRows.Count) rows." -ForegroundColor Green

    # ---- Sheet 2: Receipts, flattened to one row per line item ----
    $maxDate = Get-Date
    $minDate = $maxDate.AddDays(-1 * $DaysBack)
    $minIso  = $minDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    $maxIso  = $maxDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    Write-Host "Fetching receipts from $($minDate.ToString('yyyy-MM-dd')) to $($maxDate.ToString('yyyy-MM-dd'))..." -ForegroundColor Cyan
    $receipts = Get-LoyversePages -Path "/receipts" -QueryParams @{ created_at_min = $minIso; created_at_max = $maxIso } -ArrayKey "receipts"
    Write-Host "Loaded $($receipts.Count) receipts." -ForegroundColor Green

    $receiptRows = foreach ($r in $receipts) {
        $paymentNames = @()
        foreach ($p in $r.payments) {
            if ($paymentNameById.ContainsKey($p.payment_type_id)) { $paymentNames += $paymentNameById[$p.payment_type_id] }
            else { $paymentNames += $p.payment_type_id }
        }
        $paymentStr = $paymentNames -join ", "

        $createdStr = ""
        if ($r.created_at) { $createdStr = ([datetime]$r.created_at).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") }
        $cancelledStr = ""
        if ($r.cancelled_at) { $cancelledStr = ([datetime]$r.cancelled_at).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") }

        foreach ($li in $r.line_items) {
            $vinfo = $null
            if ($variantInfo.ContainsKey($li.variant_id)) { $vinfo = $variantInfo[$li.variant_id] }

            $category  = ""
            $isbn      = ""
            $sku       = $li.sku
            $bookTitle = $li.item_name
            if ($vinfo) {
                if (-not $sku) { $sku = $vinfo.SKU }
                $isbn = $vinfo.Barcode
                if (-not $bookTitle) { $bookTitle = $vinfo.Title }
                if ($categoryNameById.ContainsKey($vinfo.CategoryId)) { $category = $categoryNameById[$vinfo.CategoryId] }
            }

            [PSCustomObject]@{
                "Receipt Number"    = $r.receipt_number
                "Receipt Type"      = $r.receipt_type
                "Note"              = $r.note
                "Order"             = $r.order
                "Book Title"        = $bookTitle
                "Category"          = $category
                "SKU"               = $sku
                "ISBN"              = (ConvertTo-IsbnNumber $isbn)
                "Qty Sold"          = $li.quantity
                "Price Per Item"    = $li.price
                "Cost per Item"     = $li.cost
                "Payment Method"    = $paymentStr
                "Transaction Date"  = $createdStr
                "Cancellation Date" = $cancelledStr
                "Total Money"       = $r.total_money
            }
        }
    }
    Write-Host "Built receipts sheet: $($receiptRows.Count) line-item rows." -ForegroundColor Green

    # ---- Write the workbook ----
    if (Test-Path -LiteralPath $OutputPath) { Remove-Item -LiteralPath $OutputPath -Force }
    $sheet1Name = "Current Inventory"
    $sheet2Name = "Receipts (Last $($DaysBack)d)"

    # -PassThru keeps the package open (unsaved) so we can format columns before writing to disk
    $pkg = $inventoryRows | Export-Excel -Path $OutputPath -WorksheetName $sheet1Name -AutoSize -BoldTopRow -FreezeTopRow -PassThru
    $pkg = $receiptRows   | Export-Excel -ExcelPackage $pkg -WorksheetName $sheet2Name -AutoSize -BoldTopRow -FreezeTopRow -PassThru

    # Format the ISBN column as a whole number (0 decimal places) on both sheets.
    # Found by header text rather than a hardcoded column letter, so it keeps working if columns are reordered later.
    foreach ($sheetName in @($sheet1Name, $sheet2Name)) {
        $ws = $pkg.Workbook.Worksheets[$sheetName]
        if ($ws.Dimension) {
            $lastRow = $ws.Dimension.End.Row
            $lastCol = $ws.Dimension.End.Column
            for ($col = 1; $col -le $lastCol; $col++) {
                if ($ws.Cells[1, $col].Text -eq "ISBN") {
                    if ($lastRow -gt 1) {
                        $ws.Cells[2, $col, $lastRow, $col].Style.Numberformat.Format = "0"
                    }
                    break
                }
            }
        }
    }

    Close-ExcelPackage $pkg

    Write-Host "Done. Wrote $OutputPath" -ForegroundColor Green
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
