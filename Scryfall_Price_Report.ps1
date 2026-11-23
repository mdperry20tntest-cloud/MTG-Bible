# ==============================================================================
# CONFIGURATION & VARIABLES
# ==============================================================================
$Aud         = 1.4233     # fallback USD->AUD rate, only used if the live rate can't be fetched
$UseLiveRate = $true      # $true = fetch today's USD->AUD rate each run; $false = always use the fixed rate above
$OutputPath  = "D:\MTG\Scryfall_Price_Report.html"
$CsvPath     = "D:\MTG\ManaBox_Collection.csv"
$SnapshotPath = "D:\MTG\Price_Snapshot.json"   # remembers prices between runs for the Change column
$HistoryPath  = "D:\MTG\Price_History.json"    # one total-collection-value entry per day, for the value-over-time chart
$ReportTitle = "TCG Bible"

# Total amount you have spent on the collection, in AUD. Shown as a "Total Spent" pill next to the Total pill:
# green when the collection is worth MORE than you spent, red when it is worth LESS. Set to 0 to hide the pill.
$TotalSpent  = 0

# The script's own folder. Columns.ps1 and the Assets folder (report.html / report.css / report.js) live next to it.
$ScriptDir  = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$AssetsPath = Join-Path $ScriptDir 'Assets'

# eBay Search column: the icon opens an eBay search for the card name plus this extra word
# (keeps results to Magic cards). Set $EbaySearchSuffix to '' to search by the card name only.
$EbayDomain       = 'www.ebay.com.au'
$EbaySearchSuffix = 'mtg'

# $true = only show eBay listings located in Australia (adds eBay's "Australia Only" filter to both eBay links).
$EbayAustraliaOnly = $true
$EbayLocationParam = if ($EbayAustraliaOnly) { '&LH_PrefLoc=1' } else { '' }

# $true = open the finished report in your default browser as soon as the script finishes.
$AutoOpenReport = $true

# Rows whose value in this column is above the minimum get a bold gold border.
# The column must be a number column (Type = 'number'); the minimum is in AUD, as shown in the report.
# Set $HighValueColumn to '' to turn the highlight off.
$HighValueColumn = 'Price'
$HighValueMin    = 200

# ==============================================================================
# REFRESH BUTTON SUPPORT (registers the tcgrefresh:// link for the current user)
# ==============================================================================
# Runs every time the script runs, so moving/renaming the script just needs one
# manual run to re-point the link. No admin rights needed (HKCU only).
if ($PSCommandPath) {
    try {
        $ProtoKey = "HKCU:\Software\Classes\tcgrefresh"
        New-Item -Path "$ProtoKey\shell\open\command" -Force | Out-Null
        Set-ItemProperty -Path $ProtoKey -Name "(Default)" -Value "URL:TCG Refresh"
        Set-ItemProperty -Path $ProtoKey -Name "URL Protocol" -Value ""
        Set-ItemProperty -Path "$ProtoKey\shell\open\command" -Name "(Default)" `
            -Value "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    }
    catch {
        Write-Warning "Could not register the Refresh button handler: $($_.Exception.Message)"
    }
}

# ==============================================================================
# LIVE EXCHANGE RATE (USD -> AUD)
# ==============================================================================
# Fetched from the free Frankfurter API (no key needed). If the request fails or returns
# something odd, the fixed $Aud rate from the configuration section is used instead.
function Get-LiveUsdAudRate {
    $urls = @(
        'https://api.frankfurter.dev/v2/rate/usd/aud',
        'https://api.frankfurter.dev/v1/latest?base=USD&symbols=AUD'
    )
    foreach ($url in $urls) {
        try {
            $r = Invoke-RestMethod -Uri $url -TimeoutSec 15 -Headers @{ Accept = 'application/json' }
            $rate = 0.0
            if ($r.rate)                     { $rate = [double]$r.rate }
            elseif ($r.rates -and $r.rates.AUD) { $rate = [double]$r.rates.AUD }
            if ($rate -gt 0.5 -and $rate -lt 5) { return $rate }      # sanity check
        }
        catch { }
    }
    return $null
}

$AudSource = 'fixed'
if ($UseLiveRate) {
    Write-Host "Fetching live USD to AUD exchange rate..." -ForegroundColor Cyan
    $liveRate = Get-LiveUsdAudRate
    if ($liveRate) {
        $Aud       = [math]::Round($liveRate, 4)
        $AudSource = 'live'
        Write-Host "Using live rate: 1 USD = $Aud AUD" -ForegroundColor Cyan
    }
    else {
        Write-Warning "Could not fetch a live exchange rate - using the fixed rate of $Aud instead."
    }
}

# ==============================================================================
# PRICE HELPER (used by the Price and Total columns)
# ==============================================================================
# Raw Scryfall price text for the finish ($null when Scryfall has no price for it).
# Foil-only printings (e.g. Galaxy Foil Ancient Tomb, Stellar Sights #91) have no nonfoil version, so a row marked
# 'normal' in the CSV is treated as foil for them. Without this they would show 'No price'.
function Get-EffectiveFinish {
    param($card, [string]$finish)
    if ($finish -eq 'normal' -and $card.finishes -and ($card.finishes -notcontains 'nonfoil') -and ($card.finishes -contains 'foil')) { return 'foil' }
    return $finish
}

function Get-RawPriceUsd {
    param($card, [string]$finish)
    $finish = Get-EffectiveFinish $card $finish
    switch ($finish) {
        'foil'   { if ($card.prices.usd_foil) { return $card.prices.usd_foil } else { return $card.prices.usd } }
        'etched' { return $card.prices.usd_etched }
        default  { return $card.prices.usd }
    }
}

# $true if Scryfall has a price for this card in this finish.
function Test-HasPrice {
    param($card, [string]$finish)
    return [bool](Get-RawPriceUsd $card $finish)
}

# Returns the price of ONE copy in USD, based on the finish (0 when there is no price).
function Get-UnitPriceUsd {
    param($card, [string]$finish)
    return [double](Get-RawPriceUsd $card $finish)
}

# Same price converted to AUD.
function Get-UnitPriceAud {
    param($card, [string]$finish)
    return (Get-UnitPriceUsd $card $finish) * $Aud
}

# ==============================================================================
# COLUMN DEFINITIONS  <-- live in Columns.ps1 (the file to edit to add/change columns)
# ==============================================================================
$ColumnsFile = Join-Path $ScriptDir 'Columns.ps1'
if (-not (Test-Path -Path $ColumnsFile)) {
    Write-Error "Could not find Columns.ps1 at: $ColumnsFile (it must sit next to this script)"
    return
}
. $ColumnsFile      # dot-sourced: defines $Columns using the helpers and settings above

# ==============================================================================
# INPUT VALIDATION
# ==============================================================================
if (-not (Test-Path -Path $CsvPath)) {
    Write-Error "Could not find your source CSV file at: $CsvPath"
    return
}

Write-Host "Importing list and querying Scryfall API..." -ForegroundColor Cyan
$idnumbers = Import-Csv -Path $CsvPath

# ==============================================================================
# DATA COLLECTION (BATCHED API QUERIES)
# ==============================================================================
# Uses Scryfall's /cards/collection endpoint: up to 75 cards per request, so
# 5000 cards need ~67 requests instead of 5000.
$BatchSize    = 75     # Scryfall's maximum per request - do not raise
$BatchDelayMs = 600    # polite pause between batch requests

$headers1 = @{
    "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) MyCustomPowerShellScript/1.0"
    "Accept"     = "application/json"
}

function Invoke-ScryfallCollection {
    param([string[]]$Ids)

    $body = @{ identifiers = @($Ids | ForEach-Object { @{ id = $_ } }) } | ConvertTo-Json -Depth 4 -Compress

    for ($attempt = 1; $attempt -le 4; $attempt++) {
        try {
            return Invoke-RestMethod -Uri "https://api.scryfall.com/cards/collection" -Method Post `
                -Headers $headers1 -ContentType "application/json" -Body $body
        }
        catch {
            $status = 0
            if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
            $retryable = ($status -eq 429) -or ($status -ge 500)
            if (-not $retryable -or $attempt -eq 4) { throw }
            Write-Warning "Scryfall returned HTTP $status; retrying in $($attempt * 2) seconds..."
            Start-Sleep -Seconds ($attempt * 2)
        }
    }
}

# 0. Load the price snapshot from earlier runs (used by the Change column).
#    Prices are compared against the last run from a PREVIOUS day, so running the script
#    several times in one day keeps comparing against the same baseline instead of showing zero change.
#    Snapshot file layout: today's prices, plus the baseline prices they are compared with.
#    The date is stored as yyyyMMdd text so JSON parsing never turns it into a DateTime.
$TodayStamp   = Get-Date -Format "yyyyMMdd"
$SnapPrices   = @{}     # "id|finish" -> USD price of one copy, as saved by the most recent run
$PrevPrices   = @{}     # the baseline prices the Change column compares against
$BaselineDate = ''

if (Test-Path -Path $SnapshotPath) {
    try {
        $snap = Get-Content -Path $SnapshotPath -Raw | ConvertFrom-Json
        if ($snap.prices) {
            foreach ($p in $snap.prices.PSObject.Properties) { $SnapPrices[$p.Name] = [double]$p.Value }
        }
        if ([string]$snap.date -eq $TodayStamp) {
            $baselineSource = $snap.baseline
            $BaselineDate   = [string]$snap.baselineDate
        }
        else {
            $baselineSource = $snap.prices
            $BaselineDate   = [string]$snap.date
        }
        if ($baselineSource) {
            foreach ($p in $baselineSource.PSObject.Properties) { $PrevPrices[$p.Name] = [double]$p.Value }
        }
    }
    catch {
        Write-Warning "Could not read the price snapshot ($SnapshotPath): $($_.Exception.Message). Change column will be blank."
        $SnapPrices = @{}
        $PrevPrices = @{}
        $BaselineDate = ''
    }
}
if ($PrevPrices.Count -eq 0) { $BaselineDate = '' }

# 1. Read and validate the CSV rows (keeps CSV order and duplicates, so owning 3 copies = 3 rows).
#    Each row remembers its finish (Foil column), because the normal and foil copies
#    of the same card have different prices.
$CsvRows = @($idnumbers)
if ($CsvRows.Count -gt 0 -and -not ($CsvRows[0].PSObject.Properties.Name -contains 'Foil')) {
    Write-Warning "No 'Foil' column found in the CSV - every card will be priced as normal."
}

$QtyColumn = $null
if ($CsvRows.Count -gt 0) {
    $QtyColumn = $CsvRows[0].PSObject.Properties.Name | Where-Object { $_ -in @('Quantity', 'Qty') } | Select-Object -First 1
    if (-not $QtyColumn) {
        Write-Warning "No 'Quantity' column found in the CSV - every row will be counted as 1 copy."
    }
}

$OrderedEntries = New-Object System.Collections.Generic.List[PSCustomObject]
$VariantsById   = @{}      # card id -> set of "finish|quantity" combinations wanted for that card
$UnknownFinish  = @{}
$BadQuantity    = 0
$InvalidIdCount = 0        # CSV rows skipped because the Scryfall ID was not a valid ID
$FailedBatches  = 0        # Scryfall batch requests that failed completely

foreach ($csvRow in $CsvRows) {
    $raw = $csvRow.'Scryfall ID'
    if ([string]::IsNullOrWhiteSpace($raw)) { continue }

    $parsed = [guid]::Empty
    if (-not [guid]::TryParse($raw.Trim(), [ref]$parsed)) {
        Write-Warning "Skipping invalid Scryfall ID: $($raw.Trim())"
        $InvalidIdCount++
        continue
    }
    $cardId = $parsed.ToString()          # normalised: lowercase, hyphenated

    $foilText = ([string]$csvRow.Foil).Trim().ToLower()
    if ($foilText -match 'etched') { $finish = 'etched' }
    elseif ($foilText -match 'foil' -and $foilText -notin @('nonfoil', 'non-foil', 'non foil')) { $finish = 'foil' }
    else {
        $finish = 'normal'
        if ($foilText -notin @('', 'normal', 'nonfoil', 'non-foil')) { $UnknownFinish[$foilText] = $true }
    }

    # Quantity: blank means 1; anything that is not a whole number >= 0 is treated as 1 (with a warning)
    $qty = 1
    if ($QtyColumn) {
        $qtyText = ([string]$csvRow.$QtyColumn).Trim()
        if ($qtyText -ne '') {
            $qtyParsed = 0
            if ([int]::TryParse($qtyText, [ref]$qtyParsed) -and $qtyParsed -ge 0) { $qty = $qtyParsed }
            else { $BadQuantity++ }
        }
    }

    # Date added (from the CSV's Added column), kept as yyyy-MM-dd so it sorts correctly as text
    $added = ''
    if (([string]$csvRow.Added).Trim() -match '^(\d{4}-\d{2}-\d{2})') { $added = $Matches[1] }

    $OrderedEntries.Add([PSCustomObject]@{ Id = $cardId; Finish = $finish; Qty = $qty; Added = $added })
    if (-not $VariantsById.ContainsKey($cardId)) {
        $VariantsById[$cardId] = New-Object System.Collections.Generic.HashSet[string]
    }
    [void]$VariantsById[$cardId].Add("$finish|$qty|$added")
}

if ($BadQuantity -gt 0) {
    Write-Warning "$BadQuantity row(s) had an invalid Quantity and were counted as 1."
}

if ($UnknownFinish.Count -gt 0) {
    Write-Warning ("Unrecognised Foil value(s) treated as normal: " +
        (($UnknownFinish.Keys | ForEach-Object { "'$_'" }) -join ', '))
}

$UniqueIds = @($OrderedEntries | ForEach-Object { $_.Id } | Select-Object -Unique)
Write-Host "Found $($OrderedEntries.Count) rows ($($UniqueIds.Count) unique cards)." -ForegroundColor Cyan

# 2. Fetch unique IDs in batches. Rows are built immediately (one per finish wanted)
#    so the large raw API objects can be discarded (keeps memory low).
$RowLookup  = @{}      # "id|finish|qty|added" -> row
$FetchedIds = @{}
$CurrentPrices = @{}   # "id|finish" -> USD price of one copy (saved as the next snapshot)
$BatchCount = [math]::Ceiling($UniqueIds.Count / $BatchSize)

for ($b = 0; $b -lt $BatchCount; $b++) {
    $first = $b * $BatchSize
    $last  = [math]::Min($first + $BatchSize, $UniqueIds.Count) - 1
    $batch = $UniqueIds[$first..$last]

    Write-Progress -Activity "Querying Scryfall" -Status "Batch $($b + 1) of $BatchCount" `
        -PercentComplete ([int](($b / $BatchCount) * 100))

    try {
        $response = Invoke-ScryfallCollection -Ids $batch
        foreach ($card in $response.data) {
            $cardId = ([string]$card.id).ToLower()
            $FetchedIds[$cardId] = $true
            foreach ($variant in $VariantsById[$cardId]) {
                $finish, $qtyText, $added = $variant -split '\|'
                $qty = [int]$qtyText
                if (Test-HasPrice $card $finish) { $CurrentPrices["$cardId|$finish"] = Get-UnitPriceUsd $card $finish }
                $row = [ordered]@{}
                foreach ($col in $Columns) {
                    $row[$col.Key] = & $col.Value $card ($card.id) $finish $qty $added
                }
                $RowLookup["$cardId|$finish|$qty|$added"] = [PSCustomObject]$row
            }
        }
    }
    catch {
        Write-Warning "Batch $($b + 1) of $BatchCount failed: $($_.Exception.Message)"
        $FailedBatches++
    }

    if ($b -lt $BatchCount - 1) { Start-Sleep -Milliseconds $BatchDelayMs }
}
Write-Progress -Activity "Querying Scryfall" -Completed

# 3. Report anything that did not come back
$Missing = @($UniqueIds | Where-Object { -not $FetchedIds.ContainsKey($_) })
if ($Missing.Count -gt 0) {
    Write-Warning ("$($Missing.Count) Scryfall ID(s) could not be fetched (not found or batch failed). First few: " +
        (($Missing | Select-Object -First 10) -join ', '))
}

# 4. Assemble rows in original CSV order
$CardList = New-Object System.Collections.Generic.List[PSCustomObject]
foreach ($entry in $OrderedEntries) {
    $key = "$($entry.Id)|$($entry.Finish)|$($entry.Qty)|$($entry.Added)"
    if ($RowLookup.ContainsKey($key)) { $CardList.Add($RowLookup[$key]) }
}

# 5. Work out what to warn about, both in this window and in a banner at the top of the report.
$NoPriceCount = @($CardList | Where-Object { $_.Price -eq 'No price' }).Count
if ($NoPriceCount -gt 0) {
    Write-Warning "$NoPriceCount row(s) have no Scryfall price for their finish and are shown as 'No price' (counted as `$0.00 in totals)."
}

$BannerLines = New-Object System.Collections.Generic.List[string]
if ($Missing.Count -gt 0) {
    $firstFew = [System.Net.WebUtility]::HtmlEncode((($Missing | Select-Object -First 5) -join ', '))
    $reason   = if ($FailedBatches -gt 0) { "$FailedBatches lookup batch(es) failed, and some IDs may not exist" } else { "Scryfall did not recognise them" }
    $BannerLines.Add("<strong>$($Missing.Count) of $($UniqueIds.Count) unique cards could not be loaded from Scryfall</strong> ($reason). They are missing from this report, so the table and totals are incomplete. Click Refresh data to retry. First few IDs: $firstFew")
}
if ($InvalidIdCount -gt 0) {
    $BannerLines.Add("<strong>$InvalidIdCount CSV row(s) were skipped</strong> because the Scryfall ID was not valid.")
}
$BannerHtml = ''
if ($BannerLines.Count -gt 0) {
    $BannerHtml = "<div class='warning-banner'>" + (($BannerLines | ForEach-Object { "<div>$_</div>" }) -join '') + "</div>"
}

# 6. Save the snapshot for the next run. Cards that failed to fetch this time keep their old saved price.
if ($CurrentPrices.Count -gt 0) {
    try {
        foreach ($k in $CurrentPrices.Keys) { $SnapPrices[$k] = $CurrentPrices[$k] }
        [ordered]@{
            date         = $TodayStamp
            prices       = $SnapPrices
            baselineDate = $BaselineDate
            baseline     = $PrevPrices
        } | ConvertTo-Json -Depth 4 -Compress | Out-File -FilePath $SnapshotPath -Encoding utf8 -Force
    }
    catch {
        Write-Warning "Could not save the price snapshot ($SnapshotPath): $($_.Exception.Message)"
    }
}

# 7. Collection value history: one entry per day (re-running on the same day replaces that day's entry).
#    The total is stored in USD so exchange-rate moves do not look like price moves; the report converts it
#    at today's rate. A day is NOT recorded if any cards failed to load, because the total would be too low
#    and would show up as a false dip in the chart.
#
#    The file is READ with a regex and WRITTEN by hand (no ConvertFrom-Json / ConvertTo-Json), because
#    Windows PowerShell 5.1 wraps arrays as {"value":[...],"Count":n}, which corrupted this file on every run.
#    The regex only picks out well-formed {"date":"yyyyMMdd","usd":number} entries, so an old corrupted file
#    is repaired automatically and its real entries are kept.
$Inv        = [System.Globalization.CultureInfo]::InvariantCulture
$HistoryMap = [ordered]@{}     # yyyyMMdd -> collection total in USD

if (Test-Path -Path $HistoryPath) {
    try {
        $histText = Get-Content -Path $HistoryPath -Raw
        foreach ($m in [regex]::Matches($histText, '"date"\s*:\s*"(\d{8})"\s*,\s*"usd"\s*:\s*([0-9.]+)')) {
            $HistoryMap[$m.Groups[1].Value] = [double]::Parse($m.Groups[2].Value, $Inv)
        }
    }
    catch {
        Write-Warning "Could not read the value history ($HistoryPath): $($_.Exception.Message)"
    }
}

if ($CardList.Count -gt 0 -and $Missing.Count -eq 0 -and $FailedBatches -eq 0) {
    $TotalAudToday = 0.0
    foreach ($row in $CardList) {
        if ($row.LineTotal -ne 'No price') { $TotalAudToday += [double](([string]$row.LineTotal) -replace ',', '') }
    }
    $HistoryMap[$TodayStamp] = [math]::Round($TotalAudToday / $Aud, 2)
}
elseif ($CardList.Count -gt 0) {
    Write-Warning "Today's collection value was NOT added to the history because some cards could not be loaded."
}

$HistoryItems = @($HistoryMap.Keys | Sort-Object | ForEach-Object {
    '{{"date":"{0}","usd":{1}}}' -f $_, $HistoryMap[$_].ToString('0.00', $Inv)
})
$JsonHistory = '[' + ($HistoryItems -join ',') + ']'      # always a flat array, e.g. [{"date":"20261030","usd":3680.60}]

if ($HistoryMap.Count -gt 0) {
    try { $JsonHistory | Out-File -FilePath $HistoryPath -Encoding utf8 -Force }
    catch { Write-Warning "Could not save the value history ($HistoryPath): $($_.Exception.Message)" }
}

# ==============================================================================
# HTML ENGINE & REPORT GENERATION
# ==============================================================================
# The page itself lives in the Assets folder:
#   report.html  page layout, with {{PLACEHOLDERS}} filled in below
#   report.css   all styling
#   report.js    all page behaviour (table, search, sort, chart, Excel export)
# This section only prepares the data and joins the pieces into ONE self-contained HTML file.
function Read-Asset {
    param([string]$Name)
    $path = Join-Path $AssetsPath $Name
    if (-not (Test-Path -Path $path)) { throw "Missing asset file: $path" }
    return [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
}

if ($CardList.Count -gt 0) {
    Write-Host "Compiling data and assembling sortable HTML report..." -ForegroundColor Cyan

    # Data and column metadata are both passed to the page as JSON.
    # '</' is escaped so a card name can never close the <script> tag early.
    $JsonCardsData = (ConvertTo-Json @($CardList) -Compress) -replace '</', '<\/'

    $ColumnMeta = foreach ($col in $Columns) {
        [ordered]@{
            key        = $col.Key
            header     = $col.Header
            type       = if ($col.Type) { $col.Type } else { 'text' }
            prefix     = if ($col.Prefix) { $col.Prefix } else { '' }
            searchable = [bool]$col.Searchable
            total      = [bool]$col.Total
            badge      = [bool]$col.Badge
            delta      = [bool]$col.Delta
            deltaTotal = [bool]$col.DeltaTotal
        }
    }
    $JsonColumns = (ConvertTo-Json @($ColumnMeta) -Compress) -replace '</', '<\/'

    $SafeTitle   = [System.Net.WebUtility]::HtmlEncode($ReportTitle)
    $AudText     = $Aud.ToString("0.0000", [System.Globalization.CultureInfo]::InvariantCulture)
    $GeneratedOn = Get-Date -Format "d MMM yyyy"
    $AudNote     = if ($AudSource -eq 'live') { ' (live rate)' } else { ' (fixed rate)' }
    $ChangeNote  = ''
    if ($BaselineDate) {
        $baselineShown = [datetime]::ParseExact($BaselineDate, 'yyyyMMdd', [System.Globalization.CultureInfo]::InvariantCulture).ToString("d MMM yyyy")
        $ChangeNote  = " &middot; Change compared with prices from $baselineShown"
    }
    $FooterHtml = "Prices from Scryfall (USD), converted to AUD at 1 USD = $AudText AUD$AudNote &middot; Generated $GeneratedOn$ChangeNote<br>Unofficial fan-made tool, not affiliated with or endorsed by Wizards of the Coast. Magic: The Gathering is a trademark of Wizards of the Coast LLC. Card data and prices courtesy of Scryfall."

    $TotalSpentText   = ([double]$TotalSpent).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $HighValueMinText = ([double]$HighValueMin).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $HighValueKeyJs   = ($HighValueColumn -replace "[\\']", '')

    # Everything the page needs from PowerShell, as JavaScript variables.
    # ($JsonHistory was built in section 7 as a flat, hand-written JSON array.)
    $DataJs = @(
        "var cardData = $JsonCardsData;"
        "var columns = $JsonColumns;"
        "var highValueKey = '$HighValueKeyJs';"
        "var highValueMin = $HighValueMinText;"
        "var valueHistory = $JsonHistory;"
        "var audRate = $AudText;"
        "var totalSpent = $TotalSpentText;"
    ) -join "`n"

    try {
        $Html = Read-Asset 'report.html'
        $Css  = Read-Asset 'report.css'
        $Js   = Read-Asset 'report.js'
    }
    catch {
        Write-Error $_.Exception.Message
        return
    }

    # String.Replace is literal (no regex), and the big pieces go in last so card names can never be re-processed.
    $Html = $Html.Replace('{{TITLE}}',    $SafeTitle).
                  Replace('{{WARNINGS}}', $BannerHtml).
                  Replace('{{FOOTER}}',   $FooterHtml).
                  Replace('{{CSS}}',      $Css).
                  Replace('{{JS}}',       $Js).
                  Replace('{{DATA}}',     $DataJs)

    # Output file deployment
    $Html | Out-File -FilePath $OutputPath -Encoding utf8 -Force
    Write-Host "Interactive secure HTML report successfully generated at: $OutputPath" -ForegroundColor Green
    if ($AutoOpenReport) {
        try { Start-Process -FilePath $OutputPath }
        catch { Write-Warning "Could not open the report automatically: $($_.Exception.Message)" }
    }
} else {
    Write-Error "Process failed: No valid data objects were collected from the Scryfall API lookup."
}