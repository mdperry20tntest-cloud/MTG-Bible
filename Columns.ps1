# ==============================================================================
# COLUMN DEFINITIONS  <-- THE ONLY PLACE YOU NEED TO EDIT TO ADD/CHANGE COLUMNS
# ==============================================================================
# Key        : property name used in the data (must be unique)
# Header     : text shown in the table header
# Value      : scriptblock receiving ($card = Scryfall response, $id = the card's Scryfall ID,
#              $finish = 'normal', 'foil' or 'etched' from the CSV's Foil column,
#              $qty = number of copies from the CSV's Quantity column, 1 if blank,
#              $added = date from the CSV's Added column as yyyy-MM-dd, blank if missing)
# Type       : 'text' (default), 'number' (sorts numerically), 'html' (value is raw HTML)
# Prefix     : optional text shown before the value in the table (e.g. '$')
# Searchable : $true to include the column in the search box
# Total      : $true to show the sum of this (number) column under the search box
# Badge      : $true to show the value as a colour-coded pill (styled for MTG rarities)
# Delta      : $true to show a signed number (e.g. +1.50 / -0.25) as green (up) or red (down)
# DeltaTotal : $true to show a gain/loss pill: the sum of this Delta column x Qty (green = gain, red = loss)
$Columns = @(


    # The link carries the card image URL (data-img) so the report can pop the picture up on hover.
    # Double-faced cards have no top-level image_uris, so the front face is used.
    @{ Key = 'Name';   Header = 'Name';   Searchable = $true; Type = 'html'
       Value = { param($card, $id)
           $img = if ($card.image_uris) { $card.image_uris.normal }
                  elseif ($card.card_faces) { $card.card_faces[0].image_uris.normal }
           "<a href='$($card.scryfall_uri)' data-img='$img'>$($card.name)</a>" } }

    # Price (each) depends on the finish listed in the CSV: normal -> usd, foil -> usd_foil (falls back to usd
    # if the card has no foil price), etched -> usd_etched
    @{ Key = 'Price';  Header = 'Price';  Type = 'number'; Prefix = '$'
       Value = { param($card, $id, $finish, $qty)
           if (-not (Test-HasPrice $card $finish)) { return 'No price' }
           "{0:N2}" -f (Get-UnitPriceAud $card $finish) } }

    # Change in the price of one copy (AUD) since the previous snapshot, e.g. +1.50 or -0.25.
    # Blank on the very first run; 'New' for a card that was not in the previous snapshot.
    # Compared in USD so exchange-rate changes do not show up as price changes.
    @{ Key = 'Change'; Header = 'Change'; Type = 'number'; Delta = $true; DeltaTotal = $true
       Value = { param($card, $id, $finish, $qty)
           if ($PrevPrices.Count -eq 0) { return '' }
           if (-not (Test-HasPrice $card $finish)) { return '' }
           $key = "$(([string]$id).ToLower())|$finish"
           if (-not $PrevPrices.ContainsKey($key)) { return 'New' }
           $d = ((Get-UnitPriceUsd $card $finish) - $PrevPrices[$key]) * $Aud
           if ([math]::Abs($d) -lt 0.005) { return '0.00' }
           "{0:+0.00;-0.00}" -f $d } }

    # Number of copies, from the CSV's Quantity column
    @{ Key = 'Qty';    Header = 'Qty';    Type = 'number'
       Value = { param($card, $id, $finish, $qty) $qty } }

    # Price x Quantity. This is the column that feeds the grand total under the search box.
    @{ Key = 'LineTotal'; Header = 'Total'; Type = 'number'; Prefix = '$'; Total = $true
       Value = { param($card, $id, $finish, $qty)
           if (-not (Test-HasPrice $card $finish)) { return 'No price' }
           "{0:N2}" -f ((Get-UnitPriceAud $card $finish) * $qty) } }

    @{ Key = 'collector_number'; Header = 'Card Number'; Searchable = $true
       Value = { param($card, $id) $card.set+"#"+$card.collector_number } }

    @{ Key = 'Rarity'; Header = 'Rarity'; Badge = $true; Searchable = $true
    Value = { param($card, $id) $card.rarity } }

    # Finish from the CSV's Foil column: Normal, Foil or Etched
    @{ Key = 'Finish'; Header = 'Finish'; Searchable = $true
       Value = { param($card, $id, $finish)
           $f = Get-EffectiveFinish $card $finish
           $f.Substring(0, 1).ToUpper() + $f.Substring(1) } }

    # Special treatments from Scryfall's promo_types, e.g. galaxyfoil, surgefoil, textured, serialized.
    # 'boosterfun' is left out because it only means "found in Play Boosters", not a visible treatment.
    @{ Key = 'Treatment'; Header = 'Treatment'; Searchable = $true
       Value = { param($card, $id)
           $t = @($card.promo_types | Where-Object { $_ -and $_ -ne 'boosterfun' } | ForEach-Object {
               $w = $_ -replace '([a-z])foil', '$1 foil'
               $w.Substring(0, 1).ToUpper() + $w.Substring(1) })
           $t -join ', ' } }

    # Date the card was added, from the CSV's Added column (yyyy-MM-dd)
    @{ Key = 'Added';  Header = 'Date Added';
       Value = { param($card, $id, $finish, $qty, $added) $added } }

    # Clickable eBay icon that opens an eBay search for the card name (front face only for double-faced
    # cards, since eBay copes badly with the "A // B" form). Opens in a new tab.
    @{ Key = 'Ebay'; Header = 'eBay Search'; Type = 'html'
       Value = { param($card, $id)
           $searchName = ($card.name -split ' // ')[0]
           $q = [uri]::EscapeDataString((("$searchName $($card.collector_number) $EbaySearchSuffix") -replace '\s+', ' ').Trim())
           "<a class='ebay-link' href='https://$EbayDomain/sch/i.html?_nkw=$q$EbayLocationParam' target='_blank' rel='noopener' title='Search eBay'><svg xmlns='http://www.w3.org/2000/svg' width='40' height='18' viewBox='0 0 40 18'><text x='0' y='14' font-family='Arial,Helvetica,sans-serif' font-size='16' font-weight='bold' letter-spacing='-0.5'><tspan fill='#e53238'>e</tspan><tspan fill='#0064d2'>b</tspan><tspan fill='#f5af02'>a</tspan><tspan fill='#86b817'>y</tspan></text></svg></a>" } }

    # Same search, limited to SOLD listings (LH_Sold + LH_Complete), so you see what cards actually sold for.
    @{ Key = 'EbaySold'; Header = 'eBay Sold'; Type = 'html'
       Value = { param($card, $id)
           $searchName = ($card.name -split ' // ')[0]
           $q = [uri]::EscapeDataString((("$searchName $($card.collector_number) $EbaySearchSuffix") -replace '\s+', ' ').Trim())
           "<a class='ebay-link' href='https://$EbayDomain/sch/i.html?_nkw=$q&LH_Sold=1&LH_Complete=1$EbayLocationParam' target='_blank' rel='noopener' title='Search eBay sold listings'><svg xmlns='http://www.w3.org/2000/svg' width='36' height='18' viewBox='0 0 36 18'><text x='0' y='14' font-family='Arial,Helvetica,sans-serif' font-size='15' font-weight='bold' fill='#c0392b'>sold</text></svg></a>" } }
)