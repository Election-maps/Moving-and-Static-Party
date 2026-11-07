Add-Type -AssemblyName System.IO.Compression.FileSystem

# 1. Parse Mandal Centroids from mandal_borders.js
$mandalBordersRaw = Get-Content mandal_borders.js -Raw
$jsonStart = $mandalBordersRaw.IndexOf('{')
$jsonEnd = $mandalBordersRaw.LastIndexOf('}')
$mandalJson = $mandalBordersRaw.Substring($jsonStart, $jsonEnd - $jsonStart + 1) | ConvertFrom-Json

$mandalCoords = @{}
foreach ($f in $mandalJson.features) {
    $mName = $f.properties.name.ToUpper().Trim()
    $coords = $f.geometry.coordinates
    $allPts = @()
    if ($f.geometry.type -eq 'Polygon') {
        foreach ($ring in $coords) { foreach ($pt in $ring) { $allPts += ,@($pt[1], $pt[0]) } }
    } elseif ($f.geometry.type -eq 'MultiPolygon') {
        foreach ($poly in $coords) { foreach ($ring in $poly) { foreach ($pt in $ring) { $allPts += ,@($pt[1], $pt[0]) } } }
    }
    if ($allPts.Count -gt 0) {
        $avgLat = ($allPts | ForEach-Object { $_[0] } | Measure-Object -Average).Average
        $avgLng = ($allPts | ForEach-Object { $_[1] } | Measure-Object -Average).Average
        $mandalCoords[$mName] = @{ lat = [math]::Round($avgLat, 6); lng = [math]::Round($avgLng, 6) }
    }
}

$psMandalMap = @{
    "Srikakulam I Town"   = @{ mandal = "SRIKAKULAM"; offsetLat = 0.005; offsetLng = 0.0 }
    "Srikakulam Rural"    = @{ mandal = "SRIKAKULAM"; offsetLat = -0.005; offsetLng = -0.005 }
    "Gara"                = @{ mandal = "GARA"; offsetLat = 0.0; offsetLng = 0.0 }
    "Srikakulam II Town"  = @{ mandal = "SRIKAKULAM"; offsetLat = -0.002; offsetLng = 0.008 }
    "Amadalavalasa"       = @{ mandal = "AMADALAVALASA"; offsetLat = 0.0; offsetLng = 0.0 }
    "Sarubujjili"         = @{ mandal = "SARUBUJJILI"; offsetLat = 0.0; offsetLng = 0.0 }
    "Burja"               = @{ mandal = "BURJA"; offsetLat = 0.0; offsetLng = 0.0 }
    "Ponduru"             = @{ mandal = "PONDURU"; offsetLat = 0.0; offsetLng = 0.0 }
    "J.R.Puram"           = @{ mandal = "RANASTALAM"; offsetLat = 0.003; offsetLng = 0.0 }
    "Etcherla"            = @{ mandal = "ETCHERLA"; offsetLat = 0.0; offsetLng = 0.0 }
    "Laveru"              = @{ mandal = "LAVERU"; offsetLat = 0.0; offsetLng = 0.0 }
    "G. Sigadam"          = @{ mandal = "GANGUVARISIGADAM"; offsetLat = 0.0; offsetLng = 0.0 }
    "Tekkali PS"          = @{ mandal = "TEKKALI"; offsetLat = 0.0; offsetLng = 0.0 }
    "Santhabommali"       = @{ mandal = "SANTHABOMMALI"; offsetLat = 0.0; offsetLng = 0.0 }
    "Naupada"             = @{ mandal = "SANTHABOMMALI"; offsetLat = -0.01; offsetLng = 0.015 }
    "Kotabommali"         = @{ mandal = "KOTABOMMALI"; offsetLat = 0.0; offsetLng = 0.0 }
    "Nandigam"            = @{ mandal = "NANDIGAM"; offsetLat = 0.0; offsetLng = 0.0 }
    "Narasannapeta"       = @{ mandal = "NARASANNAPETA"; offsetLat = 0.0; offsetLng = 0.0 }
    "Jalumuru"            = @{ mandal = "JALUMURU"; offsetLat = 0.0; offsetLng = 0.0 }
    "Polaki"              = @{ mandal = "POLAKI"; offsetLat = 0.0; offsetLng = 0.0 }
    "Saravakota"          = @{ mandal = "SARAVAKOTA"; offsetLat = 0.0; offsetLng = 0.0 }
    "Pathapatnam"         = @{ mandal = "PATHAPATNAM"; offsetLat = 0.0; offsetLng = 0.0 }
    "Meliaputti"          = @{ mandal = "MELIAPUTTI"; offsetLat = 0.0; offsetLng = 0.0 }
    "Kothuru"             = @{ mandal = "KOTHURU"; offsetLat = 0.0; offsetLng = 0.0 }
    "Hiramandalam"        = @{ mandal = "HIRAMANDALAM"; offsetLat = 0.0; offsetLng = 0.0 }
    "Kasibugga"           = @{ mandal = "PALASA"; offsetLat = 0.003; offsetLng = 0.0 }
    "Vajrapukothuru"      = @{ mandal = "VAJRAPUKOTHURU"; offsetLat = 0.0; offsetLng = 0.0 }
    "Mandasa"             = @{ mandal = "MANDASA"; offsetLat = 0.0; offsetLng = 0.0 }
    "Sompeta"             = @{ mandal = "SOMPETA"; offsetLat = 0.0; offsetLng = 0.0 }
    "Kanchili"            = @{ mandal = "KANCHILI"; offsetLat = 0.0; offsetLng = 0.0 }
    "Baruva"              = @{ mandal = "SOMPETA"; offsetLat = -0.015; offsetLng = 0.02 }
    "Itchapuram Town"     = @{ mandal = "ICHCHAPURAM"; offsetLat = 0.006; offsetLng = 0.0 }
    "Itchapuram Rural"    = @{ mandal = "ICHCHAPURAM"; offsetLat = -0.006; offsetLng = -0.005 }
    "Kaviti"              = @{ mandal = "KAVITI"; offsetLat = 0.0; offsetLng = 0.0 }
}

# 2. Read Moving Partis and Static Points.xlsx with InnerText
$xlsxPath = "Moving Partis and Static Points.xlsx"
$zip = [System.IO.Compression.ZipFile]::OpenRead($xlsxPath)
$sEntry = $zip.GetEntry('xl/sharedStrings.xml')
$strings = @()
if ($sEntry) {
    $sr = New-Object System.IO.StreamReader($sEntry.Open(), [System.Text.Encoding]::UTF8)
    [xml]$xml = $sr.ReadToEnd()
    $sr.Close()
    foreach ($si in $xml.sst.si) {
        $strings += $si.InnerText
    }
}
$sheetEntry = $zip.GetEntry('xl/worksheets/sheet1.xml')
$sr2 = New-Object System.IO.StreamReader($sheetEntry.Open(), [System.Text.Encoding]::UTF8)
[xml]$sheet = $sr2.ReadToEnd()
$sr2.Close()
$zip.Dispose()

$rawRows = @()
$currentPS = ""
foreach ($row in $sheet.worksheet.sheetData.row) {
    $r = [int]$row.r
    if ($r -lt 3) { continue }
    $cDict = @{}
    foreach ($c in $row.c) {
        $colLetter = [regex]::Match($c.r, '^[A-Z]+').Value
        $val = $c.v
        if ($c.t -eq 's' -and $val -ne $null) { $val = $strings[[int]$val] }
        if ($val) { $val = $val.Trim() }
        $cDict[$colLetter] = $val
    }
    
    if ($cDict['B']) { $currentPS = $cDict['B'] }

    $rawRows += [pscustomobject]@{
        Row = $r
        PS = $currentPS
        MP_Name = $cDict['C']
        Vehicle = $cDict['D']
        MP_Officer = $cDict['E']
        SP_Name = $cDict['F']
        Village = $cDict['G']
        SP_Officer = $cDict['H']
    }
}

Write-Host "Read $($rawRows.Count) rows from Excel."

# Group rows by PS
$grouped = $rawRows | Group-Object PS

$allDuties = @()

foreach ($grp in $grouped) {
    $psName = $grp.Name
    if (-not $psName) { continue }

    # Find mandal mapping
    $mapping = $psMandalMap[$psName]
    if (-not $mapping) {
        foreach ($k in $psMandalMap.Keys) {
            if ($k.Trim().ToLower() -eq $psName.Trim().ToLower()) {
                $mapping = $psMandalMap[$k]
                break
            }
        }
    }
    if (-not $mapping) {
        $mapping = @{ mandal = "SRIKAKULAM"; offsetLat = 0.0; offsetLng = 0.0 }
    }

    $mName = $mapping.mandal
    $baseCoords = $mandalCoords[$mName]
    if (-not $baseCoords) {
        $baseCoords = @{ lat = 18.2969; lng = 83.8967 }
    }
    $psBaseLat = $baseCoords.lat + $mapping.offsetLat
    $psBaseLng = $baseCoords.lng + $mapping.offsetLng

    $mpCounter = 0
    $spCounter = 0

    foreach ($r in $grp.Group) {
        # Check Moving Party
        $mpName = $r.MP_Name
        $veh = $r.Vehicle
        $mpOff = $r.MP_Officer

        $isMpValid = ($mpName -or $mpOff) -and ($mpName -notmatch '^(nil|nill)$') -and ($mpOff -notmatch '^(nil|nill)$')
        if ($isMpValid) {
            $mpCounter++
            $displayName = if ($mpName) { $mpName } else { "Moving Party-$mpCounter" }
            $cleanPhone = ""
            $phoneMatch = [regex]::Match($mpOff, '\b[6-9]\d{9}\b')
            if ($phoneMatch.Success) { $cleanPhone = $phoneMatch.Value }

            $vehicleStr = if ($veh) { $veh } else { "4W" }
            $officerStr = if ($mpOff) { $mpOff } else { "Duty Incharge" }
            $supportStr = "Vehicle: " + $vehicleStr + " | PS: " + $psName

            $slug = ($psName -replace '[^a-zA-Z0-9]', '_') + "_MP_$mpCounter"
            $latOffset = (0.0035 * ($mpCounter % 4)) - 0.005
            $lngOffset = (0.0045 * [math]::Floor($mpCounter / 4)) + 0.004

            $allDuties += [pscustomobject]@{
                id          = "DUTY_$slug"
                category    = "Moving Parties"
                categoryKey = "MP"
                ps          = $psName
                mandal      = $mName
                unitName    = "$displayName ($psName)"
                route       = $displayName
                vehicle     = $vehicleStr
                village     = ""
                officer     = $officerStr
                phone       = $cleanPhone
                supporting  = $supportStr
                lat         = [math]::Round($psBaseLat + $latOffset, 6)
                lng         = [math]::Round($psBaseLng + $lngOffset, 6)
            }
        }

        # Check Static Point
        $spName = $r.SP_Name
        $vill = $r.Village
        $spOff = $r.SP_Officer

        $isSpValid = ($spName -or $vill -or $spOff) -and 
                     ($spName -notmatch '^(nil|nill|0)$') -and 
                     ($vill -notmatch '^(nil|nill|0)$') -and 
                     ($spOff -notmatch '^(nil|nill|0)$')

        if ($isSpValid) {
            $spCounter++
            $displayName = if ($spName -and $spName -notmatch '^\d+$') { 
                $spName 
            } elseif ($spName) { 
                "Static Point-$spName" 
            } else { 
                "Static Point-$spCounter" 
            }

            if ($vill) {
                $displayName += ": $vill"
            }

            $cleanPhone = ""
            $phoneMatch = [regex]::Match($spOff, '\b[6-9]\d{9}\b')
            if ($phoneMatch.Success) { $cleanPhone = $phoneMatch.Value }

            $villageStr = if ($vill) { $vill } else { "Fixed Point" }
            $officerStr = if ($spOff) { $spOff } else { "Duty Incharge" }
            $villField = if ($vill) { $vill } else { "" }
            $supportStr = "Village: " + $villageStr + " | PS: " + $psName

            $slug = ($psName -replace '[^a-zA-Z0-9]', '_') + "_SP_$spCounter"
            $latOffset = -(0.0035 * ($spCounter % 4)) + 0.003
            $lngOffset = -(0.0045 * [math]::Floor($spCounter / 4)) - 0.004

            $allDuties += [pscustomobject]@{
                id          = "DUTY_$slug"
                category    = "Static Points"
                categoryKey = "SP"
                ps          = $psName
                mandal      = $mName
                unitName    = "$displayName ($psName)"
                route       = $displayName
                vehicle     = "Fixed Point"
                village     = $villField
                officer     = $officerStr
                phone       = $cleanPhone
                supporting  = $supportStr
                lat         = [math]::Round($psBaseLat + $latOffset, 6)
                lng         = [math]::Round($psBaseLng + $lngOffset, 6)
            }
        }
    }

    # If this PS had no MP and no SP entries in Excel (e.g. Srikakulam Rural, G. Sigadam, Kothuru, etc.)
    # Provide 1 pending entry so that the PS appears in the dropdown & directory
    if ($mpCounter -eq 0 -and $spCounter -eq 0) {
        $slugMP = ($psName -replace '[^a-zA-Z0-9]', '_') + "_MP_1"
        $allDuties += [pscustomobject]@{
            id          = "DUTY_$slugMP"
            category    = "Moving Parties"
            categoryKey = "MP"
            ps          = $psName
            mandal      = $mName
            unitName    = "Moving Party-1 ($psName)"
            route       = "Moving Party-1"
            vehicle     = "4W"
            village     = ""
            officer     = "Duty Incharge (Pending Assignment)"
            phone       = ""
            supporting  = "Vehicle: 4W | PS: $psName"
            lat         = [math]::Round($psBaseLat + 0.003, 6)
            lng         = [math]::Round($psBaseLng + 0.003, 6)
        }

        $slugSP = ($psName -replace '[^a-zA-Z0-9]', '_') + "_SP_1"
        $allDuties += [pscustomobject]@{
            id          = "DUTY_$slugSP"
            category    = "Static Points"
            categoryKey = "SP"
            ps          = $psName
            mandal      = $mName
            unitName    = "Static Point-1 ($psName)"
            route       = "Static Point-1"
            vehicle     = "Fixed Point"
            village     = ""
            officer     = "Duty Incharge (Pending Assignment)"
            phone       = ""
            supporting  = "Type: Static Point | PS: $psName"
            lat         = [math]::Round($psBaseLat - 0.003, 6)
            lng         = [math]::Round($psBaseLng - 0.003, 6)
        }
    }
}

Write-Host "--------------------------------------------------------"
Write-Host "Total duties generated from Excel: $($allDuties.Count)"
$mpCount = ($allDuties | Where-Object { $_.categoryKey -eq "MP" }).Count
$spCount = ($allDuties | Where-Object { $_.categoryKey -eq "SP" }).Count
Write-Host " - Moving Parties: $mpCount"
Write-Host " - Static Points: $spCount"

# Output moving_static.json (UTF-8)
$jsonStr = $allDuties | ConvertTo-Json -Depth 5
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText("$PSScriptRoot\moving_static.json", $jsonStr, $utf8NoBom)

# Output moving_static_data.js (UTF-8)
$jsContent = "window.MOVING_STATIC_DATA = $jsonStr;`nwindow.ROUTE_MOBILES_DATA = window.MOVING_STATIC_DATA;"
[System.IO.File]::WriteAllText("$PSScriptRoot\moving_static_data.js", $jsContent, $utf8NoBom)

Write-Host "Successfully generated moving_static.json and moving_static_data.js matching Excel exactly!"
