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

# 2. Read Moving Partis and Static Points.xlsx
$xlsxPath = "Moving Partis and Static Points.xlsx"
$zip = [System.IO.Compression.ZipFile]::OpenRead($xlsxPath)
$sEntry = $zip.GetEntry('xl/sharedStrings.xml')
$strings = @()
if ($sEntry) {
    $sr = New-Object System.IO.StreamReader($sEntry.Open())
    [xml]$xml = $sr.ReadToEnd()
    $sr.Close()
    foreach ($si in $xml.sst.si) {
        if ($si.t) { $strings += $si.t }
        elseif ($si.r) { $strings += (($si.r | ForEach-Object { $_.t }) -join '') }
        else { $strings += '' }
    }
}
$sheetEntry = $zip.GetEntry('xl/worksheets/sheet1.xml')
$sr2 = New-Object System.IO.StreamReader($sheetEntry.Open())
[xml]$sheet = $sr2.ReadToEnd()
$sr2.Close()
$zip.Dispose()

# Parse excel rows
$rowsData = @()
$currentPS = ""
foreach ($row in $sheet.worksheet.sheetData.row) {
    $r = [int]$row.r
    if ($r -lt 3) { continue } # skip title & header
    $cDict = @{}
    foreach ($c in $row.c) {
        $colLetter = [regex]::Match($c.r, '^[A-Z]+').Value
        $val = $c.v
        if ($c.t -eq 's' -and $val -ne $null) { $val = $strings[[int]$val] }
        if ($val) { $val = $val.Trim() }
        $cDict[$colLetter] = $val
    }
    if ($cDict['B']) {
        $currentPS = $cDict['B']
    }
    if ($currentPS) {
        $rowsData += [pscustomobject]@{
            rowNum    = $r
            ps        = $currentPS
            mpName    = $cDict['C']
            vehicle   = $cDict['D']
            mpOfficer = $cDict['E']
            spName    = $cDict['F']
            spOfficer = $cDict['G']
        }
    }
}

Write-Host "Read $($rowsData.Count) rows from Excel."

# Group by PS
$grouped = $rowsData | Group-Object ps

$allDuties = @()
$dutyCounter = 1

foreach ($psName in $psMandalMap.Keys) {
    $mapping = $psMandalMap[$psName]
    $mName = $mapping.mandal
    $baseCoords = $mandalCoords[$mName]
    if (-not $baseCoords) {
        $baseCoords = @{ lat = 18.2969; lng = 83.8967 } # fallback Srikakulam
    }
    $psBaseLat = $baseCoords.lat + $mapping.offsetLat
    $psBaseLng = $baseCoords.lng + $mapping.offsetLng

    # Check if this PS has rows in Excel
    $excelRows = $grouped | Where-Object { $_.Name.Trim().ToLower() -eq $psName.Trim().ToLower() }
    
    $mpList = @()
    $spList = @()

    if ($excelRows -and $excelRows.Group.Count -gt 0) {
        foreach ($r in $excelRows.Group) {
            if ($r.mpName -or $r.mpOfficer) {
                $mpList += @{
                    name = if ($r.mpName) { $r.mpName } else { "Moving Party-$($mpList.Count + 1)" }
                    vehicle = if ($r.vehicle) { $r.vehicle } else { "4W" }
                    officer = if ($r.mpOfficer) { $r.mpOfficer } else { "Duty Incharge" }
                }
            }
            if ($r.spName -or $r.spOfficer) {
                $spList += @{
                    name = if ($r.spName) { $r.spName } else { "Static Point-$($spList.Count + 1)" }
                    vehicle = "Fixed Point"
                    officer = if ($r.spOfficer) { $r.spOfficer } else { "Duty Incharge" }
                }
            }
        }
    }

    # If no specific entries yet, provide standard default points
    if ($mpList.Count -eq 0) {
        $mpList += @{ name = "Moving Party-1"; vehicle = "4W"; officer = "Duty Incharge (To be assigned)" }
        $mpList += @{ name = "Moving Party-2"; vehicle = "2W"; officer = "Duty Incharge (To be assigned)" }
    }
    if ($spList.Count -eq 0) {
        $spList += @{ name = "Static Point-1"; vehicle = "Fixed Point"; officer = "Duty Incharge (To be assigned)" }
        $spList += @{ name = "Static Point-2"; vehicle = "Fixed Point"; officer = "Duty Incharge (To be assigned)" }
    }

    # Add Moving Parties
    $mpIdx = 0
    foreach ($mp in $mpList) {
        $mpIdx++
        $phone = ""
        $cleanPhoneMatch = [regex]::Match($mp.officer, '\b[6-9]\d{9}\b')
        if ($cleanPhoneMatch.Success) {
            $phone = $cleanPhoneMatch.Value
        }
        $slug = ($psName -replace '[^a-zA-Z0-9]', '_') + "_MP_$mpIdx"
        $allDuties += [pscustomobject]@{
            id          = "DUTY_$slug"
            category    = "Moving Parties"
            categoryKey = "MP"
            ps          = $psName
            mandal      = $mName
            unitName    = "$($mp.name) ($psName)"
            route       = $mp.name
            vehicle     = $mp.vehicle
            officer     = $mp.officer
            phone       = $phone
            supporting  = "Vehicle: $($mp.vehicle) • PS: $psName"
            lat         = [math]::Round($psBaseLat + (0.003 * $mpIdx), 6)
            lng         = [math]::Round($psBaseLng + (0.004 * $mpIdx), 6)
        }
    }

    # Add Static Points
    $spIdx = 0
    foreach ($sp in $spList) {
        $spIdx++
        $phone = ""
        $cleanPhoneMatch = [regex]::Match($sp.officer, '\b[6-9]\d{9}\b')
        if ($cleanPhoneMatch.Success) {
            $phone = $cleanPhoneMatch.Value
        }
        $slug = ($psName -replace '[^a-zA-Z0-9]', '_') + "_SP_$spIdx"
        $allDuties += [pscustomobject]@{
            id          = "DUTY_$slug"
            category    = "Static Points"
            categoryKey = "SP"
            ps          = $psName
            mandal      = $mName
            unitName    = "$($sp.name) ($psName)"
            route       = $sp.name
            vehicle     = "Fixed Point"
            officer     = $sp.officer
            phone       = $phone
            supporting  = "Type: Static Point • PS: $psName"
            lat         = [math]::Round($psBaseLat - (0.003 * $spIdx), 6)
            lng         = [math]::Round($psBaseLng - (0.004 * $spIdx), 6)
        }
    }
}

Write-Host "Total duties generated: $($allDuties.Count)"
$mpCount = ($allDuties | Where-Object { $_.categoryKey -eq "MP" }).Count
$spCount = ($allDuties | Where-Object { $_.categoryKey -eq "SP" }).Count
Write-Host " - Moving Parties: $mpCount"
Write-Host " - Static Points: $spCount"

# Output moving_static.json
$jsonStr = $allDuties | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText("$PSScriptRoot\moving_static.json", $jsonStr, [System.Text.Encoding]::UTF8)

# Output moving_static_data.js
$jsContent = "window.MOVING_STATIC_DATA = $jsonStr;`nwindow.ROUTE_MOBILES_DATA = window.MOVING_STATIC_DATA;"
[System.IO.File]::WriteAllText("$PSScriptRoot\moving_static_data.js", $jsContent, [System.Text.Encoding]::UTF8)

Write-Host "Successfully generated moving_static.json and moving_static_data.js"
