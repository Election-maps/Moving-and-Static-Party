Add-Type -AssemblyName System.IO.Compression.FileSystem

# 1. Parse Mandal Centroids from mandal_borders.js
$mandalBordersRaw = Get-Content "$PSScriptRoot\mandal_borders.js" -Raw
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

# 2. Open Excel file
$xlsxPath = "$PSScriptRoot\Moving Partis and Static Points.xlsx"
$tempPath = [System.IO.Path]::GetTempFileName() + ".xlsx"

$fsIn = New-Object System.IO.FileStream($xlsxPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
$fsOut = New-Object System.IO.FileStream($tempPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
$fsIn.CopyTo($fsOut)
$fsIn.Close()
$fsOut.Close()

$zip = [System.IO.Compression.ZipFile]::OpenRead($tempPath)
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
$srSheet = New-Object System.IO.StreamReader($sheetEntry.Open(), [System.Text.Encoding]::UTF8)
[xml]$sheetXml = $srSheet.ReadToEnd()
$srSheet.Close()
$zip.Dispose()
Remove-Item -Force $tempPath -ErrorAction SilentlyContinue

# Parse Rows
$rawRows = @()
$currentPS = ""
foreach ($row in $sheetXml.worksheet.sheetData.row) {
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
    
    if ($cDict['B']) {
        $currentPS = $cDict['B'].Trim()
        if ($currentPS -replace '\s+', ' ' -eq 'G Sigadam') {
            $currentPS = 'G. Sigadam'
        }
    }

    $rawRows += [pscustomobject]@{
        Row = $r
        PS = $currentPS
        MP_Num = $cDict['C']
        Veh = $cDict['D']
        MP_Off = $cDict['E']
        SP_Num = $cDict['F']
        Vill = $cDict['G']
        SP_Off = $cDict['H']
    }
}

Write-Host "Total rows read from Excel: $($rawRows.Count)"

# Helper function to extract clean 10-digit mobile
function Extract-Phone($text) {
    if (-not $text) { return "" }
    $m = [regex]::Match($text, '\b[6-9]\d{9}\b')
    if ($m.Success) { return $m.Value }
    $m2 = [regex]::Match($text, '[6-9]\s*\d\s*\d\s*\d\s*\d\s*\d\s*\d\s*\d\s*\d\s*\d')
    if ($m2.Success) { return ($m2.Value -replace '\s+', '') }
    return ""
}

# Clean Party Names: convert typos like 'Moving Partyy-1', 'Moving Paerty -1', 'Moveing Party-1' to 'Moving Party-X'
function Normalize-PartyName($text, $fallbackIndex) {
    if (-not $text) { return "Moving Party-$fallbackIndex" }
    $t = $text.Trim()
    if ($t -match '(?i)(?:mov(?:e)?ing\s*pa(?:e)?rt(?:y|yy|ies)?)\s*[-:]?\s*(\d+)') {
        return "Moving Party-" + $matches[1]
    }
    if ($t -match '^\d+$') {
        return "Moving Party-$t"
    }
    return $t
}

function Sanitize-Text($str) {
    if (-not $str) { return "" }
    $s = [regex]::Replace($str, "[\u2018\u2019\u201A\u201B`]", "'")
    $s = [regex]::Replace($s, "[\u201C\u201D]", '"')
    $s = [regex]::Replace($s, "[\u2013\u2014]", "-")
    $s = [regex]::Replace($s, "[^\x20-\x7E]", " ")
    $s = [regex]::Replace($s, "PC\s*'\s*s\s*", "PC ")
    return ($s -replace '\s+', ' ').Trim()
}

# Group by PS preserving order of appearance in Excel
$psList = @()
foreach ($r in $rawRows) {
    if ($r.PS -and $psList -notcontains $r.PS) {
        $psList += $r.PS
    }
}

$allDuties = @()

foreach ($psName in $psList) {
    $psRows = $rawRows | Where-Object { $_.PS -eq $psName }

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

    $psMovingDuties = @()
    $psStaticDuties = @()

    $mpCount = 0
    $spCount = 0

    foreach ($r in $psRows) {
        # 1. Process Moving Party
        $mpNum = $r.MP_Num
        $veh = $r.Veh
        $mpOff = $r.MP_Off

        $isMpValid = ($mpNum -or $mpOff) -and 
                     ($mpNum -notmatch '^(?i)(nil|nill|0)$') -and 
                     ($mpOff -notmatch '^(?i)(nil|nill|0)$')

        if ($isMpValid) {
            $mpCount++
            $cleanVeh = if ($veh) { $veh.Trim() -replace '(?i)\s*wheeler', '' } else { "4W" }
            if ($cleanVeh -match '(?i)4w') { $cleanVeh = "4W" }
            elseif ($cleanVeh -match '(?i)2w') { $cleanVeh = "2W" }

            $cleanPhone = Extract-Phone $mpOff
            $cleanOfficer = if ($mpOff) { Sanitize-Text ($mpOff -replace '\r?\n', ', ') } else { "Duty Incharge ($psName)" }

            $dispName = Normalize-PartyName $mpNum $mpCount

            # Check if this row also contains village route for Moving Party
            $mpVillage = ""
            if ($r.Vill -and $r.SP_Num -match '^(?i)(nil|nill|0)$') {
                $mpVillage = Sanitize-Text $r.Vill
            }

            $supportStr = "Vehicle: $cleanVeh | PS: $psName"
            if ($mpVillage) {
                $supportStr += " | Route: $mpVillage"
            }

            $slug = ($psName -replace '[^a-zA-Z0-9]', '_') + "_MP_$mpCount"
            $latOffset = (0.0035 * ($mpCount % 4)) - 0.005
            $lngOffset = (0.0045 * [math]::Floor($mpCount / 4)) + 0.004

            $psMovingDuties += [pscustomobject]@{
                id          = "DUTY_$slug"
                category    = "Moving Parties"
                categoryKey = "MP"
                ps          = $psName
                mandal      = $mName
                unitName    = "$dispName ($psName)"
                route       = $dispName
                vehicle     = $cleanVeh
                village     = $mpVillage
                officer     = $cleanOfficer
                phone       = $cleanPhone
                supporting  = $supportStr
                lat         = [math]::Round($psBaseLat + $latOffset, 6)
                lng         = [math]::Round($psBaseLng + $lngOffset, 6)
            }
        }

        # 2. Process Static Point
        $spNum = $r.SP_Num
        $vill = $r.Vill
        $spOff = $r.SP_Off

        $isSpValid = ($vill -or $spOff) -and 
                     ($vill -notmatch '^(?i)(nil|nill|0)$') -and 
                     ($spOff -notmatch '^(?i)(nil|nill|0)$')

        if ($spNum -eq $psName -and -not $spOff -and -not $vill) {
            $isSpValid = $false
        }

        if ($isSpValid) {
            $spCount++
            $cleanPhone = Extract-Phone $spOff
            $cleanOfficer = if ($spOff) { Sanitize-Text ($spOff -replace '\r?\n', ', ') } else { "Duty Incharge" }
            $cleanVill = if ($vill) { Sanitize-Text $vill } else { "" }

            # Extract village from spOff if village column was empty (e.g. Tekkali)
            if (-not $cleanVill -and $cleanOfficer -match '^(?:\d+\.\s*)?([a-zA-Z\s]+?)[,\.]?\s+(?:Hc|Pc|Asi|Si)\b') {
                $cleanVill = Sanitize-Text $matches[1]
            }

            $dispName = "Static Point-$spCount"
            if ($cleanVill) {
                $dispName += ": $cleanVill"
            }

            $supportStr = if ($cleanVill) { "Village: $cleanVill | PS: $psName" } else { "Static Point | PS: $psName" }

            $slug = ($psName -replace '[^a-zA-Z0-9]', '_') + "_SP_$spCount"
            $latOffset = -(0.0035 * ($spCount % 4)) + 0.003
            $lngOffset = -(0.0045 * [math]::Floor($spCount / 4)) - 0.004

            $psStaticDuties += [pscustomobject]@{
                id          = "DUTY_$slug"
                category    = "Static Points"
                categoryKey = "SP"
                ps          = $psName
                mandal      = $mName
                unitName    = "$dispName ($psName)"
                route       = $dispName
                vehicle     = "Fixed Point"
                village     = $cleanVill
                officer     = $cleanOfficer
                phone       = $cleanPhone
                supporting  = $supportStr
                lat         = [math]::Round($psBaseLat + $latOffset, 6)
                lng         = [math]::Round($psBaseLng + $lngOffset, 6)
            }
        }
    }

    # Fallback if station had no entries at all (e.g. Kothuru)
    if ($mpCount -eq 0 -and $spCount -eq 0) {
        $slugMP = ($psName -replace '[^a-zA-Z0-9]', '_') + "_MP_1"
        $psMovingDuties += [pscustomobject]@{
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
        $psStaticDuties += [pscustomobject]@{
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

    $allDuties += $psMovingDuties
    $allDuties += $psStaticDuties
}

Write-Host "=========================================================="
Write-Host "Total duties successfully generated: $($allDuties.Count)"
$mpTotal = ($allDuties | Where-Object { $_.categoryKey -eq "MP" }).Count
$spTotal = ($allDuties | Where-Object { $_.categoryKey -eq "SP" }).Count
Write-Host " - Moving Parties: $mpTotal"
Write-Host " - Static Points: $spTotal"
Write-Host " - Total Police Stations: $($psList.Count)"
Write-Host "=========================================================="

# Output moving_static.json (UTF-8 without BOM)
$jsonStr = $allDuties | ConvertTo-Json -Depth 5
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText("$PSScriptRoot\moving_static.json", $jsonStr, $utf8NoBom)

# Output moving_static_data.js (UTF-8 without BOM)
$jsContent = "window.MOVING_STATIC_DATA = $jsonStr;`nwindow.ROUTE_MOBILES_DATA = window.MOVING_STATIC_DATA;"
[System.IO.File]::WriteAllText("$PSScriptRoot\moving_static_data.js", $jsContent, $utf8NoBom)

# Update officer.html EMBEDDED_DUTIES_DATA
$compactJson = ($allDuties | ConvertTo-Json -Compress)
$officerPath = "$PSScriptRoot\officer.html"
$officerContent = [System.IO.File]::ReadAllText($officerPath, [System.Text.Encoding]::UTF8)
$officerContent = [regex]::Replace($officerContent, 'const EMBEDDED_DUTIES_DATA = \[.*?\];', "const EMBEDDED_DUTIES_DATA = $compactJson;", [System.Text.RegularExpressions.RegexOptions]::Singleline)
[System.IO.File]::WriteAllText($officerPath, $officerContent, $utf8NoBom)

# Generate Srikakulam_Police_Duties_Online_Offline_Report.csv
$csvLines = @()
$csvLines += 'Sl No,Duty Category,Police Station,Party or Point Name,Duty In-charge Officer,Mobile Number,Vehicle or Type,Headquarters or Village,Live Status,Latitude,Longitude'
$sl = 1
foreach ($d in $allDuties) {
    $cat = $d.category
    $ps = if ($d.ps -match '(?i)\s+PS$') { $d.ps } else { $d.ps + " PS" }
    $unit = $d.unitName
    $off = ($d.officer -replace '"', '""').Trim()
    $ph = $d.phone
    $veh = $d.vehicle
    $vill = ($d.village -replace '"', '""').Trim()
    $lat = $d.lat
    $lng = $d.lng
    $csvLines += "`"$sl`",`"$cat`",`"$ps`",`"$unit`",`"$off`",`"$ph`",`"$veh`",`"$vill`",`"OFFLINE`",`"$lat`",`"$lng`""
    $sl++
}
$csvText = [string]::Join("`r`n", $csvLines)
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText("$PSScriptRoot\Srikakulam_Police_Duties_Online_Offline_Report.csv", $csvText, $utf8Bom)

Write-Host "Successfully regenerated moving_static.json, moving_static_data.js, officer.html, and CSV Report!"
