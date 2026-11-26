<#
.SYNOPSIS
  Cordaves Avicary — Read-only Windows triage probe for remote support sessions.
.DESCRIPTION
  Collects system specs, RAM/pagefile pressure, internal & external USB drives,
  disk health status, disk/USB/power system errors (14 days), app crashes/hangs
  (14 days), Reliability Monitor records, top RAM processes, and basic AV/Adobe
  environment signals — without modifying system settings or scanning user files.

  Saves a UTF-8 (BOM) report to the Desktop and copies it to the clipboard.
#>

[CmdletBinding()]
param(
  [int]$Days = 14,
  [string]$OutputPath = "",
  [switch]$NoClipboard
)

$ErrorActionPreference = 'SilentlyContinue'
$now = Get-Date
$since = $now.AddDays(-$Days)
$flags = New-Object System.Collections.Generic.List[string]
$sections = New-Object System.Collections.Generic.List[string]

function Add-Section {
  param([string]$Title, [string]$Body)
  $sections.Add("================================================================")
  $sections.Add($Title)
  $sections.Add("================================================================")
  $sections.Add(($Body.TrimEnd()))
  $sections.Add("")
}

function Clean-Msg {
  param($Text, [int]$MaxLen = 160)
  $s = ("$Text" -replace '\s+', ' ').Trim()
  if ($s.Length -gt $MaxLen) { return $s.Substring(0, $MaxLen) + "..." }
  return $s
}

Write-Host "Avicary: Starte Nur-Lese-Diagnose (letzte $Days Tage)..." -ForegroundColor Cyan

# 1. SYSTEM, UPTIME & RAM
$os = Get-CimInstance Win32_OperatingSystem
$cs = Get-CimInstance Win32_ComputerSystem
$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
$pf = Get-CimInstance Win32_PageFileUsage | Select-Object -First 1

$totalRamGB = if ($os.TotalVisibleMemorySize) { [math]::Round($os.TotalVisibleMemorySize / 1MB, 1) } else { 0 }
$freeRamGB  = if ($os.FreePhysicalMemory)     { [math]::Round($os.FreePhysicalMemory / 1MB, 1) }     else { 0 }
$freeRamPct = if ($totalRamGB -gt 0)          { [math]::Round(($freeRamGB / $totalRamGB) * 100, 0) } else { 0 }
$uptimeDays = if ($os.LastBootUpTime)         { [math]::Round(($now - $os.LastBootUpTime).TotalDays, 1) } else { -1 }

if ($freeRamPct -lt 15 -and $totalRamGB -gt 0) {
  $flags.Add("[WARNUNG] Sehr wenig freier Arbeitsspeicher: $freeRamGB / $totalRamGB GB ($freeRamPct% frei)")
}
if ($uptimeDays -gt 7) {
  $flags.Add("[HINWEIS] Letzter echter Neustart vor $uptimeDays Tagen (Windows-Schnellstart aktiv?)")
}

$sysLines = @(
  "Erstellt:        $($now.ToString('yyyy-MM-dd HH:mm:ss')) (Zeitzone: $([TimeZoneInfo]::Local.DisplayName))"
  "Geraet:          $($cs.Manufacturer) $($cs.Model)"
  "OS:              $($os.Caption) | Version $($os.Version) (Build $($os.BuildNumber))"
  "Architektur:     $($os.OSArchitecture)"
  "CPU:             $($cpu.Name) ($($cpu.NumberOfCores) Kerne / $($cpu.NumberOfLogicalProcessors) Threads, Last: $($cpu.LoadPercentage)%)"
  "RAM:             $freeRamGB GB frei von $totalRamGB GB ($freeRamPct% frei)"
  "Auslagerung:     $(if ($pf) { "$($pf.CurrentUsage) MB belegt / $($pf.AllocatedBaseSize) MB zugewiesen (Peak: $($pf.PeakUsage) MB)" } else { 'Nicht abrufbar' })"
  "Letzter Boot:    $($os.LastBootUpTime) (Laufzeit: $uptimeDays Tage)"
)
Add-Section "1. SYSTEM & ARBEITSSPEICHER" ($sysLines -join "`r`n")

# 2. LOGICAL DRIVES & FREE SPACE
$driveTypeMap = @{
  0 = 'Unbekannt'
  1 = 'Kein Root'
  2 = 'Wechsel/USB'
  3 = 'Festplatte'
  4 = 'Netzlaufwerk'
  5 = 'CD/DVD'
  6 = 'RAM-Disk'
}

$logical = Get-CimInstance Win32_LogicalDisk | ForEach-Object {
  $sizeGB = if ($_.Size) { [math]::Round($_.Size / 1GB, 1) } else { 0 }
  $freeGB = if ($_.FreeSpace) { [math]::Round($_.FreeSpace / 1GB, 1) } else { 0 }
  $freePct = if ($sizeGB -gt 0) { [math]::Round(($freeGB / $sizeGB) * 100, 0) } else { 0 }

  if ($_.DeviceID -eq 'C:' -and $sizeGB -gt 0 -and ($freeGB -lt 15 -or $freePct -lt 10)) {
    $flags.Add("[WARNUNG] Systemlaufwerk C: hat wenig freien Platz: $freeGB GB ($freePct% frei)")
  }
  elseif ($_.DriveType -in 2, 3 -and $sizeGB -gt 0 -and $freePct -lt 8) {
    $flags.Add("[WARNUNG] Laufwerk $($_.DeviceID) fast voll: $freeGB GB von $sizeGB GB frei ($freePct%)")
  }

  [PSCustomObject]@{
    Laufwerk   = $_.DeviceID
    Name       = $_.VolumeName
    Typ        = $driveTypeMap[[int]$_.DriveType]
    Dateisyst  = $_.FileSystem
    GroesseGB  = $sizeGB
    FreiGB     = $freeGB
    FreiPct    = "$freePct%"
  }
}
Add-Section "2. LAUFWERKE & SPEICHERPLATZ" ($logical | Format-Table -AutoSize | Out-String -Width 240)

# 3. PHYSICAL DISKS & HEALTH
$physDisks = Get-CimInstance Win32_DiskDrive | ForEach-Object {
  if ($_.Status -and $_.Status -ne 'OK') {
    $flags.Add("[KRITISCH] Festplatte '$($_.Model)' meldet Status '$($_.Status)'!")
  }
  [PSCustomObject]@{
    Index     = $_.Index
    Modell    = $_.Model
    Anschluss = $_.InterfaceType
    Typ       = $_.MediaType
    Status    = $_.Status
    GroesseGB = if ($_.Size) { [math]::Round($_.Size / 1GB, 1) } else { 0 }
  }
}
$diskText = ($physDisks | Format-Table -AutoSize | Out-String -Width 240)

try {
  $msftDisks = Get-PhysicalDisk -ErrorAction Stop | Select-Object FriendlyName, BusType, MediaType, HealthStatus, OperationalStatus
  if ($msftDisks) {
    foreach ($d in $msftDisks) {
      if ($d.HealthStatus -and $d.HealthStatus -ne 'Healthy') {
        $flags.Add("[KRITISCH] Storage-Health '$($d.FriendlyName)': $($d.HealthStatus) ($($d.OperationalStatus))")
      }
    }
    $diskText += "`r`n--- Windows Storage Health ---`r`n" + ($msftDisks | Format-Table -AutoSize | Out-String -Width 240)
  }
} catch {
  $diskText += "`r`n(Erweiterter Storage-Health-Status benoetigt ggf. Admin-Rechte - uebersprungen.)"
}
Add-Section "3. PHYSISCHE FESTPLATTEN & USB-LAUFWERKE" $diskText

# 4. SYSTEM EVENT LOG: DISK / USB / NTFS / POWER / BSOD
$sysErrText = ""
try {
  $sysEvents = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; Level = 1, 2, 3; StartTime = $since } -ErrorAction Stop |
    Where-Object { $_.ProviderName -match 'disk|ntfs|storahci|storport|usb|volmgr|partmgr|Kernel-Power|nvme|volsnap|BugCheck|WHEA' })

  if ($sysEvents.Count -gt 0) {
    $diskErrCount = @($sysEvents | Where-Object { $_.ProviderName -match 'disk|ntfs|storahci|storport|usb|volmgr|nvme' }).Count
    if ($diskErrCount -gt 0) {
      $flags.Add("[KRITISCH] $diskErrCount Festplatten-/USB-/Dateisystem-Fehler in den letzten $Days Tagen gefunden!")
    }
    $powerCount = @($sysEvents | Where-Object { $_.ProviderName -match 'Kernel-Power' -and $_.Id -eq 41 }).Count
    if ($powerCount -gt 0) {
      $flags.Add("[WARNUNG] $powerCount unerwartete Abstuerze/Stromausfaelle (Kernel-Power 41) in den letzten $Days Tagen")
    }

    $summary = $sysEvents | Group-Object ProviderName, Id, LevelDisplayName | Sort-Object Count -Descending | ForEach-Object {
      $parts = $_.Name -split ', '
      [PSCustomObject]@{
        Anzahl   = $_.Count
        Quelle   = $parts[0]
        EventID  = $parts[1]
        Stufe    = $parts[2]
      }
    }
    $sysErrText = "Gesamt relevante Eintraege: $($sysEvents.Count)`r`n`r`n--- Haeufigkeit ---`r`n"
    $sysErrText += ($summary | Select-Object -First 15 | Format-Table -AutoSize | Out-String -Width 240)
    $sysErrText += "`r`n--- Letzte 25 Eintraege ---`r`n"
    $sysErrText += ($sysEvents | Select-Object -First 25 TimeCreated, LevelDisplayName, ProviderName, Id, @{ N = 'Meldung'; E = { Clean-Msg $_.Message 145 } } | Format-Table -AutoSize | Out-String -Width 240)
  } else {
    $sysErrText = "Keine Disk-/USB-/NTFS-/Power-Fehler in den letzten $Days Tagen."
  }
} catch {
  $sysErrText = "Keine Eintraege gefunden oder System-Log nicht lesbar ($($_.Exception.Message))."
}
Add-Section "4. SYSTEMPROTOKOLL: FESTPLATTE / USB / ABSTUERZE ($Days TAGE)" $sysErrText

# 5. APPLICATION CRASHES, HANGS & ADOBE SIGNALS
$appErrText = ""
try {
  $appEvents = @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; Level = 1, 2; StartTime = $since } -ErrorAction Stop |
    Where-Object {
      $_.ProviderName -match 'Application Error|Application Hang|Windows Error Reporting|Adobe' -or
      "$($_.Message)" -match 'InDesign|Photoshop|Illustrator|Acrobat|Adobe|CS6'
    })

  if ($appEvents.Count -gt 0) {
    $flags.Add("[HINWEIS] $($appEvents.Count) Programm-Abstuerze oder Haenger im Anwendungsprotokoll ($Days Tage)")
    $appErrText = "Gesamt Eintraege: $($appEvents.Count)`r`n"
    $appErrText += ($appEvents | Select-Object -First 25 TimeCreated, ProviderName, Id, @{ N = 'Meldung'; E = { Clean-Msg $_.Message 165 } } | Format-Table -AutoSize | Out-String -Width 240)
  } else {
    $appErrText = "Keine Anwendungs-Abstuerze oder Haenger in den letzten $Days Tagen."
  }
} catch {
  $appErrText = "Keine Eintraege gefunden oder Anwendungs-Log nicht lesbar ($($_.Exception.Message))."
}
Add-Section "5. PROGRAMM-ABSTUERZE & HAENGER ($Days TAGE)" $appErrText

# 6. RELIABILITY MONITOR (ZUVERLAESSIGKEITSVERLAUF)
$relText = ""
try {
  $rel = @(Get-CimInstance -ClassName Win32_ReliabilityRecords -ErrorAction Stop |
    Where-Object { $_.TimeGenerated -ge $since } |
    Sort-Object TimeGenerated -Descending |
    Select-Object -First 15 TimeGenerated, SourceName, ProductName, @{ N = 'Meldung'; E = { Clean-Msg $_.Message 140 } })
  if ($rel.Count -gt 0) {
    $relText = ($rel | Format-Table -AutoSize | Out-String -Width 240)
  } else {
    $relText = "Keine aktuellen Eintraege im Zuverlaessigkeitsverlauf ($Days Tage)."
  }
} catch {
  $relText = "Zuverlaessigkeitsverlauf nicht abrufbar ($($_.Exception.Message))."
}
Add-Section "6. ZUVERLAESSIGKEITSVERLAUF (LETZTE 15)" $relText

# 7. ADOBE / ANTIVIRUS ENVIRONMENT (NON-RECURSIVE)
$envLines = New-Object System.Collections.Generic.List[string]
foreach ($basePath in @("C:\Program Files\Adobe", "C:\Program Files (x86)\Adobe")) {
  if (Test-Path $basePath) {
    $dirs = (Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name) -join ", "
    $envLines.Add("Adobe-Ordner ($basePath): $(if ($dirs) { $dirs } else { '(leer)' })")
  }
}
$adobeProcs = @(Get-Process | Where-Object { $_.ProcessName -match 'InDesign|Photoshop|Illustrator|Acrobat|Adobe|Creative|CS6|AAM|AGM|AGS' })
if ($adobeProcs.Count -gt 0) {
  $envLines.Add("Laufende Adobe-Prozesse:")
  foreach ($p in $adobeProcs) {
    $envLines.Add("  - $($p.ProcessName) (PID $($p.Id), RAM: $([math]::Round($p.WorkingSet64 / 1MB, 0)) MB)")
  }
} else {
  $envLines.Add("Laufende Adobe-Prozesse: Keine aktiv")
}

try {
  $av = @(Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction Stop | Select-Object -ExpandProperty displayName)
  $envLines.Add("Installierte Antivirus-Software: $(if ($av.Count -gt 0) { $av -join ', ' } else { 'Keine gemeldet' })")
} catch {
  $envLines.Add("Installierte Antivirus-Software: Nicht abrufbar")
}
Add-Section "7. ADOBE- & SCHUTZPROGRAMM-UMGEBUNG" ($envLines -join "`r`n")

# 8. TOP RAM PROCESSES
$topProcs = Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 12 `
  ProcessName, Id,
  @{ N = 'RAM_MB'; E = { [math]::Round($_.WorkingSet64 / 1MB, 0) } },
  @{ N = 'CPU_Sek'; E = { if ($_.CPU) { [math]::Round($_.CPU, 1) } else { 0 } } }
Add-Section "8. TOP 12 PROZESSE NACH ARBEITSSPEICHER" ($topProcs | Format-Table -AutoSize | Out-String -Width 240)

# BUILD FINAL REPORT WITH SUMMARY AT TOP
$flagBlock = if ($flags.Count -gt 0) {
  ($flags -join "`r`n")
} else {
  "[OK] Keine offensichtlichen kritischen Auffaelligkeiten automatisch erkannt."
}

$header = @"
================================================================
AVICARY TRIAGE-KURZZUSAMMENFASSUNG (NUR-LESE-DIAGNOSE)
================================================================
$flagBlock

"@

$report = $header + ($sections -join "`r`n")

# Save UTF-8 with BOM so Windows 10 & 11 Notepad display Umlaute cleanly
if ($OutputPath) {
  $outPath = $OutputPath
} else {
  $desktop = [Environment]::GetFolderPath('Desktop')
  if (-not $desktop -or -not (Test-Path $desktop)) { $desktop = $env:USERPROFILE }
  $fileName = "Avicary-Bericht-$($now.ToString('yyyyMMdd-HHmm')).txt"
  $outPath = Join-Path $desktop $fileName
}

[System.IO.File]::WriteAllText($outPath, $report, (New-Object System.Text.UTF8Encoding($true)))

$copied = $false
if (-not $NoClipboard) {
  try {
    Set-Clipboard -Value $report -ErrorAction Stop
    $copied = $true
  } catch {}
}

Write-Host ""
Write-Host "=== KURZBEFUND ===" -ForegroundColor Yellow
if ($flags.Count -gt 0) {
  foreach ($f in $flags) {
    $color = if ($f -like '[KRITISCH]*') { 'Red' } elseif ($f -like '[WARNUNG]*') { 'Yellow' } else { 'Cyan' }
    Write-Host "  $f" -ForegroundColor $color
  }
} else {
  Write-Host "  [OK] Keine automatischen Warnungen." -ForegroundColor Green
}
Write-Host ""
Write-Host "Bericht gespeichert: $outPath" -ForegroundColor Green
if ($copied) {
  Write-Host "Bericht wurde ausserdem in die Zwischenablage kopiert (Strg+V)." -ForegroundColor Green
}
