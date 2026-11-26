<#
.SYNOPSIS
  Cordaves Avicary — Read-only Windows triage probe for remote support sessions.
.DESCRIPTION
  Collects system specs, RAM/pagefile pressure, internal & external USB drives
  (with physical-to-letter mapping and USB Selective Suspend check), disk health
  status, disk/USB/power system errors (14 days), app crashes/hangs (14 days),
  Reliability Monitor records, InDesign/CS6 bonus checks (recent .indd paths,
  missing/swapped drive letters, orphaned .idlk lock files, crash recovery files),
  backup readiness with a ready-to-copy safe robocopy command, and top RAM
  processes — strictly read-only, without modifying system settings or running
  slow recursive scans across external drives.

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

$logicalRaw = @(Get-CimInstance Win32_LogicalDisk)
$mountedLetters = @($logicalRaw | ForEach-Object { $_.DeviceID })

$logical = $logicalRaw | ForEach-Object {
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

# 3. PHYSICAL DISKS, USB-TO-LETTER MAPPING & USB SLEEP CHECK
$physDisksRaw = @(Get-CimInstance Win32_DiskDrive)
$physDisks = $physDisksRaw | ForEach-Object {
  if ($_.Status -and $_.Status -ne 'OK') {
    $flags.Add("[KRITISCH] Festplatte '$($_.Model)' meldet Status '$($_.Status)'!")
  }
  $letters = @()
  try {
    $parts = Get-CimAssociatedInstance -InputObject $_ -ResultClassName Win32_DiskPartition -ErrorAction Stop
    foreach ($pt in $parts) {
      $lds = Get-CimAssociatedInstance -InputObject $pt -ResultClassName Win32_LogicalDisk -ErrorAction Stop
      foreach ($ld in $lds) { $letters += "$($ld.DeviceID) ($($ld.VolumeName))" }
    }
  } catch {}
  [PSCustomObject]@{
    Index     = $_.Index
    Modell    = $_.Model
    Anschluss = $_.InterfaceType
    Laufwerke = if ($letters.Count -gt 0) { $letters -join ', ' } else { '-' }
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

# Read-only USB Selective Suspend check (active power plan)
try {
  $usbPow = (powercfg /Q SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 2>&1) -join "`n"
  $acMatch = [regex]::Match($usbPow, '(?m)AC.*?(0x[0-9a-fA-F]{8})|Netzbetrieb.*?(0x[0-9a-fA-F]{8})')
  $dcMatch = [regex]::Match($usbPow, '(?m)DC.*?(0x[0-9a-fA-F]{8})|Akkubetrieb.*?(0x[0-9a-fA-F]{8})')
  $allHex = @([regex]::Matches($usbPow, '0x[0-9a-fA-F]{8}') | ForEach-Object { $_.Value })
  $acVal = if ($acMatch.Success) { if ($acMatch.Groups[1].Value) { $acMatch.Groups[1].Value } else { $acMatch.Groups[2].Value } } elseif ($allHex.Count -ge 2) { $allHex[-2] } else { '' }
  $dcVal = if ($dcMatch.Success) { if ($dcMatch.Groups[1].Value) { $dcMatch.Groups[1].Value } else { $dcMatch.Groups[2].Value } } elseif ($allHex.Count -ge 1) { $allHex[-1] } else { '' }
  $acState = if ($acVal -eq '0x00000001') { 'AKTIV (kann USB-HDD schlafen legen)' } elseif ($acVal -eq '0x00000000') { 'Deaktiviert (OK)' } else { "Unbekannt ($acVal)" }
  $dcState = if ($dcVal -eq '0x00000001') { 'AKTIV' } elseif ($dcVal -eq '0x00000000') { 'Deaktiviert (OK)' } else { "Unbekannt ($dcVal)" }
  $diskText += "`r`n--- USB-Energiesparmodus (Selective Suspend) ---`r`nNetzbetrieb (AC): $acState`r`nAkkubetrieb (DC): $dcState"
  $hasUsbDisk = @($physDisksRaw | Where-Object { $_.InterfaceType -eq 'USB' }).Count -gt 0
  if ($hasUsbDisk -and $acVal -eq '0x00000001') {
    $flags.Add("[HINWEIS] Externe USB-Festplatte angeschlossen und USB-Energiesparmodus (AC) ist aktiv (moegliche Ursache fuer Haenger)")
  }
} catch {}
Add-Section "3. PHYSISCHE FESTPLATTEN, USB-ZUORDNUNG & ENERGIESPARMODUS" $diskText

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

# 7. ADOBE / CREATIVE SUITE DEEP CHECK (BONUS CHECKS: RECENT .INDD, MISSING DRIVES, .IDLK, RECOVERY)
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

# Check InDesign Crash Recovery & Preferences folders (fast local C: only)
$recoveryFiles = @()
foreach ($idRoot in @("$env:LOCALAPPDATA\Adobe\InDesign", "$env:APPDATA\Adobe\InDesign")) {
  if (Test-Path $idRoot) {
    $vers = @(Get-ChildItem -Path $idRoot -Directory -ErrorAction SilentlyContinue)
    foreach ($v in $vers) {
      $envLines.Add("InDesign-Profil gefunden: $($v.FullName)")
      $recDirs = @(Get-ChildItem -Path $v.FullName -Directory -Recurse -Depth 2 -Filter "InDesign Recovery" -ErrorAction SilentlyContinue)
      foreach ($rd in $recDirs) {
        $rf = @(Get-ChildItem -Path $rd.FullName -File -ErrorAction SilentlyContinue)
        if ($rf.Count -gt 0) {
          $recoveryFiles += $rf
          $envLines.Add("  -> Absturz-Wiederherstellung ($($rd.FullName)): $($rf.Count) Datei(en), zuletzt $($rf | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty LastWriteTime)")
        }
      }
      $defFiles = @(Get-ChildItem -Path $v.FullName -File -Recurse -Depth 2 -Filter "InDesign Defaults" -ErrorAction SilentlyContinue)
      foreach ($df in $defFiles) {
        $envLines.Add("  -> InDesign Defaults: $([math]::Round($df.Length / 1KB, 0)) KB (geaendert: $($df.LastWriteTime))")
      }
    }
  }
}
if ($recoveryFiles.Count -gt 0) {
  $flags.Add("[WARNUNG] $($recoveryFiles.Count) InDesign-Wiederherstellungsdatei(en) gefunden (Hinweis auf kuerzliche Abstuerze/Haenger)")
}

# Resolve Windows Recent shortcuts (.lnk) for InDesign / layout / photo files (instant, zero disk scan!)
$recentDir = Join-Path $env:APPDATA "Microsoft\Windows\Recent"
$recentProjects = New-Object System.Collections.Generic.List[object]
$checkedDirs = @{}
$lockFilesFound = New-Object System.Collections.Generic.List[string]
$missingDrives = New-Object System.Collections.Generic.HashSet[string]

if (Test-Path $recentDir) {
  try {
    $wsh = New-Object -ComObject WScript.Shell
    $lnks = @(Get-ChildItem -Path $recentDir -Filter "*.lnk" -ErrorAction SilentlyContinue |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 120)

    foreach ($lnk in $lnks) {
      try {
        $target = $wsh.CreateShortcut($lnk.FullName).TargetPath
        if ($target -and ($target -match '\.(indd|idml|indt|indb|psd|ai|pdf)$')) {
          $drv = if ($target -match '^([A-Za-z]:)') { $Matches[1].ToUpper() } else { '' }
          $driveMounted = if ($drv) { $mountedLetters -contains $drv } else { $true }
          $exists = if ($driveMounted) { Test-Path -LiteralPath $target } else { $false }
          $sizeMB = if ($exists) { [math]::Round(((Get-Item -LiteralPath $target -ErrorAction SilentlyContinue).Length) / 1MB, 1) } else { $null }

          if (-not $driveMounted -and $drv -and $target -match '\.(indd|idml|indb)$') {
            [void]$missingDrives.Add($drv)
          }

          $recentProjects.Add([PSCustomObject]@{
            Zuletzt   = $lnk.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
            Laufwerk  = $drv
            Status    = if ($exists) { "OK ($sizeMB MB)" } elseif (-not $driveMounted) { "LAUFWERK FEHLT ($drv)" } else { "DATEI FEHLT" }
            Pfad      = $target
          })

          # Check exact parent folder (non-recursive!) for orphaned .idlk lock files
          if ($exists -and $target -match '\.(indd|indb)$') {
            $parent = Split-Path -Path $target -Parent
            if ($parent -and -not $checkedDirs.ContainsKey($parent)) {
              $checkedDirs[$parent] = $true
              $idlks = @(Get-ChildItem -LiteralPath $parent -Filter "*.idlk" -File -ErrorAction SilentlyContinue)
              foreach ($lk in $idlks) {
                $lockFilesFound.Add("$($lk.FullName) (geaendert: $($lk.LastWriteTime))")
              }
            }
          }
        }
      } catch {}
    }
  } catch {}
}

if ($missingDrives.Count -gt 0) {
  $drvList = ($missingDrives | Sort-Object) -join ', '
  $flags.Add("[KRITISCH] Kuerzlich geoeffnete InDesign-Datei(en) liegen auf nicht angeschlossenem Laufwerk: $drvList (andere USB-HDD oder geaenderter Laufwerksbuchstabe?)")
}
if ($lockFilesFound.Count -gt 0) {
  $flags.Add("[WARNUNG] $($lockFilesFound.Count) InDesign-Sperrdatei(en) (.idlk) gefunden (blockieren Dateien, falls InDesign abgestuerzt ist)")
}

$envLines.Add("")
$envLines.Add("--- Zuletzt geoeffnete Layout-/InDesign-Dateien (Windows-Verlauf) ---")
if ($recentProjects.Count -gt 0) {
  $envLines.Add(($recentProjects | Select-Object -First 15 | Format-Table -AutoSize | Out-String -Width 240).TrimEnd())
} else {
  $envLines.Add("Keine .indd/.idml/.psd/.pdf-Verknuepfungen im Windows-Verlauf gefunden.")
}

if ($lockFilesFound.Count -gt 0) {
  $envLines.Add("")
  $envLines.Add("--- Gefundene InDesign-Sperrdateien (.idlk) ---")
  foreach ($lf in $lockFilesFound) { $envLines.Add("  - $lf") }
}

$envLines.Add("")
$envLines.Add("Helfer-Tipp (InDesign CS6): Falls Bilder auf dem Notebook haengen -> Ansicht -> Anzeigeleistung -> 'Normale Anzeige' pruefen + Backup als .idml (Datei -> Exportieren -> IDML).")
Add-Section "7. ADOBE / INDESIGN CS6 BONUS-CHECK & SCHUTZPROGRAMME" ($envLines -join "`r`n")

# 8. BACKUP READINESS & READY-TO-COPY SAFE ROBOCOPY COMMAND (NEVER AUTO-RUNS)
$bkLines = New-Object System.Collections.Generic.List[string]

$oneDrivePath = if ($env:OneDrive -and (Test-Path $env:OneDrive)) { $env:OneDrive } else { $null }
$bkLines.Add("OneDrive-Ordner: $(if ($oneDrivePath) { $oneDrivePath } else { 'Nicht eingerichtet' })")

$safeTargets = @($logicalRaw | Where-Object { $_.DriveType -in 2, 3 -and $_.FreeSpace -gt 5GB } | ForEach-Object {
  "$($_.DeviceID) '$($_.VolumeName)' ($([math]::Round($_.FreeSpace / 1GB, 1)) GB frei)"
})
$bkLines.Add("Laufwerke mit > 5 GB freiem Speicher: $(if ($safeTargets.Count -gt 0) { $safeTargets -join ' | ' } else { 'Keine!' })")

$primaryIndd = $recentProjects | Where-Object { $_.Pfad -match '\.(indd|indb)$' -and $_.Status -like 'OK*' } | Select-Object -First 1
if ($primaryIndd) {
  $srcFolder = Split-Path -Path $primaryIndd.Pfad -Parent
  $srcFiles = @(Get-ChildItem -LiteralPath $srcFolder -File -ErrorAction SilentlyContinue)
  $srcMB = [math]::Round((($srcFiles | Measure-Object -Property Length -Sum).Sum) / 1MB, 1)
  $bkLines.Add("")
  $bkLines.Add("Erkannter aktueller Projektordner: $srcFolder")
  $bkLines.Add("Dateien direkt in diesem Ordner: $($srcFiles.Count) Datei(en), ca. $srcMB MB")

  $srcDrive = $primaryIndd.Laufwerk
  $altDrive = $logicalRaw | Where-Object { $_.DriveType -in 2, 3 -and $_.DeviceID -ne $srcDrive -and $_.FreeSpace -gt 5GB } | Select-Object -First 1
  $dstRoot = if ($altDrive) { Join-Path "$($altDrive.DeviceID)\" "Chronik-Backup-$($now.ToString('yyyy-MM-dd'))" }
             elseif ($oneDrivePath) { Join-Path $oneDrivePath "Chronik-Backup-$($now.ToString('yyyy-MM-dd'))" }
             else { Join-Path $env:USERPROFILE "Desktop\Chronik-Backup-$($now.ToString('yyyy-MM-dd'))" }

  $bkLines.Add("")
  $bkLines.Add("Sicherer Kopierbefehl fuer den Helfer (nicht automatisch ausgefuehrt; /R:1 /W:1 haengt nicht bei defekten Sektoren):")
  $bkLines.Add("  robocopy `"$srcFolder`" `"$dstRoot`" /E /R:1 /W:1 /NP /MT:4")
} else {
  $bkLines.Add("Kein erreichbarer .indd-Projektordner automatisch erkannt (ggf. zuerst richtige USB-Festplatte anstecken).")
}
Add-Section "8. BACKUP-BEREITSCHAFT (CHRONIK-SICHERUNG)" ($bkLines -join "`r`n")

# 9. TOP RAM PROCESSES
$topProcs = Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 12 `
  ProcessName, Id,
  @{ N = 'RAM_MB'; E = { [math]::Round($_.WorkingSet64 / 1MB, 0) } },
  @{ N = 'CPU_Sek'; E = { if ($_.CPU) { [math]::Round($_.CPU, 1) } else { 0 } } }
Add-Section "9. TOP 12 PROZESSE NACH ARBEITSSPEICHER" ($topProcs | Format-Table -AutoSize | Out-String -Width 240)

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
