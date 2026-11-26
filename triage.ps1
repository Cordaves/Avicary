<#
.SYNOPSIS
  Cordaves Avicary - Read-only Windows triage probe for remote support sessions.
.DESCRIPTION
  Collects system specs, RAM/pagefile pressure, internal & external USB drives
  (with physical-to-letter mapping and USB Selective Suspend check), disk health
  status, disk/USB/power system errors (14 days), app crashes/hangs (14 days),
  Reliability Monitor records, InDesign/CS6 bonus checks (recent .indd paths,
  missing/swapped drive letters, orphaned .idlk lock files, crash recovery files),
  backup readiness with a ready-to-copy safe robocopy command, and top RAM
  processes - strictly read-only, without modifying system settings or running
  slow recursive scans across external drives.

  Supports German (default, 'de') and English ('en') output via -Lang,
  $AvicaryLang, or $env:AVICARY_LANG.
  Saves a UTF-8 (BOM) report to the Desktop and copies it to the clipboard.
#>

[CmdletBinding()]
param(
  [int]$Days = 14,
  [string]$OutputPath = "",
  [switch]$NoClipboard,
  [ValidateSet('', 'de', 'en')]
  [string]$Lang = ""
)

$ErrorActionPreference = 'SilentlyContinue'

# Resolve report language: -Lang > $AvicaryLang > $env:AVICARY_LANG > 'de'
$resolvedLang = $Lang
if (-not $resolvedLang) {
  $varLang = Get-Variable -Name AvicaryLang -ValueOnly -ErrorAction SilentlyContinue
  if ($varLang) {
    $resolvedLang = "$varLang"
    Remove-Variable -Name AvicaryLang -ErrorAction SilentlyContinue
  } elseif ($env:AVICARY_LANG) {
    $resolvedLang = "$($env:AVICARY_LANG)"
  }
}
if ("$resolvedLang".ToLowerInvariant() -eq 'en') {
  $resolvedLang = 'en'
} else {
  $resolvedLang = 'de'
}
$isEn = ($resolvedLang -eq 'en')

$T = if ($isEn) {
  @{
    StartMsg            = "Avicary: Starting read-only diagnostics (last {0} days)..."
    FlagLowRam          = "[WARNING] Very low free memory: {0} / {1} GB ({2}% free)"
    FlagUptime          = "[NOTE] Last full reboot was {0} days ago (Windows Fast Startup active?)"
    LblCreated          = "Generated:"
    LblTimeZone         = "Time zone:"
    LblDevice           = "Device:"
    LblArch             = "Architecture:"
    LblCores            = "cores"
    LblLoad             = "Load:"
    RamFmt              = "{0} GB free of {1} GB ({2}% free)"
    LblPagefile         = "Pagefile:"
    PagefileFmt         = "{0} MB used / {1} MB allocated (Peak: {2} MB)"
    Unavailable         = "Unavailable"
    LblLastBoot         = "Last boot:"
    UptimeFmt           = "Uptime: {0} days"
    Sec1                = "1. SYSTEM & MEMORY"
    FlagLowC            = "[WARNING] System drive C: is low on free space: {0} GB ({1}% free)"
    FlagLowDrive        = "[WARNING] Drive {0} is nearly full: {1} GB of {2} GB free ({3}%)"
    Sec2                = "2. DRIVES & STORAGE SPACE"
    FlagDiskStatus      = "[CRITICAL] Disk '{0}' reports status '{1}'!"
    FlagStorageHealth   = "[CRITICAL] Storage health '{0}': {1} ({2})"
    StorageHealthSkip   = "(Extended Storage Health status may require Admin rights - skipped.)"
    UsbAcActive         = "ENABLED (may put USB HDD to sleep)"
    UsbDcActive         = "ENABLED"
    UsbDisabled         = "Disabled (OK)"
    UsbUnknown          = "Unknown ({0})"
    UsbHeader           = "--- USB Selective Suspend (Power Saving) ---"
    UsbAcLbl            = "Plugged in (AC):"
    UsbDcLbl            = "On battery (DC):"
    FlagUsbSleep        = "[NOTE] External USB drive connected and USB Selective Suspend (AC) is enabled (possible cause of freezes)"
    Sec3                = "3. PHYSICAL DISKS, USB MAPPING & POWER SAVING"
    FlagDiskErrors      = "[CRITICAL] Found {0} disk/USB/filesystem error(s) in the last {1} days!"
    FlagPowerErrors     = "[WARNING] {0} unexpected shutdown(s)/power loss events (Kernel-Power 41) in the last {1} days"
    TotalRelevant       = "Total relevant entries: {0}"
    SubFreq             = "--- Frequency ---"
    SubLatest25         = "--- Latest 25 Entries ---"
    ColMessage          = "Message"
    NoSysErrors         = "No disk/USB/NTFS/power errors in the last {0} days."
    SysLogErr           = "No entries found or System log unreadable ({0})."
    Sec4                = "4. SYSTEM EVENT LOG: DISK / USB / CRASHES ({0} DAYS)"
    FlagAppEvents       = "[NOTE] {0} application crash(es) or hang(s) in Application log ({1} days)"
    TotalEntries        = "Total entries: {0}"
    NoAppErrors         = "No application crashes or hangs in the last {0} days."
    AppLogErr           = "No entries found or Application log unreadable ({0})."
    Sec5                = "5. APPLICATION CRASHES & HANGS ({0} DAYS)"
    NoReliability       = "No recent entries in Reliability Monitor ({0} days)."
    ReliabilityErr      = "Reliability Monitor unavailable ({0})."
    Sec6                = "6. RELIABILITY MONITOR (LATEST 15)"
    AdobeFolder         = "Adobe folder ({0}): {1}"
    EmptyParens         = "(empty)"
    RunningAdobe        = "Running Adobe processes:"
    NoRunningAdobe      = "Running Adobe processes: None active"
    InstalledAv         = "Installed antivirus software: {0}"
    NoneReported        = "None reported"
    InDesignProfile     = "InDesign profile found: {0}"
    RecoveryLine        = "  -> Crash recovery ({0}): {1} file(s), latest {2}"
    DefaultsLine        = "  -> InDesign Defaults: {0} KB (modified: {1})"
    FlagRecovery        = "[WARNING] Found {0} InDesign crash recovery file(s) (indicates recent crashes/hangs)"
    StatusDriveMissing  = "DRIVE MISSING ({0})"
    StatusFileMissing   = "FILE MISSING"
    ModifiedWord        = "modified"
    FlagMissingDrives   = "[CRITICAL] Recently opened InDesign file(s) are on a disconnected drive: {0} (different USB HDD or changed drive letter?)"
    FlagLockFiles       = "[WARNING] Found {0} InDesign lock file(s) (.idlk) (can block files after an InDesign crash)"
    SubRecentLayout     = "--- Recently Opened Layout/InDesign Files (Windows Recent) ---"
    NoRecentLayout      = "No .indd/.idml/.psd/.pdf shortcuts found in Windows Recent."
    SubLockFiles        = "--- Found InDesign Lock Files (.idlk) ---"
    HelperTip           = "Helper tip (InDesign CS6): If images lag on the laptop -> check View -> Display Performance -> 'Typical Display' + export backup as .idml (File -> Export -> IDML)."
    Sec7                = "7. ADOBE / INDESIGN CS6 BONUS CHECK & ANTIVIRUS"
    OneDriveFolder      = "OneDrive folder: {0}"
    NotConfigured       = "Not configured"
    FreeWord            = "free"
    DrivesGt5GB         = "Drives with > 5 GB free space: {0}"
    NoneExcl            = "None!"
    DetectedProject     = "Detected current project folder: {0}"
    FilesInFolder       = "Files directly in this folder: {0} file(s), approx. {1} MB"
    BackupPrefix        = "Project-Backup"
    SafeCopyIntro       = "Safe copy command for the helper (not auto-executed; /R:1 /W:1 avoids hanging on bad sectors):"
    NoProjectDetected   = "No reachable .indd project folder detected automatically (plug in the external USB drive first if needed)."
    Sec8                = "8. BACKUP READINESS"
    Sec9                = "9. TOP 12 PROCESSES BY MEMORY"
    NoFlagsReport       = "[OK] No obvious critical issues automatically detected."
    ReportHeader        = "AVICARY TRIAGE SUMMARY (READ-ONLY DIAGNOSTICS)"
    FilePrefix          = "Avicary-Report"
    ConsoleSummaryTitle = "=== QUICK SUMMARY ==="
    ConsoleNoFlags      = "  [OK] No automatic warnings."
    ConsoleSaved        = "Report saved: {0}"
    ConsoleCopied       = "Report was also copied to the clipboard (Ctrl+V)."
  }
} else {
  @{
    StartMsg            = "Avicary: Starte Nur-Lese-Diagnose (letzte {0} Tage)..."
    FlagLowRam          = "[WARNUNG] Sehr wenig freier Arbeitsspeicher: {0} / {1} GB ({2}% frei)"
    FlagUptime          = "[HINWEIS] Letzter echter Neustart vor {0} Tagen (Windows-Schnellstart aktiv?)"
    LblCreated          = "Erstellt:"
    LblTimeZone         = "Zeitzone:"
    LblDevice           = "Geraet:"
    LblArch             = "Architektur:"
    LblCores            = "Kerne"
    LblLoad             = "Last:"
    RamFmt              = "{0} GB frei von {1} GB ({2}% frei)"
    LblPagefile         = "Auslagerung:"
    PagefileFmt         = "{0} MB belegt / {1} MB zugewiesen (Peak: {2} MB)"
    Unavailable         = "Nicht abrufbar"
    LblLastBoot         = "Letzter Boot:"
    UptimeFmt           = "Laufzeit: {0} Tage"
    Sec1                = "1. SYSTEM & ARBEITSSPEICHER"
    FlagLowC            = "[WARNUNG] Systemlaufwerk C: hat wenig freien Platz: {0} GB ({1}% frei)"
    FlagLowDrive        = "[WARNUNG] Laufwerk {0} fast voll: {1} GB von {2} GB frei ({3}%)"
    Sec2                = "2. LAUFWERKE & SPEICHERPLATZ"
    FlagDiskStatus      = "[KRITISCH] Festplatte '{0}' meldet Status '{1}'!"
    FlagStorageHealth   = "[KRITISCH] Storage-Health '{0}': {1} ({2})"
    StorageHealthSkip   = "(Erweiterter Storage-Health-Status benoetigt ggf. Admin-Rechte - uebersprungen.)"
    UsbAcActive         = "AKTIV (kann USB-HDD schlafen legen)"
    UsbDcActive         = "AKTIV"
    UsbDisabled         = "Deaktiviert (OK)"
    UsbUnknown          = "Unbekannt ({0})"
    UsbHeader           = "--- USB-Energiesparmodus (Selective Suspend) ---"
    UsbAcLbl            = "Netzbetrieb (AC):"
    UsbDcLbl            = "Akkubetrieb (DC):"
    FlagUsbSleep        = "[HINWEIS] Externe USB-Festplatte angeschlossen und USB-Energiesparmodus (AC) ist aktiv (moegliche Ursache fuer Haenger)"
    Sec3                = "3. PHYSISCHE FESTPLATTEN, USB-ZUORDNUNG & ENERGIESPARMODUS"
    FlagDiskErrors      = "[KRITISCH] {0} Festplatten-/USB-/Dateisystem-Fehler in den letzten {1} Tagen gefunden!"
    FlagPowerErrors     = "[WARNUNG] {0} unerwartete Abstuerze/Stromausfaelle (Kernel-Power 41) in den letzten {1} Tagen"
    TotalRelevant       = "Gesamt relevante Eintraege: {0}"
    SubFreq             = "--- Haeufigkeit ---"
    SubLatest25         = "--- Letzte 25 Eintraege ---"
    ColMessage          = "Meldung"
    NoSysErrors         = "Keine Disk-/USB-/NTFS-/Power-Fehler in den letzten {0} Tagen."
    SysLogErr           = "Keine Eintraege gefunden oder System-Log nicht lesbar ({0})."
    Sec4                = "4. SYSTEMPROTOKOLL: FESTPLATTE / USB / ABSTUERZE ({0} TAGE)"
    FlagAppEvents       = "[HINWEIS] {0} Programm-Abstuerze oder Haenger im Anwendungsprotokoll ({1} Tage)"
    TotalEntries        = "Gesamt Eintraege: {0}"
    NoAppErrors         = "Keine Anwendungs-Abstuerze oder Haenger in den letzten {0} Tagen."
    AppLogErr           = "Keine Eintraege gefunden oder Anwendungs-Log nicht lesbar ({0})."
    Sec5                = "5. PROGRAMM-ABSTUERZE & HAENGER ({0} TAGE)"
    NoReliability       = "Keine aktuellen Eintraege im Zuverlaessigkeitsverlauf ({0} Tage)."
    ReliabilityErr      = "Zuverlaessigkeitsverlauf nicht abrufbar ({0})."
    Sec6                = "6. ZUVERLAESSIGKEITSVERLAUF (LETZTE 15)"
    AdobeFolder         = "Adobe-Ordner ({0}): {1}"
    EmptyParens         = "(leer)"
    RunningAdobe        = "Laufende Adobe-Prozesse:"
    NoRunningAdobe      = "Laufende Adobe-Prozesse: Keine aktiv"
    InstalledAv         = "Installierte Antivirus-Software: {0}"
    NoneReported        = "Keine gemeldet"
    InDesignProfile     = "InDesign-Profil gefunden: {0}"
    RecoveryLine        = "  -> Absturz-Wiederherstellung ({0}): {1} Datei(en), zuletzt {2}"
    DefaultsLine        = "  -> InDesign Defaults: {0} KB (geaendert: {1})"
    FlagRecovery        = "[WARNUNG] {0} InDesign-Wiederherstellungsdatei(en) gefunden (Hinweis auf kuerzliche Abstuerze/Haenger)"
    StatusDriveMissing  = "LAUFWERK FEHLT ({0})"
    StatusFileMissing   = "DATEI FEHLT"
    ModifiedWord        = "geaendert"
    FlagMissingDrives   = "[KRITISCH] Kuerzlich geoeffnete InDesign-Datei(en) liegen auf nicht angeschlossenem Laufwerk: {0} (andere USB-HDD oder geaenderter Laufwerksbuchstabe?)"
    FlagLockFiles       = "[WARNUNG] {0} InDesign-Sperrdatei(en) (.idlk) gefunden (blockieren Dateien, falls InDesign abgestuerzt ist)"
    SubRecentLayout     = "--- Zuletzt geoeffnete Layout-/InDesign-Dateien (Windows-Verlauf) ---"
    NoRecentLayout      = "Keine .indd/.idml/.psd/.pdf-Verknuepfungen im Windows-Verlauf gefunden."
    SubLockFiles        = "--- Gefundene InDesign-Sperrdateien (.idlk) ---"
    HelperTip           = "Helfer-Tipp (InDesign CS6): Falls Bilder auf dem Notebook haengen -> Ansicht -> Anzeigeleistung -> 'Normale Anzeige' pruefen + Backup als .idml (Datei -> Exportieren -> IDML)."
    Sec7                = "7. ADOBE / INDESIGN CS6 BONUS-CHECK & SCHUTZPROGRAMME"
    OneDriveFolder      = "OneDrive-Ordner: {0}"
    NotConfigured       = "Nicht eingerichtet"
    FreeWord            = "frei"
    DrivesGt5GB         = "Laufwerke mit > 5 GB freiem Speicher: {0}"
    NoneExcl            = "Keine!"
    DetectedProject     = "Erkannter aktueller Projektordner: {0}"
    FilesInFolder       = "Dateien direkt in diesem Ordner: {0} Datei(en), ca. {1} MB"
    BackupPrefix        = "Chronik-Backup"
    SafeCopyIntro       = "Sicherer Kopierbefehl fuer den Helfer (nicht automatisch ausgefuehrt; /R:1 /W:1 haengt nicht bei defekten Sektoren):"
    NoProjectDetected   = "Kein erreichbarer .indd-Projektordner automatisch erkannt (ggf. zuerst richtige USB-Festplatte anstecken)."
    Sec8                = "8. BACKUP-BEREITSCHAFT (CHRONIK-SICHERUNG)"
    Sec9                = "9. TOP 12 PROZESSE NACH ARBEITSSPEICHER"
    NoFlagsReport       = "[OK] Keine offensichtlichen kritischen Auffaelligkeiten automatisch erkannt."
    ReportHeader        = "AVICARY TRIAGE-KURZZUSAMMENFASSUNG (NUR-LESE-DIAGNOSE)"
    FilePrefix          = "Avicary-Bericht"
    ConsoleSummaryTitle = "=== KURZBEFUND ==="
    ConsoleNoFlags      = "  [OK] Keine automatischen Warnungen."
    ConsoleSaved        = "Bericht gespeichert: {0}"
    ConsoleCopied       = "Bericht wurde ausserdem in die Zwischenablage kopiert (Strg+V)."
  }
}

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

Write-Host ($T.StartMsg -f $Days) -ForegroundColor Cyan

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
  $flags.Add(($T.FlagLowRam -f $freeRamGB, $totalRamGB, $freeRamPct))
}
if ($uptimeDays -gt 7) {
  $flags.Add(($T.FlagUptime -f $uptimeDays))
}

$pfText = if ($pf) { $T.PagefileFmt -f $pf.CurrentUsage, $pf.AllocatedBaseSize, $pf.PeakUsage } else { $T.Unavailable }
$sysLines = @(
  ("{0,-16} {1} ({2} {3})" -f $T.LblCreated, $now.ToString('yyyy-MM-dd HH:mm:ss'), $T.LblTimeZone, ([TimeZoneInfo]::Local.DisplayName))
  ("{0,-16} {1} {2}" -f $T.LblDevice, $cs.Manufacturer, $cs.Model)
  ("{0,-16} {1} | Version {2} (Build {3})" -f "OS:", $os.Caption, $os.Version, $os.BuildNumber)
  ("{0,-16} {1}" -f $T.LblArch, $os.OSArchitecture)
  ("{0,-16} {1} ({2} {3} / {4} Threads, {5} {6}%)" -f "CPU:", $cpu.Name, $cpu.NumberOfCores, $T.LblCores, $cpu.NumberOfLogicalProcessors, $T.LblLoad, $cpu.LoadPercentage)
  ("{0,-16} {1}" -f "RAM:", ($T.RamFmt -f $freeRamGB, $totalRamGB, $freeRamPct))
  ("{0,-16} {1}" -f $T.LblPagefile, $pfText)
  ("{0,-16} {1} ({2})" -f $T.LblLastBoot, $os.LastBootUpTime, ($T.UptimeFmt -f $uptimeDays))
)
Add-Section $T.Sec1 ($sysLines -join "`r`n")

# 2. LOGICAL DRIVES & FREE SPACE
$driveTypeMap = if ($isEn) {
  @{
    0 = 'Unknown'
    1 = 'No Root'
    2 = 'Removable/USB'
    3 = 'Local Disk'
    4 = 'Network'
    5 = 'CD/DVD'
    6 = 'RAM Disk'
  }
} else {
  @{
    0 = 'Unbekannt'
    1 = 'Kein Root'
    2 = 'Wechsel/USB'
    3 = 'Festplatte'
    4 = 'Netzlaufwerk'
    5 = 'CD/DVD'
    6 = 'RAM-Disk'
  }
}

$logicalRaw = @(Get-CimInstance Win32_LogicalDisk)
$mountedLetters = @($logicalRaw | ForEach-Object { $_.DeviceID })

$logical = $logicalRaw | ForEach-Object {
  $sizeGB = if ($_.Size) { [math]::Round($_.Size / 1GB, 1) } else { 0 }
  $freeGB = if ($_.FreeSpace) { [math]::Round($_.FreeSpace / 1GB, 1) } else { 0 }
  $freePct = if ($sizeGB -gt 0) { [math]::Round(($freeGB / $sizeGB) * 100, 0) } else { 0 }

  if ($_.DeviceID -eq 'C:' -and $sizeGB -gt 0 -and ($freeGB -lt 15 -or $freePct -lt 10)) {
    $flags.Add(($T.FlagLowC -f $freeGB, $freePct))
  }
  elseif ($_.DriveType -in 2, 3 -and $sizeGB -gt 0 -and $freePct -lt 8) {
    $flags.Add(($T.FlagLowDrive -f $_.DeviceID, $freeGB, $sizeGB, $freePct))
  }

  if ($isEn) {
    [PSCustomObject]@{
      Drive      = $_.DeviceID
      Label      = $_.VolumeName
      Type       = $driveTypeMap[[int]$_.DriveType]
      FileSystem = $_.FileSystem
      SizeGB     = $sizeGB
      FreeGB     = $freeGB
      FreePct    = "$freePct%"
    }
  } else {
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
}
Add-Section $T.Sec2 ($logical | Format-Table -AutoSize | Out-String -Width 240)

# 3. PHYSICAL DISKS, USB-TO-LETTER MAPPING & USB SLEEP CHECK
$physDisksRaw = @(Get-CimInstance Win32_DiskDrive)
$physDisks = $physDisksRaw | ForEach-Object {
  if ($_.Status -and $_.Status -ne 'OK') {
    $flags.Add(($T.FlagDiskStatus -f $_.Model, $_.Status))
  }
  $letters = @()
  try {
    $parts = Get-CimAssociatedInstance -InputObject $_ -ResultClassName Win32_DiskPartition -ErrorAction Stop
    foreach ($pt in $parts) {
      $lds = Get-CimAssociatedInstance -InputObject $pt -ResultClassName Win32_LogicalDisk -ErrorAction Stop
      foreach ($ld in $lds) { $letters += "$($ld.DeviceID) ($($ld.VolumeName))" }
    }
  } catch {}
  $lettersStr = if ($letters.Count -gt 0) { $letters -join ', ' } else { '-' }
  $sizeVal = if ($_.Size) { [math]::Round($_.Size / 1GB, 1) } else { 0 }
  if ($isEn) {
    [PSCustomObject]@{
      Index     = $_.Index
      Model     = $_.Model
      Interface = $_.InterfaceType
      Drives    = $lettersStr
      Status    = $_.Status
      SizeGB    = $sizeVal
    }
  } else {
    [PSCustomObject]@{
      Index     = $_.Index
      Modell    = $_.Model
      Anschluss = $_.InterfaceType
      Laufwerke = $lettersStr
      Status    = $_.Status
      GroesseGB = $sizeVal
    }
  }
}
$diskText = ($physDisks | Format-Table -AutoSize | Out-String -Width 240)

try {
  $msftDisks = Get-PhysicalDisk -ErrorAction Stop | Select-Object FriendlyName, BusType, MediaType, HealthStatus, OperationalStatus
  if ($msftDisks) {
    foreach ($d in $msftDisks) {
      if ($d.HealthStatus -and $d.HealthStatus -ne 'Healthy') {
        $flags.Add(($T.FlagStorageHealth -f $d.FriendlyName, $d.HealthStatus, $d.OperationalStatus))
      }
    }
    $diskText += "`r`n--- Windows Storage Health ---`r`n" + ($msftDisks | Format-Table -AutoSize | Out-String -Width 240)
  }
} catch {
  $diskText += "`r`n" + $T.StorageHealthSkip
}

# Read-only USB Selective Suspend check (active power plan)
try {
  $usbPow = (powercfg /Q SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 2>&1) -join "`n"
  $acMatch = [regex]::Match($usbPow, '(?m)AC.*?(0x[0-9a-fA-F]{8})|Netzbetrieb.*?(0x[0-9a-fA-F]{8})')
  $dcMatch = [regex]::Match($usbPow, '(?m)DC.*?(0x[0-9a-fA-F]{8})|Akkubetrieb.*?(0x[0-9a-fA-F]{8})')
  $allHex = @([regex]::Matches($usbPow, '0x[0-9a-fA-F]{8}') | ForEach-Object { $_.Value })
  $acVal = if ($acMatch.Success) { if ($acMatch.Groups[1].Value) { $acMatch.Groups[1].Value } else { $acMatch.Groups[2].Value } } elseif ($allHex.Count -ge 2) { $allHex[-2] } else { '' }
  $dcVal = if ($dcMatch.Success) { if ($dcMatch.Groups[1].Value) { $dcMatch.Groups[1].Value } else { $dcMatch.Groups[2].Value } } elseif ($allHex.Count -ge 1) { $allHex[-1] } else { '' }
  $acState = if ($acVal -eq '0x00000001') { $T.UsbAcActive } elseif ($acVal -eq '0x00000000') { $T.UsbDisabled } else { ($T.UsbUnknown -f $acVal) }
  $dcState = if ($dcVal -eq '0x00000001') { $T.UsbDcActive } elseif ($dcVal -eq '0x00000000') { $T.UsbDisabled } else { ($T.UsbUnknown -f $dcVal) }
  $diskText += "`r`n$($T.UsbHeader)`r`n$($T.UsbAcLbl) $acState`r`n$($T.UsbDcLbl) $dcState"
  $hasUsbDisk = @($physDisksRaw | Where-Object { $_.InterfaceType -eq 'USB' }).Count -gt 0
  if ($hasUsbDisk -and $acVal -eq '0x00000001') {
    $flags.Add($T.FlagUsbSleep)
  }
} catch {}
Add-Section $T.Sec3 $diskText

# 4. SYSTEM EVENT LOG: DISK / USB / NTFS / POWER / BSOD
$sysErrText = ""
try {
  $sysEvents = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; Level = 1, 2, 3; StartTime = $since } -ErrorAction Stop |
    Where-Object { $_.ProviderName -match 'disk|ntfs|storahci|storport|usb|volmgr|partmgr|Kernel-Power|nvme|volsnap|BugCheck|WHEA' })

  if ($sysEvents.Count -gt 0) {
    $diskErrCount = @($sysEvents | Where-Object { $_.ProviderName -match 'disk|ntfs|storahci|storport|usb|volmgr|nvme' }).Count
    if ($diskErrCount -gt 0) {
      $flags.Add(($T.FlagDiskErrors -f $diskErrCount, $Days))
    }
    $powerCount = @($sysEvents | Where-Object { $_.ProviderName -match 'Kernel-Power' -and $_.Id -eq 41 }).Count
    if ($powerCount -gt 0) {
      $flags.Add(($T.FlagPowerErrors -f $powerCount, $Days))
    }

    $summary = $sysEvents | Group-Object ProviderName, Id, LevelDisplayName | Sort-Object Count -Descending | ForEach-Object {
      $parts = $_.Name -split ', '
      if ($isEn) {
        [PSCustomObject]@{
          Count   = $_.Count
          Source  = $parts[0]
          EventID = $parts[1]
          Level   = $parts[2]
        }
      } else {
        [PSCustomObject]@{
          Anzahl   = $_.Count
          Quelle   = $parts[0]
          EventID  = $parts[1]
          Stufe    = $parts[2]
        }
      }
    }
    $sysErrText = ($T.TotalRelevant -f $sysEvents.Count) + "`r`n`r`n" + $T.SubFreq + "`r`n"
    $sysErrText += ($summary | Select-Object -First 15 | Format-Table -AutoSize | Out-String -Width 240)
    $sysErrText += "`r`n" + $T.SubLatest25 + "`r`n"
    $sysErrText += ($sysEvents | Select-Object -First 25 TimeCreated, LevelDisplayName, ProviderName, Id, @{ N = $T.ColMessage; E = { Clean-Msg $_.Message 145 } } | Format-Table -AutoSize | Out-String -Width 240)
  } else {
    $sysErrText = ($T.NoSysErrors -f $Days)
  }
} catch {
  $sysErrText = ($T.SysLogErr -f $_.Exception.Message)
}
Add-Section ($T.Sec4 -f $Days) $sysErrText

# 5. APPLICATION CRASHES, HANGS & ADOBE SIGNALS
$appErrText = ""
try {
  $appEvents = @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; Level = 1, 2; StartTime = $since } -ErrorAction Stop |
    Where-Object {
      $_.ProviderName -match 'Application Error|Application Hang|Windows Error Reporting|Adobe' -or
      "$($_.Message)" -match 'InDesign|Photoshop|Illustrator|Acrobat|Adobe|CS6'
    })

  if ($appEvents.Count -gt 0) {
    $flags.Add(($T.FlagAppEvents -f $appEvents.Count, $Days))
    $appErrText = ($T.TotalEntries -f $appEvents.Count) + "`r`n"
    $appErrText += ($appEvents | Select-Object -First 25 TimeCreated, ProviderName, Id, @{ N = $T.ColMessage; E = { Clean-Msg $_.Message 165 } } | Format-Table -AutoSize | Out-String -Width 240)
  } else {
    $appErrText = ($T.NoAppErrors -f $Days)
  }
} catch {
  $appErrText = ($T.AppLogErr -f $_.Exception.Message)
}
Add-Section ($T.Sec5 -f $Days) $appErrText

# 6. RELIABILITY MONITOR (ZUVERLAESSIGKEITSVERLAUF)
$relText = ""
try {
  $rel = @(Get-CimInstance -ClassName Win32_ReliabilityRecords -ErrorAction Stop |
    Where-Object { $_.TimeGenerated -ge $since } |
    Sort-Object TimeGenerated -Descending |
    Select-Object -First 15 TimeGenerated, SourceName, ProductName, @{ N = $T.ColMessage; E = { Clean-Msg $_.Message 140 } })
  if ($rel.Count -gt 0) {
    $relText = ($rel | Format-Table -AutoSize | Out-String -Width 240)
  } else {
    $relText = ($T.NoReliability -f $Days)
  }
} catch {
  $relText = ($T.ReliabilityErr -f $_.Exception.Message)
}
Add-Section $T.Sec6 $relText

# 7. ADOBE / CREATIVE SUITE DEEP CHECK (BONUS CHECKS: RECENT .INDD, MISSING DRIVES, .IDLK, RECOVERY)
$envLines = New-Object System.Collections.Generic.List[string]
foreach ($basePath in @("C:\Program Files\Adobe", "C:\Program Files (x86)\Adobe")) {
  if (Test-Path $basePath) {
    $dirs = (Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name) -join ", "
    $envLines.Add(($T.AdobeFolder -f $basePath, $(if ($dirs) { $dirs } else { $T.EmptyParens })))
  }
}

$adobeProcs = @(Get-Process | Where-Object { $_.ProcessName -match 'InDesign|Photoshop|Illustrator|Acrobat|Adobe|Creative|CS6|AAM|AGM|AGS' })
if ($adobeProcs.Count -gt 0) {
  $envLines.Add($T.RunningAdobe)
  foreach ($p in $adobeProcs) {
    $envLines.Add("  - $($p.ProcessName) (PID $($p.Id), RAM: $([math]::Round($p.WorkingSet64 / 1MB, 0)) MB)")
  }
} else {
  $envLines.Add($T.NoRunningAdobe)
}

try {
  $av = @(Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction Stop | Select-Object -ExpandProperty displayName)
  $envLines.Add(($T.InstalledAv -f $(if ($av.Count -gt 0) { $av -join ', ' } else { $T.NoneReported })))
} catch {
  $envLines.Add(($T.InstalledAv -f $T.Unavailable))
}

# Check InDesign Crash Recovery & Preferences folders (fast local C: only)
$recoveryFiles = @()
foreach ($idRoot in @("$env:LOCALAPPDATA\Adobe\InDesign", "$env:APPDATA\Adobe\InDesign")) {
  if (Test-Path $idRoot) {
    $vers = @(Get-ChildItem -Path $idRoot -Directory -ErrorAction SilentlyContinue)
    foreach ($v in $vers) {
      $envLines.Add(($T.InDesignProfile -f $v.FullName))
      $recDirs = @(Get-ChildItem -Path $v.FullName -Directory -Recurse -Depth 2 -Filter "InDesign Recovery" -ErrorAction SilentlyContinue)
      foreach ($rd in $recDirs) {
        $rf = @(Get-ChildItem -Path $rd.FullName -File -ErrorAction SilentlyContinue)
        if ($rf.Count -gt 0) {
          $recoveryFiles += $rf
          $latestRec = $rf | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty LastWriteTime
          $envLines.Add(($T.RecoveryLine -f $rd.FullName, $rf.Count, $latestRec))
        }
      }
      $defFiles = @(Get-ChildItem -Path $v.FullName -File -Recurse -Depth 2 -Filter "InDesign Defaults" -ErrorAction SilentlyContinue)
      foreach ($df in $defFiles) {
        $envLines.Add(($T.DefaultsLine -f [math]::Round($df.Length / 1KB, 0), $df.LastWriteTime))
      }
    }
  }
}
if ($recoveryFiles.Count -gt 0) {
  $flags.Add(($T.FlagRecovery -f $recoveryFiles.Count))
}

# Resolve Windows Recent shortcuts (.lnk) for InDesign / layout / photo files (instant, zero disk scan!)
$recentDir = Join-Path $env:APPDATA "Microsoft\Windows\Recent"
$recentProjects = New-Object System.Collections.Generic.List[object]
$primaryInddPath = $null
$primaryInddDrive = $null
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
          if (-not $primaryInddPath -and $exists -and $target -match '\.(indd|indb)$') {
            $primaryInddPath = $target
            $primaryInddDrive = $drv
          }

          $statusText = if ($exists) { "OK ($sizeMB MB)" } elseif (-not $driveMounted) { ($T.StatusDriveMissing -f $drv) } else { $T.StatusFileMissing }
          if ($isEn) {
            $recentProjects.Add([PSCustomObject]@{
              Recent = $lnk.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
              Drive  = $drv
              Status = $statusText
              Path   = $target
            })
          } else {
            $recentProjects.Add([PSCustomObject]@{
              Zuletzt  = $lnk.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
              Laufwerk = $drv
              Status   = $statusText
              Pfad     = $target
            })
          }

          # Check exact parent folder (non-recursive!) for orphaned .idlk lock files
          if ($exists -and $target -match '\.(indd|indb)$') {
            $parent = Split-Path -Path $target -Parent
            if ($parent -and -not $checkedDirs.ContainsKey($parent)) {
              $checkedDirs[$parent] = $true
              $idlks = @(Get-ChildItem -LiteralPath $parent -Filter "*.idlk" -File -ErrorAction SilentlyContinue)
              foreach ($lk in $idlks) {
                $lockFilesFound.Add("$($lk.FullName) ($($T.ModifiedWord): $($lk.LastWriteTime))")
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
  $flags.Add(($T.FlagMissingDrives -f $drvList))
}
if ($lockFilesFound.Count -gt 0) {
  $flags.Add(($T.FlagLockFiles -f $lockFilesFound.Count))
}

$envLines.Add("")
$envLines.Add($T.SubRecentLayout)
if ($recentProjects.Count -gt 0) {
  $envLines.Add(($recentProjects | Select-Object -First 15 | Format-Table -AutoSize | Out-String -Width 240).TrimEnd())
} else {
  $envLines.Add($T.NoRecentLayout)
}

if ($lockFilesFound.Count -gt 0) {
  $envLines.Add("")
  $envLines.Add($T.SubLockFiles)
  foreach ($lf in $lockFilesFound) { $envLines.Add("  - $lf") }
}

$envLines.Add("")
$envLines.Add($T.HelperTip)
Add-Section $T.Sec7 ($envLines -join "`r`n")

# 8. BACKUP READINESS & READY-TO-COPY SAFE ROBOCOPY COMMAND (NEVER AUTO-RUNS)
$bkLines = New-Object System.Collections.Generic.List[string]

$oneDrivePath = if ($env:OneDrive -and (Test-Path $env:OneDrive)) { $env:OneDrive } else { $null }
$bkLines.Add(($T.OneDriveFolder -f $(if ($oneDrivePath) { $oneDrivePath } else { $T.NotConfigured })))

$safeTargets = @($logicalRaw | Where-Object { $_.DriveType -in 2, 3 -and $_.FreeSpace -gt 5GB } | ForEach-Object {
  "$($_.DeviceID) '$($_.VolumeName)' ($([math]::Round($_.FreeSpace / 1GB, 1)) GB $($T.FreeWord))"
})
$bkLines.Add(($T.DrivesGt5GB -f $(if ($safeTargets.Count -gt 0) { $safeTargets -join ' | ' } else { $T.NoneExcl })))

if ($primaryInddPath) {
  $srcFolder = Split-Path -Path $primaryInddPath -Parent
  $srcFiles = @(Get-ChildItem -LiteralPath $srcFolder -File -ErrorAction SilentlyContinue)
  $srcMB = [math]::Round((($srcFiles | Measure-Object -Property Length -Sum).Sum) / 1MB, 1)
  $bkLines.Add("")
  $bkLines.Add(($T.DetectedProject -f $srcFolder))
  $bkLines.Add(($T.FilesInFolder -f $srcFiles.Count, $srcMB))

  $altDrive = $logicalRaw | Where-Object { $_.DriveType -in 2, 3 -and $_.DeviceID -ne $primaryInddDrive -and $_.FreeSpace -gt 5GB } | Select-Object -First 1
  $backupName = "$($T.BackupPrefix)-$($now.ToString('yyyy-MM-dd'))"
  $dstRoot = if ($altDrive) { Join-Path "$($altDrive.DeviceID)\" $backupName }
             elseif ($oneDrivePath) { Join-Path $oneDrivePath $backupName }
             else { Join-Path $env:USERPROFILE "Desktop\$backupName" }

  $bkLines.Add("")
  $bkLines.Add($T.SafeCopyIntro)
  $bkLines.Add("  robocopy `"$srcFolder`" `"$dstRoot`" /E /R:1 /W:1 /NP /MT:4")
} else {
  $bkLines.Add($T.NoProjectDetected)
}
Add-Section $T.Sec8 ($bkLines -join "`r`n")

# 9. TOP RAM PROCESSES
$cpuColName = if ($isEn) { 'CPU_Sec' } else { 'CPU_Sek' }
$topProcs = Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 12 `
  ProcessName, Id,
  @{ N = 'RAM_MB'; E = { [math]::Round($_.WorkingSet64 / 1MB, 0) } },
  @{ N = $cpuColName; E = { if ($_.CPU) { [math]::Round($_.CPU, 1) } else { 0 } } }
Add-Section $T.Sec9 ($topProcs | Format-Table -AutoSize | Out-String -Width 240)

# BUILD FINAL REPORT WITH SUMMARY AT TOP
$flagBlock = if ($flags.Count -gt 0) {
  ($flags -join "`r`n")
} else {
  $T.NoFlagsReport
}

$header = @"
================================================================
$($T.ReportHeader)
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
  $fileName = "$($T.FilePrefix)-$($now.ToString('yyyyMMdd-HHmm')).txt"
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
Write-Host $T.ConsoleSummaryTitle -ForegroundColor Yellow
if ($flags.Count -gt 0) {
  foreach ($f in $flags) {
    $color = if ($f -like '[KRITISCH]*' -or $f -like '[CRITICAL]*') { 'Red' } elseif ($f -like '[WARNUNG]*' -or $f -like '[WARNING]*') { 'Yellow' } else { 'Cyan' }
    Write-Host "  $f" -ForegroundColor $color
  }
} else {
  Write-Host $T.ConsoleNoFlags -ForegroundColor Green
}
Write-Host ""
Write-Host ($T.ConsoleSaved -f $outPath) -ForegroundColor Green
if ($copied) {
  Write-Host $T.ConsoleCopied -ForegroundColor Green
}
