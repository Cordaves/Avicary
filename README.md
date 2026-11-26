<p align="left">
  <img src="avicary-logo.svg" alt="Avicary Logo" width="56" height="56" />
</p>

# Avicary (Cordaves)

**Remote IT Diagnosis Tool** — *read-only remote-support landing page & Windows triage probe, part of the Cordaves ecosystem alongside **Wingward** (planned remediation & watchdog companion).*

Styled with the Cordaves palette (**Deep blue `#0135A8`** main accent — Logo Blade 2, **Paper `#F8FAF7`** light ground, and the four angular Cordaves wing blades in **Deep green `#02592A`**, **Deep blue `#0135A8`**, **Deep violet `#531395`**, and **Crimson `#B0122E`**).

## Funktionsweise

1. **1-Klick-Start der eingebauten Windows-Remotehilfe (`index.html`)**
   - Nutzt das offizielle Windows-Protokoll `ms-quick-assist:`.
   - Kein Download, keine Installation, keine SmartScreen-Warnung.
   - Enthält einen rein lokalen Browser-Kurzcheck (OS, Bildschirm-Skalierung, Grafik, Systemzeit) sowie Fallback-Anleitungen und Direktlinks zu Windows-Einstellungen (`ms-settings:`).
2. **Nur-Lese-Diagnose für den Helfer (`triage.ps1`)**
   - Erst ausführen, **nachdem** Sie per Remotehilfe verbunden sind (PowerShell öffnen, kein Admin nötig):
     ```powershell
     irm 'https://cordaves.github.io/Avicary/triage.ps1' | iex
     ```
   - Prüft Arbeitsspeicher, Auslagerungsdatei, interne & externe USB-Laufwerke, Festplatten-Status, `disk`/`ntfs`/`usb`/`Kernel-Power`-Ereignisse (14 Tage), Programm-Abstürze/-Hänger, Zuverlässigkeitsverlauf, laufende Adobe-Prozesse und Top-RAM-Verbraucher.
   - Verändert nichts am System und durchsucht keine persönlichen Benutzerdateien.
   - Speichert `Avicary-Bericht-YYYYMMDD-HHMM.txt` auf dem Desktop und kopiert den Bericht in die Zwischenablage.

## Auf GitHub Pages veröffentlichen

1. Repository auf GitHub unter `https://github.com/Cordaves/Avicary` anlegen (**Public** für kostenloses GitHub Pages).
2. Lokal pushen:
   ```powershell
   git remote add origin https://github.com/Cordaves/Avicary.git
   git push -u origin main
   ```
3. In GitHub unter **Settings → Pages**:
   - **Source:** `Deploy from a branch`
   - **Branch:** `main` / `/ (root)` → **Save**
4. Live-URL nach ca. 1 Minute:
   `https://cordaves.github.io/Avicary/`

## Datenschutz & Sicherheit

- Enthält bewusst keine personenbezogenen Daten, keine Cookies und kein Tracking.
- Der Browser-Kurzcheck sendet nichts automatisch an Server.
- `triage.ps1` ist vollständig transparent und rein lesend – führen Sie es nur selbst während einer Remotehilfe-Sitzung aus.
