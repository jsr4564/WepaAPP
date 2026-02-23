# Printer Supply & Tray Monitor

A desktop app for monitoring printer supplies and tray status from a web dashboard.

## Features

- Detects low toner/ink/drum/belt/fuser conditions
- Detects empty trays
- Classifies printers by service desk:
  - `Pattee Library Service Desk`
  - `Pollock Service Desk`
  - `Findlay Service Desk`
- Served-by filter for dashboard, alert history, and CSV export
- Tracks tray empty/filled history and low-supply alert lifecycle history
- Exports full alert history to CSV (tray + low-supply events)
- Generates copyable ServiceNow worknotes
- Generates copyable ServiceNow ticket descriptions
- Generates copyable printer-keys chat jokes (with occasional CS-major jokes)
- Sends native notifications on macOS, Windows, and Linux (when available)
- Includes macOS/Windows menu-bar summary views via app menu
- Includes a sleek minimal app icon for packaged builds

## Runtime Requirements

- Python 3.10+ (tested on 3.12)
- Tkinter available in your Python distribution (the app attempts a one-time auto-install if missing)

No third-party Python dependencies are required to run from source.

## Run From Source

```bash
python3 main.py
```

On first launch:
1. Paste your monitor URL into the `Monitor URL` field.
2. Click `Refresh Now`.

## Alert History

The `Alert History` tab includes:
- Empty tray events (`empty`, `filled`)
- Low-supply lifecycle events (`low_open`, `low_update`, `low_resolved`)

Double-click any current low-supply or empty-tray alert to open a generated incident report with recommended resolution steps.

## Data Storage (Portable)

The app stores state in a user-local app-data directory (not hardcoded machine paths):

- macOS: `~/Library/Application Support/printer_supply_tray_monitor/printer_monitor_state.json`
- Windows: `%APPDATA%\printer_supply_tray_monitor\printer_monitor_state.json`
- Linux: `${XDG_DATA_HOME:-~/.local/share}/printer_supply_tray_monitor/printer_monitor_state.json`

## Build Installers

Build scripts are included for:

- macOS `.pkg` + `.dmg` (installs to `/Applications`)
- Windows portable `.exe` + setup `.exe` (if Inno Setup is installed)

Install build dependency:

```bash
python3 -m pip install -r requirements-build.txt
```

### macOS

```bash
./packaging/build_macos_dmg.sh 1.0.0
```

Output:
- `release/macos/PrinterSupplyTrayMonitor-1.0.0.pkg`
- `release/macos/PrinterSupplyTrayMonitor-1.0.0.dmg`

### Windows (run on Windows)

```powershell
.\packaging\build_windows_exe.ps1 -Version 1.0.0
```

Output:
- `release\windows\PrinterSupplyTrayMonitor-1.0.0.exe` (portable)
- `release\windows\PrinterSupplyTrayMonitor-1.0.0-setup.exe` (installer, if Inno Setup is present)

## Credits

Credits: Jack Shetterly
