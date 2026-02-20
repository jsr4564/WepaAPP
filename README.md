# WepaAPP
# Printer Supply & Tray Monitor

A desktop app for monitoring printer supplies and tray status from a web dashboard.

## Features

- Detects low toner/ink/fuser conditions
- Detects empty trays
- Tracks tray empty/filled history over time
- Exports tray history to CSV
- Generates copyable ServiceNow worknotes
- Includes a sleek minimal app icon for packaged builds

## Runtime Requirements

- Python 3.10+ (tested on 3.12)
- Tkinter available in your Python distribution

No third-party Python dependencies are required to run from source.

## Run From Source

```bash
python3 main.py
```

On first launch:
1. Paste your monitor URL into the `Monitor URL` field.
2. Click `Refresh Now`.

## Minimal GitHub Upload (Source Run)

If you only want source-run support (`python3 main.py`), you can publish:

- `main.py`
- `README.md`

Optional (for nicer window icon):
- `assets/icons/PrinterSupplyTrayMonitor-256.png`

If the icon file is missing, the app still runs.

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
