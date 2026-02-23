# Printer Key Checkout Tracker

Cross-platform Tkinter desktop app for sandbox IT key checkout/return tracking.

This build uses Google OAuth + Google Drive API (personal Gmail compatible).

Credits: Jack Shetterly

## Authentication Model

- Users sign in with Google in a browser window.
- Required first-run inputs:
  - Google OAuth Desktop App Client ID
  - Google OAuth Desktop App Client Secret
  - Shared Google Drive folder URL (or folder ID)
- Token cache and settings are stored in per-user app data (not in project folder).

## Google Drive Storage Layout

Inside the configured Drive folder, the app creates/uses:

```text
<drive folder>/
  logs/
    <userid>_events.csv
  output/
    keylog_aggregate.xlsx
```

CSV columns:
`EventId, Timestamp, UserId, Action, KeyId, FromLocation, ToLocation, PrinterOrDestination, ReturnedToLocation, Notes`

## Runtime Requirements

- Python 3.10+
- `google-api-python-client`
- `google-auth-oauthlib`
- `google-auth-httplib2`
- `openpyxl`

Install:

```bash
python3 -m pip install -r requirements.txt
```

## Run

macOS/Linux:

```bash
./run_mac_linux.sh
```

Windows (PowerShell):

```powershell
.\run_windows.bat
```

Direct run (advanced):

```bash
python3 app.py
```

## Startup Flow

1. Enter Google OAuth Client ID.
2. Enter Google OAuth Client Secret.
3. Enter shared Google Drive folder URL (or folder ID).
4. Complete Google sign-in in browser.
5. Enter session UserId.

Optional environment overrides:
- `KEY_TRACKER_GOOGLE_CLIENT_ID`
- `KEY_TRACKER_GOOGLE_CLIENT_SECRET`
- `KEY_TRACKER_GOOGLE_FOLDER_URL`
- `KEY_TRACKER_GOOGLE_FOLDER_ID`

## Logic Rules

- Event logs are append-only per-user CSV files.
- Double-checkout is blocked based on latest key state.
- Return for key not currently OUT requires confirmation.
- `What's Out` scans all CSV logs and computes latest event by `KeyId`.
- Malformed CSV rows are skipped and surfaced as warnings.

## Export

Export writes to Google Drive:
- `output/keylog_aggregate.xlsx`

Sheets:
- `RawEvents`
- `WhatsOut` (`KeyId`, `CheckedOutBy`, `TimeOut`, `ToLocation`, `PrinterOrDestination`)

## macOS DMG Build

Build app + DMG:

```bash
./build_macos_dmg.sh
```

Output:
- `dist/Printer Key Checkout Tracker.app`
- `dist/PrinterKeyCheckoutTracker-macOS.dmg`

## Portable Python Zip (Cross-Machine Source Bundle)

Create a distributable source zip:

```bash
./build_portable_zip.sh
```

Output:
- `release/PrinterKeyCheckoutTracker-python-portable.zip`

The zip includes only portable runtime files:
- `app.py`, `storage.py`, `models.py`, `utils.py`
- `requirements.txt`
- `README.md`
- `run_mac_linux.sh`, `run_windows.bat`
- `resources/AppIcon-256.png`

The release zip includes:
- Core app files (`app.py`, `storage.py`, `models.py`, `utils.py`)
- Run scripts (`run_mac_linux.sh`, `run_windows.bat`)
- Build scripts (`build_portable_zip.sh`, `build_github_release_zip.sh`)
- `requirements.txt`, `README.md`, icon asset, and wiki docs (`docs/wiki`)
