"""Google Drive-backed storage and aggregation logic for printer key events."""

from __future__ import annotations

import csv
import io
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable
from urllib.parse import parse_qs, urlparse

try:
    from google.auth.transport.requests import Request as GoogleAuthRequest
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    from googleapiclient.http import MediaIoBaseUpload
except ImportError:  # pragma: no cover - dependency guard
    GoogleAuthRequest = None  # type: ignore[assignment]
    Credentials = None  # type: ignore[assignment]
    InstalledAppFlow = None  # type: ignore[assignment]
    build = None  # type: ignore[assignment]
    HttpError = None  # type: ignore[assignment]
    MediaIoBaseUpload = None  # type: ignore[assignment]

from models import EVENT_COLUMNS, KeyEvent, OutstandingKey
from utils import clean_text, new_event_id, normalize_key_id, now_iso_timestamp, parse_iso_timestamp

APP_ID = "PrinterKeyCheckoutTracker"

LOGS_DIRNAME = "logs"
OUTPUT_DIRNAME = "output"
AGGREGATE_FILENAME = "keylog_aggregate.xlsx"

GOOGLE_SCOPES = ["https://www.googleapis.com/auth/drive"]
GOOGLE_FOLDER_MIME = "application/vnd.google-apps.folder"

ENV_GOOGLE_CLIENT_ID = "KEY_TRACKER_GOOGLE_CLIENT_ID"
ENV_GOOGLE_CLIENT_SECRET = "KEY_TRACKER_GOOGLE_CLIENT_SECRET"
ENV_GOOGLE_FOLDER_URL = "KEY_TRACKER_GOOGLE_FOLDER_URL"
ENV_GOOGLE_FOLDER_ID = "KEY_TRACKER_GOOGLE_FOLDER_ID"

DEFAULT_GOOGLE_FOLDER_URL = ""

_REQUIRED_COLUMNS = {"EventId", "Timestamp", "UserId", "Action", "KeyId"}
_VALID_ACTIONS = {"OUT", "IN"}
_FALLBACK_MIN_TS = datetime(1970, 1, 1, tzinfo=timezone.utc)
_FALLBACK_MAX_TS = datetime(9999, 12, 31, 23, 59, 59, tzinfo=timezone.utc)
_DRIVE_FOLDER_ID_RE = re.compile(r"^[A-Za-z0-9_-]{10,}$")


class StorageError(Exception):
    """Raised when config, auth, log, or export operations fail."""


def _app_data_dir() -> Path:
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / APP_ID

    if os.name == "nt":
        appdata = clean_text(os.environ.get("APPDATA"))
        if appdata:
            return Path(appdata) / APP_ID
        return Path.home() / "AppData" / "Roaming" / APP_ID

    xdg = clean_text(os.environ.get("XDG_CONFIG_HOME"))
    if xdg:
        return Path(xdg) / APP_ID
    return Path.home() / ".config" / APP_ID


def _config_path() -> Path:
    return _app_data_dir() / "config.json"


def _token_path() -> Path:
    return _app_data_dir() / "google_token.json"


def load_config() -> dict[str, str]:
    config_path = _config_path()
    if not config_path.exists():
        return {}

    try:
        payload = json.loads(config_path.read_text(encoding="utf-8"))
    except OSError as exc:
        raise StorageError(f"Unable to read config file: {config_path}\n{exc}") from exc
    except json.JSONDecodeError as exc:
        raise StorageError(f"Invalid JSON in config file: {config_path}\n{exc}") from exc

    if not isinstance(payload, dict):
        raise StorageError(f"Config must contain a JSON object: {config_path}")

    return {str(key): str(value) for key, value in payload.items() if value is not None}


def save_config(config: dict[str, str]) -> None:
    config_path = _config_path()
    temp_path = config_path.with_suffix(".tmp")
    cleaned = {str(key): str(value) for key, value in config.items() if clean_text(str(value))}

    try:
        config_path.parent.mkdir(parents=True, exist_ok=True)
        temp_path.write_text(json.dumps(cleaned, indent=2), encoding="utf-8")
        temp_path.replace(config_path)
    except OSError as exc:
        raise StorageError(f"Unable to write config file: {config_path}\n{exc}") from exc


def save_setup(
    google_client_id: str,
    google_client_secret: str,
    google_folder_url: str,
    google_folder_id: str = "",
    email_hint: str = "",
) -> None:
    save_config(
        {
            "google_client_id": clean_text(google_client_id),
            "google_client_secret": clean_text(google_client_secret),
            "google_folder_url": clean_text(google_folder_url),
            "google_folder_id": clean_text(google_folder_id),
            "google_email_hint": clean_text(email_hint),
        }
    )


def get_saved_google_client_id() -> str:
    return clean_text(load_config().get("google_client_id"))


def get_saved_google_client_secret() -> str:
    return clean_text(load_config().get("google_client_secret"))


def get_saved_google_folder_url() -> str:
    return clean_text(load_config().get("google_folder_url"))


def get_saved_google_folder_id() -> str:
    return clean_text(load_config().get("google_folder_id"))


def get_saved_google_email_hint() -> str:
    return clean_text(load_config().get("google_email_hint"))


def get_env_google_client_id() -> str:
    return clean_text(os.environ.get(ENV_GOOGLE_CLIENT_ID))


def get_env_google_client_secret() -> str:
    return clean_text(os.environ.get(ENV_GOOGLE_CLIENT_SECRET))


def get_env_google_folder_url() -> str:
    return clean_text(os.environ.get(ENV_GOOGLE_FOLDER_URL))


def get_env_google_folder_id() -> str:
    return clean_text(os.environ.get(ENV_GOOGLE_FOLDER_ID))


def get_default_google_folder_url() -> str:
    env_value = get_env_google_folder_url()
    if env_value:
        return env_value
    return DEFAULT_GOOGLE_FOLDER_URL


def extract_google_drive_folder_id(folder_url_or_id: str) -> str:
    raw = clean_text(folder_url_or_id)
    if not raw:
        raise StorageError("Google Drive folder URL or folder ID is required.")

    if _DRIVE_FOLDER_ID_RE.fullmatch(raw):
        return raw

    parsed = urlparse(raw)
    if parsed.scheme.lower() not in {"http", "https"}:
        raise StorageError("Folder value must be a Drive URL or raw folder ID.")

    host = parsed.netloc.lower()
    if "drive.google.com" not in host:
        raise StorageError("Folder URL must be a drive.google.com folder URL.")

    path_parts = [part for part in parsed.path.split("/") if part]
    for index, part in enumerate(path_parts):
        if part == "folders" and index + 1 < len(path_parts):
            candidate = clean_text(path_parts[index + 1])
            if _DRIVE_FOLDER_ID_RE.fullmatch(candidate):
                return candidate

    query = parse_qs(parsed.query)
    for key in ("id", "folderid"):
        candidate = clean_text(query.get(key, [""])[0])
        if _DRIVE_FOLDER_ID_RE.fullmatch(candidate):
            return candidate

    raise StorageError(
        "Unable to extract Google Drive folder ID from URL. "
        "Use a URL like https://drive.google.com/drive/folders/<id>."
    )


def _escape_drive_query(value: str) -> str:
    return value.replace("\\", "\\\\").replace("'", "\\'")


def _extract_http_status(exc: Exception) -> int:
    response = getattr(exc, "resp", None)
    status = getattr(response, "status", 0)
    try:
        return int(status or 0)
    except (TypeError, ValueError):
        return 0


def _extract_http_error_message(exc: Exception) -> str:
    content = getattr(exc, "content", b"")
    if isinstance(content, bytes):
        text = clean_text(content.decode("utf-8", errors="replace"))
    else:
        text = clean_text(str(content))
    if text:
        return text
    return clean_text(str(exc))


class GoogleDriveStore:
    def __init__(
        self,
        google_client_id: str,
        google_client_secret: str,
        google_folder_url: str = "",
        google_folder_id: str = "",
        email_hint: str = "",
    ) -> None:
        if (
            Credentials is None
            or GoogleAuthRequest is None
            or InstalledAppFlow is None
            or build is None
            or HttpError is None
            or MediaIoBaseUpload is None
        ):
            raise StorageError(
                "Missing Google dependencies. Install with:\n"
                "pip install -r requirements.txt"
            )

        self.google_client_id = clean_text(google_client_id)
        self.google_client_secret = clean_text(google_client_secret)
        if not self.google_client_id:
            raise StorageError("Google OAuth Client ID is required.")
        if not self.google_client_secret:
            raise StorageError("Google OAuth Client Secret is required.")

        self._email_hint = clean_text(email_hint)
        self._folder_url = clean_text(google_folder_url)

        folder_id = clean_text(google_folder_id)
        if not folder_id and self._folder_url:
            folder_id = extract_google_drive_folder_id(self._folder_url)
        if not folder_id:
            raise StorageError("Google Drive folder URL or folder ID is required.")

        self._root_folder_id = folder_id
        self._credentials: Any = self._load_credentials()
        self._drive: Any = None
        self._logs_folder_id = ""
        self._output_folder_id = ""

    @property
    def root_folder_id(self) -> str:
        return self._root_folder_id

    @property
    def folder_url(self) -> str:
        if self._folder_url:
            return self._folder_url
        return f"https://drive.google.com/drive/folders/{self._root_folder_id}"

    @property
    def email_hint(self) -> str:
        return self._email_hint

    def _load_credentials(self) -> Any:
        path = _token_path()
        if not path.exists():
            return None

        try:
            creds = Credentials.from_authorized_user_file(str(path), GOOGLE_SCOPES)
        except Exception as exc:
            raise StorageError(f"Unable to read Google token cache: {path}\n{exc}") from exc

        if clean_text(getattr(creds, "client_id", "")) != self.google_client_id:
            return None
        return creds

    def _save_credentials(self) -> None:
        if self._credentials is None:
            return

        path = _token_path()
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(self._credentials.to_json(), encoding="utf-8")
        except OSError as exc:
            raise StorageError(f"Unable to write Google token cache: {path}\n{exc}") from exc

    def _client_config(self) -> dict[str, Any]:
        return {
            "installed": {
                "client_id": self.google_client_id,
                "client_secret": self.google_client_secret,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "redirect_uris": ["http://localhost"],
            }
        }

    def _acquire_credentials_interactive(self) -> None:
        try:
            flow = InstalledAppFlow.from_client_config(self._client_config(), GOOGLE_SCOPES)
            creds = flow.run_local_server(port=0, prompt="consent")
        except Exception as exc:
            raise StorageError(
                "Google interactive sign-in failed.\n"
                "Verify OAuth Client ID/Secret and Desktop App configuration."
            ) from exc

        self._credentials = creds
        self._save_credentials()

    def _ensure_credentials(self, interactive_if_needed: bool) -> None:
        if self._credentials and self._credentials.valid:
            return

        if self._credentials and self._credentials.expired and self._credentials.refresh_token:
            try:
                self._credentials.refresh(GoogleAuthRequest())
                self._save_credentials()
                return
            except Exception:
                self._credentials = None

        if interactive_if_needed:
            self._acquire_credentials_interactive()
            return

        raise StorageError("No valid Google token found. Sign-in is required.")

    def _ensure_drive(self, interactive_if_needed: bool) -> Any:
        self._ensure_credentials(interactive_if_needed=interactive_if_needed)
        if self._drive is None:
            self._drive = build("drive", "v3", credentials=self._credentials, cache_discovery=False)
        return self._drive

    def _execute(
        self,
        request_factory: Callable[[Any], Any],
        *,
        interactive_on_401: bool = True,
        not_found_ok: bool = False,
    ) -> Any:
        for attempt in range(2):
            drive = self._ensure_drive(interactive_if_needed=interactive_on_401 and attempt == 0)
            try:
                request = request_factory(drive)
                return request.execute()
            except Exception as exc:
                status = _extract_http_status(exc)
                if status == 404 and not_found_ok:
                    return None
                if status in {401, 403} and interactive_on_401 and attempt == 0:
                    self._credentials = None
                    self._drive = None
                    self._acquire_credentials_interactive()
                    continue
                if HttpError is not None and isinstance(exc, HttpError):
                    detail = _extract_http_error_message(exc)
                    raise StorageError(
                        f"Google Drive API error {status}.\n{detail}"
                    ) from exc
                raise StorageError(f"Google Drive request failed.\n{exc}") from exc

        raise StorageError("Google Drive request failed after retry.")

    def _list_files(self, query: str, fields: str) -> list[dict[str, Any]]:
        items: list[dict[str, Any]] = []
        page_token = ""
        while True:
            payload = self._execute(
                lambda drive: drive.files().list(
                    q=query,
                    spaces="drive",
                    fields=f"nextPageToken,files({fields})",
                    pageToken=page_token or None,
                ),
                interactive_on_401=True,
            )
            if not isinstance(payload, dict):
                return items
            for item in payload.get("files", []) or []:
                if isinstance(item, dict):
                    items.append(item)
            page_token = clean_text(payload.get("nextPageToken"))
            if not page_token:
                break
        return items

    def _find_child_folder_id(self, parent_id: str, name: str) -> str:
        escaped = _escape_drive_query(name)
        query = (
            f"name = '{escaped}' and '{parent_id}' in parents and trashed = false "
            f"and mimeType = '{GOOGLE_FOLDER_MIME}'"
        )
        items = self._list_files(query, fields="id,name")
        if not items:
            return ""
        return clean_text(items[0].get("id"))

    def _find_child_file(self, parent_id: str, name: str) -> dict[str, Any] | None:
        escaped = _escape_drive_query(name)
        query = (
            f"name = '{escaped}' and '{parent_id}' in parents and trashed = false "
            f"and mimeType != '{GOOGLE_FOLDER_MIME}'"
        )
        items = self._list_files(query, fields="id,name,webViewLink,mimeType")
        if not items:
            return None
        return items[0]

    def _create_folder(self, parent_id: str, name: str) -> str:
        payload = self._execute(
            lambda drive: drive.files().create(
                body={
                    "name": name,
                    "mimeType": GOOGLE_FOLDER_MIME,
                    "parents": [parent_id],
                },
                fields="id",
            ),
            interactive_on_401=True,
        )
        if not isinstance(payload, dict) or not clean_text(payload.get("id")):
            raise StorageError(f"Failed to create Google Drive folder '{name}'.")
        return clean_text(payload.get("id"))

    def _ensure_subfolders(self) -> None:
        if self._logs_folder_id and self._output_folder_id:
            return

        logs = self._find_child_folder_id(self._root_folder_id, LOGS_DIRNAME)
        if not logs:
            logs = self._create_folder(self._root_folder_id, LOGS_DIRNAME)

        output = self._find_child_folder_id(self._root_folder_id, OUTPUT_DIRNAME)
        if not output:
            output = self._create_folder(self._root_folder_id, OUTPUT_DIRNAME)

        self._logs_folder_id = logs
        self._output_folder_id = output

    def _download_file_bytes(self, file_id: str, missing_ok: bool) -> bytes | None:
        payload = self._execute(
            lambda drive: drive.files().get_media(fileId=file_id),
            interactive_on_401=True,
            not_found_ok=missing_ok,
        )
        if payload is None:
            return None
        if isinstance(payload, bytes):
            return payload
        if isinstance(payload, str):
            return payload.encode("utf-8")
        raise StorageError("Unexpected Google Drive download response.")

    def _upload_file_bytes(
        self,
        parent_id: str,
        filename: str,
        data: bytes,
        mimetype: str,
    ) -> dict[str, Any]:
        existing = self._find_child_file(parent_id, filename)
        existing_id = clean_text(existing.get("id")) if isinstance(existing, dict) else ""

        def request_factory(drive: Any) -> Any:
            media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mimetype, resumable=False)
            if existing_id:
                return drive.files().update(
                    fileId=existing_id,
                    media_body=media,
                    fields="id,name,webViewLink",
                )
            return drive.files().create(
                body={"name": filename, "parents": [parent_id]},
                media_body=media,
                fields="id,name,webViewLink",
            )

        payload = self._execute(request_factory, interactive_on_401=True)
        if not isinstance(payload, dict):
            raise StorageError(f"Upload failed for file '{filename}'.")
        if not clean_text(payload.get("id")):
            raise StorageError(f"Upload succeeded but file ID missing for '{filename}'.")
        return payload

    def validate_connection(self) -> None:
        payload = self._execute(
            lambda drive: drive.files().get(
                fileId=self._root_folder_id,
                fields="id,name,mimeType,webViewLink",
            ),
            interactive_on_401=True,
        )
        if not isinstance(payload, dict):
            raise StorageError("Unexpected Google Drive folder metadata response.")

        mime = clean_text(payload.get("mimeType"))
        if mime != GOOGLE_FOLDER_MIME:
            raise StorageError("Configured Google Drive ID is not a folder.")

        link = clean_text(payload.get("webViewLink"))
        if link:
            self._folder_url = link

    def ensure_structure(self) -> None:
        self._ensure_subfolders()

    def _file_web_link(self, file_payload: dict[str, Any], file_id: str) -> str:
        link = clean_text(file_payload.get("webViewLink"))
        if link:
            return link
        return f"https://drive.google.com/file/d/{file_id}/view"

    def ensure_user_log_exists(self, user_id: str) -> str:
        self._ensure_subfolders()
        safe_user = sanitize_userid_for_filename(user_id)
        filename = f"{safe_user}_events.csv"

        existing = self._find_child_file(self._logs_folder_id, filename)
        if existing is not None and clean_text(existing.get("id")):
            return self._file_web_link(existing, clean_text(existing.get("id")))

        output = io.StringIO()
        writer = csv.DictWriter(output, fieldnames=EVENT_COLUMNS)
        writer.writeheader()
        uploaded = self._upload_file_bytes(
            self._logs_folder_id,
            filename,
            output.getvalue().encode("utf-8"),
            mimetype="text/csv",
        )
        file_id = clean_text(uploaded.get("id"))
        return self._file_web_link(uploaded, file_id)

    def append_event(self, user_id: str, event: KeyEvent) -> str:
        self._ensure_subfolders()
        safe_user = sanitize_userid_for_filename(user_id)
        filename = f"{safe_user}_events.csv"

        existing = self._find_child_file(self._logs_folder_id, filename)
        existing_id = clean_text(existing.get("id")) if existing else ""

        existing_text = ""
        if existing_id:
            file_bytes = self._download_file_bytes(existing_id, missing_ok=True)
            if file_bytes:
                try:
                    existing_text = file_bytes.decode("utf-8-sig")
                except UnicodeDecodeError:
                    try:
                        existing_text = file_bytes.decode("utf-8")
                    except UnicodeDecodeError as exc:
                        raise StorageError(
                            f"Unable to decode event log for user '{user_id}' as UTF-8."
                        ) from exc

        output = io.StringIO()
        if clean_text(existing_text):
            output.write(existing_text)
            if not existing_text.endswith(("\n", "\r")):
                output.write("\n")
        else:
            writer = csv.DictWriter(output, fieldnames=EVENT_COLUMNS)
            writer.writeheader()

        writer = csv.DictWriter(output, fieldnames=EVENT_COLUMNS)
        writer.writerow(event.to_csv_row())

        uploaded = self._upload_file_bytes(
            self._logs_folder_id,
            filename,
            output.getvalue().encode("utf-8"),
            mimetype="text/csv",
        )
        file_id = clean_text(uploaded.get("id"))
        return self._file_web_link(uploaded, file_id)

    def read_all_events(self) -> tuple[list[KeyEvent], list[str]]:
        self._ensure_subfolders()
        all_events: list[KeyEvent] = []
        all_warnings: list[str] = []

        query = (
            f"'{self._logs_folder_id}' in parents and trashed = false "
            f"and mimeType != '{GOOGLE_FOLDER_MIME}'"
        )
        for item in self._list_files(query, fields="id,name,mimeType"):
            name = clean_text(item.get("name"))
            file_id = clean_text(item.get("id"))
            if not name.lower().endswith(".csv") or not file_id:
                continue

            file_bytes = self._download_file_bytes(file_id, missing_ok=True)
            if file_bytes is None:
                continue

            try:
                text = file_bytes.decode("utf-8-sig")
            except UnicodeDecodeError:
                try:
                    text = file_bytes.decode("utf-8")
                except UnicodeDecodeError as exc:
                    all_warnings.append(f"{name}: UTF-8 decode failed ({exc})")
                    continue

            events, warnings = _read_events_from_text(name, text)
            all_events.extend(events)
            all_warnings.extend(warnings)

        all_events.sort(key=_event_sort_key)
        return all_events, all_warnings

    def export_aggregate(self) -> tuple[str, list[str]]:
        try:
            from openpyxl import Workbook
        except ImportError as exc:
            raise StorageError(
                "openpyxl is required for export. Install it with: pip install openpyxl"
            ) from exc

        self._ensure_subfolders()
        events, warnings = self.read_all_events()
        whats_out = compute_whats_out(events)

        workbook = Workbook()
        raw_sheet = workbook.active
        raw_sheet.title = "RawEvents"
        raw_sheet.append(list(EVENT_COLUMNS))
        for event in events:
            row = event.to_csv_row()
            raw_sheet.append([row[column] for column in EVENT_COLUMNS])

        out_sheet = workbook.create_sheet("WhatsOut")
        headers = ["KeyId", "CheckedOutBy", "TimeOut", "ToLocation", "PrinterOrDestination"]
        out_sheet.append(headers)
        for item in whats_out:
            out_sheet.append(
                [
                    item.KeyId,
                    item.CheckedOutBy,
                    item.TimeOut,
                    item.ToLocation,
                    item.PrinterOrDestination,
                ]
            )

        _autosize_columns(raw_sheet)
        _autosize_columns(out_sheet)

        buffer = io.BytesIO()
        workbook.save(buffer)
        workbook.close()

        uploaded = self._upload_file_bytes(
            self._output_folder_id,
            AGGREGATE_FILENAME,
            buffer.getvalue(),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        file_id = clean_text(uploaded.get("id"))
        return self._file_web_link(uploaded, file_id), warnings


def create_google_drive_store(
    google_client_id: str,
    google_client_secret: str,
    google_folder_url: str = "",
    google_folder_id: str = "",
    email_hint: str = "",
) -> GoogleDriveStore:
    store = GoogleDriveStore(
        google_client_id=google_client_id,
        google_client_secret=google_client_secret,
        google_folder_url=google_folder_url,
        google_folder_id=google_folder_id,
        email_hint=email_hint,
    )
    store.validate_connection()
    store.ensure_structure()
    return store


def sanitize_userid_for_filename(user_id: str) -> str:
    cleaned = clean_text(user_id)
    if not cleaned:
        raise StorageError("UserId is required.")

    safe = "".join(ch if (ch.isalnum() or ch in "._-") else "_" for ch in cleaned)
    if not safe:
        raise StorageError("UserId must include at least one valid filename character.")
    return safe


def ensure_user_log_exists(store: GoogleDriveStore, user_id: str) -> str:
    return store.ensure_user_log_exists(user_id)


def append_event(store: GoogleDriveStore, user_id: str, event: KeyEvent) -> str:
    return store.append_event(user_id, event)


def make_checkout_event(
    user_id: str,
    key_id: str,
    from_location: str,
    to_location: str,
    printer_or_destination: str = "",
    notes: str = "",
) -> KeyEvent:
    return KeyEvent(
        EventId=new_event_id(),
        Timestamp=now_iso_timestamp(),
        UserId=clean_text(user_id),
        Action="OUT",
        KeyId=clean_text(key_id),
        FromLocation=clean_text(from_location),
        ToLocation=clean_text(to_location),
        PrinterOrDestination=clean_text(printer_or_destination),
        ReturnedToLocation="",
        Notes=clean_text(notes),
    )


def make_return_event(
    user_id: str,
    key_id: str,
    returned_to_location: str,
    notes: str = "",
) -> KeyEvent:
    return KeyEvent(
        EventId=new_event_id(),
        Timestamp=now_iso_timestamp(),
        UserId=clean_text(user_id),
        Action="IN",
        KeyId=clean_text(key_id),
        FromLocation="",
        ToLocation="",
        PrinterOrDestination="",
        ReturnedToLocation=clean_text(returned_to_location),
        Notes=clean_text(notes),
    )


def _event_sort_key(event: KeyEvent) -> datetime:
    parsed = parse_iso_timestamp(event.Timestamp)
    if parsed is None:
        return _FALLBACK_MIN_TS
    return parsed


def _parse_row(row: dict[str, str]) -> KeyEvent:
    normalized = {column: clean_text(row.get(column, "")) for column in EVENT_COLUMNS}
    missing = [name for name in _REQUIRED_COLUMNS if not normalized.get(name)]
    if missing:
        raise ValueError(f"Missing required values: {', '.join(sorted(missing))}")

    action = normalized["Action"].upper()
    if action not in _VALID_ACTIONS:
        raise ValueError(f"Invalid Action '{normalized['Action']}'")
    normalized["Action"] = action

    timestamp = normalized["Timestamp"]
    if parse_iso_timestamp(timestamp) is None:
        raise ValueError(f"Invalid Timestamp '{timestamp}'")

    key_id = normalized["KeyId"]
    if not key_id:
        raise ValueError("KeyId cannot be blank")

    return KeyEvent(**normalized)


def _read_events_from_text(filename: str, text: str) -> tuple[list[KeyEvent], list[str]]:
    events: list[KeyEvent] = []
    warnings: list[str] = []

    if not clean_text(text):
        return events, warnings

    reader = csv.DictReader(io.StringIO(text))
    fieldnames = set(reader.fieldnames or [])
    if not fieldnames:
        warnings.append(f"{filename}: missing header; file skipped.")
        return events, warnings

    missing_headers = sorted(_REQUIRED_COLUMNS - fieldnames)
    if missing_headers:
        warnings.append(f"{filename}: missing required columns {missing_headers}; file skipped.")
        return events, warnings

    for line_number, row in enumerate(reader, start=2):
        try:
            event = _parse_row(row)
        except Exception as exc:
            warnings.append(f"{filename}:{line_number}: {exc}")
            continue
        events.append(event)

    return events, warnings


def latest_event_per_key(events: list[KeyEvent]) -> dict[str, KeyEvent]:
    latest: dict[str, KeyEvent] = {}
    for event in sorted(events, key=_event_sort_key):
        normalized_key = normalize_key_id(event.KeyId)
        if not normalized_key:
            continue
        latest[normalized_key] = event
    return latest


def compute_whats_out(events: list[KeyEvent]) -> list[OutstandingKey]:
    latest = latest_event_per_key(events)
    rows: list[OutstandingKey] = []

    for event in latest.values():
        if event.Action != "OUT":
            continue
        rows.append(
            OutstandingKey(
                KeyId=event.KeyId,
                CheckedOutBy=event.UserId,
                TimeOut=event.Timestamp,
                ToLocation=event.ToLocation,
                PrinterOrDestination=event.PrinterOrDestination,
            )
        )

    rows.sort(key=lambda item: parse_iso_timestamp(item.TimeOut) or _FALLBACK_MAX_TS)
    return rows


def get_whats_out(store: GoogleDriveStore) -> tuple[list[OutstandingKey], list[str]]:
    events, warnings = store.read_all_events()
    return compute_whats_out(events), warnings


def get_current_checkout(store: GoogleDriveStore, key_id: str) -> tuple[KeyEvent | None, list[str]]:
    normalized = normalize_key_id(key_id)
    if not normalized:
        return None, []

    events, warnings = store.read_all_events()
    latest = latest_event_per_key(events)
    event = latest.get(normalized)
    if event is None or event.Action != "OUT":
        return None, warnings
    return event, warnings


def export_aggregate(store: GoogleDriveStore) -> tuple[str, list[str]]:
    return store.export_aggregate()


def _autosize_columns(sheet: Any) -> None:
    try:
        for column_cells in sheet.columns:
            values = ["" if cell.value is None else str(cell.value) for cell in column_cells]
            max_length = max((len(value) for value in values), default=0)
            width = min(max_length + 2, 60)
            if width < 10:
                width = 10
            sheet.column_dimensions[column_cells[0].column_letter].width = width
    except Exception:
        pass
