#!/usr/bin/env python3
"""Printer Supply & Tray Monitor.

Features:
- Pulls printer status from a WEPA monitoring page.
- Detects low toner/ink/fuser conditions.
- Detects empty trays and tracks empty/filled events over time.
- Exports tray events to CSV.
- Generates copyable ServiceNow work notes for selected empty trays.
"""

from __future__ import annotations

import csv
import json
import os
import re
import sys
import threading
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from datetime import datetime
from html import unescape
from pathlib import Path
from typing import Any

import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

APP_NAME = "Printer Supply & Tray Monitor"
APP_SLUG = "printer_supply_tray_monitor"
APP_CREDITS = "Credits: Jack Shetterly"
APP_ICON_FILE = "PrinterSupplyTrayMonitor-256.png"
DEFAULT_URL = ""
STATE_FILENAME = "printer_monitor_state.json"
DEFAULT_TONER_THRESHOLD = 15
DEFAULT_FUSER_THRESHOLD = 20
DEFAULT_INTERVAL_MINUTES = 5
MAX_EVENTS = 20000

COLOR_MAP = {
    "k": "Black",
    "c": "Cyan",
    "m": "Magenta",
    "y": "Yellow",
}


def get_app_data_dir() -> Path:
    candidates: list[Path] = []

    if sys.platform == "darwin":
        candidates.append(Path.home() / "Library" / "Application Support")
    elif sys.platform.startswith("win"):
        appdata = os.getenv("APPDATA")
        if appdata:
            candidates.append(Path(appdata))
        else:
            candidates.append(Path.home() / "AppData" / "Roaming")
    else:
        candidates.append(Path(os.getenv("XDG_DATA_HOME", str(Path.home() / ".local" / "share"))))

    # Always include a local fallback for restricted environments.
    candidates.append(Path(__file__).resolve().parent)

    for base_dir in candidates:
        app_dir = base_dir / APP_SLUG
        try:
            app_dir.mkdir(parents=True, exist_ok=True)
            return app_dir
        except OSError:
            continue

    raise OSError("Unable to create an application data directory.")


def now_iso() -> str:
    return datetime.now().astimezone().replace(microsecond=0).isoformat()


def parse_iso(value: str) -> datetime | None:
    try:
        return datetime.fromisoformat(value)
    except (TypeError, ValueError):
        return None


def display_time(value: str) -> str:
    dt = parse_iso(value)
    if dt is None:
        return value or "-"
    return dt.astimezone().strftime("%Y-%m-%d %H:%M:%S")


@dataclass
class PrinterRecord:
    printer_id: str
    description: str = ""
    status_message: str = ""
    printer_text: str = ""
    device_last_update: str = ""
    row_tail: str = ""
    levels: dict[str, int] = field(default_factory=dict)
    empty_trays: list[str] = field(default_factory=list)


@dataclass
class LowAlert:
    printer_id: str
    description: str
    item: str
    level: str
    source: str


class StateStore:
    def __init__(self, state_path: Path) -> None:
        self.state_path = state_path
        self.data: dict[str, Any] = {
            "open_empty_trays": {},
            "events": [],
            "last_scan_at": "",
        }
        self.load()

    def load(self) -> None:
        if not self.state_path.exists():
            return
        try:
            self.data = json.loads(self.state_path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            self.data = {
                "open_empty_trays": {},
                "events": [],
                "last_scan_at": "",
            }

    def save(self) -> None:
        tmp_path = self.state_path.with_suffix(".tmp")
        tmp_path.write_text(json.dumps(self.data, indent=2), encoding="utf-8")
        tmp_path.replace(self.state_path)

    def get_open_empties(self) -> dict[str, dict[str, str]]:
        return dict(self.data.get("open_empty_trays", {}))

    def get_events(self) -> list[dict[str, str]]:
        return list(self.data.get("events", []))

    def get_last_scan(self) -> str:
        return str(self.data.get("last_scan_at", ""))

    def reconcile(self, current_empties: dict[str, dict[str, str]], scan_time: str) -> dict[str, int]:
        previous: dict[str, dict[str, str]] = self.data.get("open_empty_trays", {})
        open_after: dict[str, dict[str, str]] = {}
        added = 0
        filled = 0

        for key, item in current_empties.items():
            prior = previous.get(key)
            if prior is None:
                item["since"] = scan_time
                item["last_seen"] = scan_time
                open_after[key] = item
                self._append_event(scan_time, "empty", item, "detected by monitor")
                added += 1
            else:
                item["since"] = prior.get("since", scan_time)
                item["last_seen"] = scan_time
                open_after[key] = item

        for key, prior in previous.items():
            if key not in current_empties:
                resolved_item = dict(prior)
                resolved_item["last_seen"] = scan_time
                self._append_event(scan_time, "filled", resolved_item, "no longer reported empty")
                filled += 1

        self.data["open_empty_trays"] = open_after
        self.data["last_scan_at"] = scan_time
        self._trim_events()
        self.save()
        return {"new_empties": added, "new_filled": filled}

    def manual_mark_filled(self, tray_key: str, timestamp: str) -> bool:
        current = self.data.get("open_empty_trays", {})
        item = current.pop(tray_key, None)
        if item is None:
            return False

        item = dict(item)
        item["last_seen"] = timestamp
        self._append_event(timestamp, "filled", item, "manually marked filled")
        self.data["open_empty_trays"] = current
        self.data["last_scan_at"] = timestamp
        self._trim_events()
        self.save()
        return True

    def _append_event(self, timestamp: str, event_type: str, item: dict[str, str], note: str) -> None:
        event = {
            "timestamp": timestamp,
            "event_type": event_type,
            "printer_id": item.get("printer_id", ""),
            "description": item.get("description", ""),
            "tray": item.get("tray", ""),
            "empty_since": item.get("since", ""),
            "last_seen": item.get("last_seen", ""),
            "status_message": item.get("status_message", ""),
            "printer_text": item.get("printer_text", ""),
            "note": note,
        }
        self.data.setdefault("events", []).append(event)

    def _trim_events(self) -> None:
        events = self.data.setdefault("events", [])
        if len(events) > MAX_EVENTS:
            self.data["events"] = events[-MAX_EVENTS:]


def fetch_html(url: str, timeout_seconds: int = 25) -> str:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
        },
    )
    with urllib.request.urlopen(request, timeout=timeout_seconds) as response:
        raw = response.read()
        charset = response.headers.get_content_charset() or "utf-8"
    return raw.decode(charset, errors="replace")


def html_to_lines(html: str) -> list[str]:
    cleaned = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", html)
    cleaned = re.sub(r"(?i)<br\s*/?>", "\n", cleaned)
    cleaned = re.sub(r"(?i)</(p|div|tr|td|th|li|h1|h2|h3|h4|h5|h6)>", "\n", cleaned)
    cleaned = re.sub(r"(?i)<[^>]+>", " ", cleaned)
    cleaned = unescape(cleaned)

    lines: list[str] = []
    for line in cleaned.splitlines():
        compact = re.sub(r"\s+", " ", line).strip()
        if compact:
            lines.append(compact)
    return lines


def parse_tail_metrics(row_tail: str) -> tuple[dict[str, int], str]:
    levels: dict[str, int] = {}
    last_update = ""

    time_match = re.search(r"\b(\d{2}/\d{2}/\d{2}\s+\d{2}:\d{2}:\d{2})\b", row_tail)
    if time_match:
        last_update = time_match.group(1)

    numbers = [int(value) for value in re.findall(r"\b(\d{1,3})\b", row_tail)]
    if len(numbers) >= 10:
        candidate = numbers[-10:]
        if all(0 <= value <= 100 for value in candidate):
            keys = [
                "toner_k",
                "toner_c",
                "toner_m",
                "toner_y",
                "drum_k",
                "drum_c",
                "drum_m",
                "drum_y",
                "belt",
                "fuser",
            ]
            levels = dict(zip(keys, candidate))

    return levels, last_update


def normalize_tray(value: str) -> str:
    token = re.sub(r"[^A-Za-z0-9-]", "", value).upper()
    if not token:
        return "Unknown Tray"
    if token.isdigit():
        return f"Tray {int(token)}"
    if token.startswith("TRAY"):
        token = token[4:] or "UNKNOWN"
        if token.isdigit():
            return f"Tray {int(token)}"
    return f"Tray {token}"


def detect_empty_trays(text: str) -> list[str]:
    trays: set[str] = set()

    patterns = [
        r"\btray\s*([A-Za-z0-9-]+)\s*(?:is\s*)?(?:empty|out|no\s+paper)\b",
        r"\bpaper\s*out\s*(?:in\s*)?tray\s*([A-Za-z0-9-]+)\b",
        r"\b(?:empty|out)\s*tray\s*([A-Za-z0-9-]+)\b",
    ]

    for pattern in patterns:
        for match in re.finditer(pattern, text, flags=re.IGNORECASE):
            trays.add(normalize_tray(match.group(1)))

    if not trays and re.search(r"\b(paper out|tray empty)\b", text, flags=re.IGNORECASE):
        trays.add("Unknown Tray")

    return sorted(trays)


def parse_monitor_page(html: str) -> list[PrinterRecord]:
    lines = html_to_lines(html)
    records: list[PrinterRecord] = []
    current: PrinterRecord | None = None

    id_pattern = re.compile(r"^(\d{5})\b(.*)$")

    for line in lines:
        id_match = id_pattern.match(line)
        if id_match:
            if current is not None:
                finalize_record(current)
                records.append(current)

            printer_id = id_match.group(1)
            tail = id_match.group(2).strip()
            levels, last_update = parse_tail_metrics(tail)
            current = PrinterRecord(
                printer_id=printer_id,
                row_tail=tail,
                levels=levels,
                device_last_update=last_update,
            )
            continue

        if current is None:
            continue

        lowered = line.lower()
        if lowered.startswith("description:"):
            current.description = line.split(":", 1)[1].strip()
        elif lowered.startswith("status message:"):
            current.status_message = line.split(":", 1)[1].strip()
        elif lowered.startswith("printer text:"):
            current.printer_text = line.split(":", 1)[1].strip()
        elif lowered.startswith("fuser:"):
            fuser_match = re.search(r"fuser:\s*(\d{1,3})%", line, flags=re.IGNORECASE)
            belt_match = re.search(r"belt:\s*(\d{1,3})%", line, flags=re.IGNORECASE)
            if fuser_match:
                current.levels["fuser"] = int(fuser_match.group(1))
            if belt_match:
                current.levels["belt"] = int(belt_match.group(1))

    if current is not None:
        finalize_record(current)
        records.append(current)

    # Guard against false positives from page chrome text.
    records = [record for record in records if record.printer_id.isdigit()]
    return records


def finalize_record(record: PrinterRecord) -> None:
    if not record.description:
        description_guess = re.split(
            r"\b\d{2}/\d{2}/\d{2}\s+\d{2}:\d{2}:\d{2}\b",
            record.row_tail,
            maxsplit=1,
        )[0].strip()
        record.description = description_guess or f"Printer {record.printer_id}"

    combined_text = " ".join(
        value
        for value in [record.status_message, record.printer_text, record.row_tail]
        if value and value.lower() not in {"none", "n/a"}
    )
    record.empty_trays = detect_empty_trays(combined_text)


def build_low_alerts(records: list[PrinterRecord], toner_threshold: int, fuser_threshold: int) -> list[LowAlert]:
    alerts: list[LowAlert] = []
    seen: set[tuple[str, str, str]] = set()

    for record in records:
        cmy_levels = [record.levels.get("toner_c"), record.levels.get("toner_m"), record.levels.get("toner_y")]
        mono_like = all(level == 0 for level in cmy_levels if level is not None) and all(
            level is not None for level in cmy_levels
        )

        for channel in ["k", "c", "m", "y"]:
            if mono_like and channel in {"c", "m", "y"}:
                continue
            key = f"toner_{channel}"
            level = record.levels.get(key)
            if level is not None and level <= toner_threshold:
                item = f"{COLOR_MAP[channel]} Toner"
                dedupe = (record.printer_id, item, "level")
                if dedupe not in seen:
                    alerts.append(
                        LowAlert(
                            printer_id=record.printer_id,
                            description=record.description,
                            item=item,
                            level=f"{level}%",
                            source=f"threshold <= {toner_threshold}%",
                        )
                    )
                    seen.add(dedupe)

        fuser = record.levels.get("fuser")
        if fuser is not None and fuser <= fuser_threshold:
            dedupe = (record.printer_id, "Fuser", "level")
            if dedupe not in seen:
                alerts.append(
                    LowAlert(
                        printer_id=record.printer_id,
                        description=record.description,
                        item="Fuser",
                        level=f"{fuser}%",
                        source=f"threshold <= {fuser_threshold}%",
                    )
                )
                seen.add(dedupe)

        status_blob = f"{record.status_message} {record.printer_text}".lower()
        keyword_rules = [
            (r"\blow\s+toner\b", "Toner", "status message reports low toner"),
            (r"\blow\s+ink\b", "Ink", "status message reports low ink"),
            (r"\blow\s+fuser\b", "Fuser", "status message reports low fuser"),
        ]
        for pattern, item, reason in keyword_rules:
            if re.search(pattern, status_blob):
                dedupe = (record.printer_id, item, reason)
                if dedupe not in seen:
                    alerts.append(
                        LowAlert(
                            printer_id=record.printer_id,
                            description=record.description,
                            item=item,
                            level="reported",
                            source=reason,
                        )
                    )
                    seen.add(dedupe)

    return sorted(alerts, key=lambda alert: (alert.printer_id, alert.item))


def build_current_empties(records: list[PrinterRecord]) -> dict[str, dict[str, str]]:
    empties: dict[str, dict[str, str]] = {}

    for record in records:
        for tray in record.empty_trays:
            key = f"{record.printer_id}::{tray}"
            empties[key] = {
                "key": key,
                "printer_id": record.printer_id,
                "description": record.description,
                "tray": tray,
                "status_message": record.status_message,
                "printer_text": record.printer_text,
                "device_last_update": record.device_last_update,
            }

    return empties


class PrinterMonitorApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_NAME)
        self._window_icon = None
        self._apply_window_icon()
        self.root.geometry("1500x920")

        self.app_data_dir = get_app_data_dir()
        self.state_store = StateStore(self.app_data_dir / STATE_FILENAME)
        self.state_lock = threading.Lock()

        self.url_var = tk.StringVar(value=DEFAULT_URL)
        self.toner_threshold_var = tk.IntVar(value=DEFAULT_TONER_THRESHOLD)
        self.fuser_threshold_var = tk.IntVar(value=DEFAULT_FUSER_THRESHOLD)
        self.auto_refresh_var = tk.BooleanVar(value=False)
        self.refresh_interval_var = tk.IntVar(value=DEFAULT_INTERVAL_MINUTES)
        self.worknote_mode_var = tk.StringVar(value="Detected Empty")

        self.status_var = tk.StringVar(value="Ready")
        self.last_refresh_var = tk.StringVar(value="Never")
        self.summary_printer_var = tk.StringVar(value="0")
        self.summary_low_var = tk.StringVar(value="0")
        self.summary_empty_var = tk.StringVar(value="0")

        self.records: list[PrinterRecord] = []
        self.low_alerts: list[LowAlert] = []
        self.open_empties: dict[str, dict[str, str]] = self.state_store.get_open_empties()
        self.events: list[dict[str, str]] = self.state_store.get_events()

        self.refresh_in_progress = False
        self.next_auto_refresh_epoch = 0.0

        self._build_styles()
        self._build_ui()
        self._refresh_history_tree()
        self._refresh_empty_tree()

        saved_scan = self.state_store.get_last_scan()
        if saved_scan:
            self.last_refresh_var.set(display_time(saved_scan))

        self.trigger_refresh()
        self.root.after(1000, self._auto_refresh_tick)

    def _build_styles(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        families = set(tkfont.families())
        if "SF Pro Text" in families:
            self.font_family = "SF Pro Text"
        elif "Helvetica Neue" in families:
            self.font_family = "Helvetica Neue"
        else:
            self.font_family = "Helvetica"

        if "SF Pro Display" in families:
            self.display_font_family = "SF Pro Display"
        else:
            self.display_font_family = self.font_family

        self.colors = {
            "window_bg": "#151a16",
            "chrome_bg": "#1d231e",
            "surface": "#242b25",
            "surface_soft": "#212721",
            "surface_raised": "#2a322b",
            "border": "#3a433b",
            "border_soft": "#303830",
            "title": "#f3f5f1",
            "text": "#d2d8d0",
            "muted": "#98a396",
            "accent": "#0a84ff",
            "accent_active": "#2f96ff",
            "accent_pressed": "#0070e8",
            "accent_soft": "#233140",
            "table_header": "#2d352e",
            "table_alt": "#283027",
            "selection": "#32465c",
            "footer_bg": "#1b201b",
            "footer_text": "#879182",
        }

        self.root.configure(bg=self.colors["window_bg"])
        style.configure(".", background=self.colors["window_bg"], foreground=self.colors["text"], font=(self.font_family, 11))

        style.configure("App.TFrame", background=self.colors["window_bg"])
        style.configure("Chrome.TFrame", background=self.colors["chrome_bg"], relief="flat")
        style.configure("Toolbar.TFrame", background=self.colors["chrome_bg"], relief="flat")
        style.configure("SummaryCard.TFrame", background=self.colors["surface"], relief="flat")
        style.configure("Card.TFrame", background=self.colors["surface"], relief="flat")
        style.configure("Surface.TFrame", background=self.colors["surface"], relief="flat")

        style.configure(
            "Title.TLabel",
            background=self.colors["chrome_bg"],
            foreground=self.colors["title"],
            font=(self.display_font_family, 20, "bold"),
        )
        style.configure(
            "Subtitle.TLabel",
            background=self.colors["chrome_bg"],
            foreground=self.colors["muted"],
            font=(self.font_family, 11),
        )
        style.configure(
            "Toolbar.TLabel",
            background=self.colors["chrome_bg"],
            foreground=self.colors["muted"],
            font=(self.font_family, 11, "bold"),
        )
        style.configure(
            "Value.TLabel",
            background=self.colors["chrome_bg"],
            foreground=self.colors["text"],
            font=(self.font_family, 11, "bold"),
        )
        style.configure(
            "SummaryTitle.TLabel",
            background=self.colors["surface"],
            foreground=self.colors["muted"],
            font=(self.font_family, 11, "bold"),
        )
        style.configure(
            "SummaryValue.TLabel",
            background=self.colors["surface"],
            foreground=self.colors["title"],
            font=(self.display_font_family, 31, "bold"),
        )
        style.configure(
            "SummaryHint.TLabel",
            background=self.colors["surface"],
            foreground=self.colors["muted"],
            font=(self.font_family, 10),
        )

        style.configure(
            "Section.TLabelframe",
            background=self.colors["surface"],
            bordercolor=self.colors["border"],
            relief="flat",
            padding=12,
        )
        style.configure(
            "Section.TLabelframe.Label",
            background=self.colors["surface"],
            foreground=self.colors["title"],
            font=(self.font_family, 12, "bold"),
        )

        style.configure(
            "Field.TEntry",
            fieldbackground=self.colors["surface_raised"],
            foreground=self.colors["text"],
            borderwidth=1,
            relief="flat",
            padding=10,
        )
        style.map(
            "Field.TEntry",
            bordercolor=[("focus", self.colors["accent"])],
            lightcolor=[("focus", self.colors["accent"])],
            darkcolor=[("focus", self.colors["accent"])],
        )
        style.configure(
            "Field.TCombobox",
            fieldbackground=self.colors["surface_raised"],
            background=self.colors["surface_raised"],
            foreground=self.colors["text"],
            borderwidth=1,
            relief="flat",
            padding=7,
            arrowsize=14,
        )
        style.map(
            "Field.TCombobox",
            fieldbackground=[("readonly", self.colors["surface_raised"])],
            background=[("readonly", self.colors["surface_raised"])],
            bordercolor=[("focus", self.colors["accent"])],
            lightcolor=[("focus", self.colors["accent"])],
            darkcolor=[("focus", self.colors["accent"])],
        )
        style.configure(
            "Field.TSpinbox",
            fieldbackground=self.colors["surface_raised"],
            foreground=self.colors["text"],
            borderwidth=1,
            relief="flat",
            padding=6,
            arrowsize=12,
        )
        style.map(
            "Field.TSpinbox",
            bordercolor=[("focus", self.colors["accent"])],
            lightcolor=[("focus", self.colors["accent"])],
            darkcolor=[("focus", self.colors["accent"])],
        )

        style.configure(
            "PrimaryDark.TButton",
            font=(self.font_family, 11, "bold"),
            padding=(16, 9),
            foreground="#ffffff",
            background=self.colors["accent"],
            borderwidth=0,
            relief="flat",
        )
        style.map(
            "PrimaryDark.TButton",
            background=[("pressed", self.colors["accent_pressed"]), ("active", self.colors["accent_active"])],
            foreground=[("disabled", "#e4ebf5")],
        )
        style.configure(
            "GhostDark.TButton",
            font=(self.font_family, 11, "bold"),
            padding=(14, 8),
            foreground=self.colors["text"],
            background=self.colors["surface_raised"],
            borderwidth=0,
            relief="flat",
        )
        style.map(
            "GhostDark.TButton",
            background=[("pressed", "#343f35"), ("active", "#313a33")],
            foreground=[("disabled", "#6c766c")],
        )
        style.configure(
            "Toolbar.TCheckbutton",
            background=self.colors["chrome_bg"],
            foreground=self.colors["text"],
            font=(self.font_family, 11),
        )

        style.configure(
            "Glass.TNotebook",
            background=self.colors["window_bg"],
            borderwidth=0,
            tabmargins=(0, 10, 0, 0),
        )
        style.configure(
            "Glass.TNotebook.Tab",
            font=(self.font_family, 11, "bold"),
            padding=(20, 10),
            background=self.colors["surface_soft"],
            foreground=self.colors["muted"],
            borderwidth=0,
        )
        style.map(
            "Glass.TNotebook.Tab",
            background=[("selected", self.colors["surface_raised"]), ("active", "#2d352f")],
            foreground=[("selected", self.colors["title"]), ("active", self.colors["title"])],
        )

        style.configure(
            "Data.Treeview",
            background=self.colors["surface"],
            fieldbackground=self.colors["surface"],
            foreground=self.colors["text"],
            borderwidth=0,
            relief="flat",
            rowheight=31,
            font=(self.font_family, 11),
        )
        style.configure(
            "Data.Treeview.Heading",
            background=self.colors["table_header"],
            foreground=self.colors["title"],
            borderwidth=0,
            relief="flat",
            padding=(10, 9),
            font=(self.font_family, 11, "bold"),
        )
        style.map(
            "Data.Treeview",
            background=[("selected", self.colors["selection"])],
            foreground=[("selected", self.colors["title"])],
        )
        style.map(
            "Data.Treeview.Heading",
            background=[("active", "#334038")],
        )

        style.configure(
            "TScrollbar",
            troughcolor=self.colors["surface_soft"],
            background="#475247",
            bordercolor=self.colors["surface_soft"],
            arrowcolor="#9ba79b",
            relief="flat",
        )

        style.configure(
            "Status.TLabel",
            background=self.colors["footer_bg"],
            foreground=self.colors["footer_text"],
            font=(self.font_family, 10),
            padding=(12, 9),
        )
        style.configure(
            "Credit.TLabel",
            background=self.colors["footer_bg"],
            foreground=self.colors["muted"],
            font=(self.font_family, 10, "bold"),
            padding=(12, 9),
        )

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, style="App.TFrame", padding=(24, 20, 24, 16))
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(3, weight=1)

        hero = ttk.Frame(container, style="Chrome.TFrame", padding=(18, 12))
        hero.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        hero.columnconfigure(0, weight=1)
        ttk.Label(hero, text=APP_NAME, style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(hero, text="Operational view with live alerts and tray history", style="Subtitle.TLabel").grid(
            row=1, column=0, sticky="w", pady=(2, 0)
        )

        controls = ttk.Frame(container, style="Toolbar.TFrame", padding=(18, 14))
        controls.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        controls.columnconfigure(1, weight=1)
        controls.columnconfigure(7, weight=1)

        ttk.Label(controls, text="Monitor URL", style="Toolbar.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(controls, textvariable=self.url_var, style="Field.TEntry").grid(
            row=0, column=1, columnspan=3, sticky="ew", padx=(8, 10)
        )
        ttk.Button(controls, text="Refresh Now", command=self.trigger_refresh, style="PrimaryDark.TButton").grid(
            row=0, column=4, padx=(0, 10), sticky="e"
        )
        ttk.Checkbutton(
            controls,
            text="Auto Refresh",
            variable=self.auto_refresh_var,
            style="Toolbar.TCheckbutton",
        ).grid(row=0, column=5, sticky="w")
        ttk.Label(controls, text="Every (min)", style="Toolbar.TLabel").grid(row=0, column=6, sticky="e", padx=(10, 4))
        ttk.Spinbox(
            controls,
            from_=1,
            to=180,
            textvariable=self.refresh_interval_var,
            width=5,
            style="Field.TSpinbox",
        ).grid(row=0, column=7, sticky="w")

        ttk.Label(controls, text="Toner Alert <= %", style="Toolbar.TLabel").grid(row=1, column=0, sticky="w", pady=(12, 0))
        ttk.Spinbox(
            controls,
            from_=1,
            to=100,
            textvariable=self.toner_threshold_var,
            width=6,
            style="Field.TSpinbox",
        ).grid(row=1, column=1, sticky="w", pady=(12, 0))

        ttk.Label(controls, text="Fuser Alert <= %", style="Toolbar.TLabel").grid(row=1, column=2, sticky="w", pady=(12, 0))
        ttk.Spinbox(
            controls,
            from_=1,
            to=100,
            textvariable=self.fuser_threshold_var,
            width=6,
            style="Field.TSpinbox",
        ).grid(row=1, column=3, sticky="w", pady=(12, 0))

        ttk.Label(controls, text="Last Refresh", style="Toolbar.TLabel").grid(row=1, column=4, sticky="e", pady=(12, 0), padx=(0, 6))
        ttk.Label(controls, textvariable=self.last_refresh_var, style="Value.TLabel").grid(
            row=1, column=5, columnspan=3, sticky="w", pady=(12, 0)
        )

        summary = ttk.Frame(container, style="App.TFrame")
        summary.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        summary.columnconfigure((0, 1, 2), weight=1)

        self._build_summary_card(summary, 0, "Printers Seen", self.summary_printer_var)
        self._build_summary_card(summary, 1, "Low Supply Alerts", self.summary_low_var)
        self._build_summary_card(summary, 2, "Open Empty Trays", self.summary_empty_var)

        notebook = ttk.Notebook(container, style="Glass.TNotebook")
        notebook.grid(row=3, column=0, sticky="nsew")

        dashboard = ttk.Frame(notebook, style="App.TFrame", padding=(0, 10, 0, 0))
        history_tab = ttk.Frame(notebook, style="App.TFrame", padding=(0, 10, 0, 0))
        notebook.add(dashboard, text="Dashboard")
        notebook.add(history_tab, text="Tray History")

        dashboard.columnconfigure(0, weight=1)
        dashboard.columnconfigure(1, weight=1)
        dashboard.rowconfigure(0, weight=1)
        dashboard.rowconfigure(1, weight=1)

        low_frame = ttk.LabelFrame(dashboard, text="Low Supplies", style="Section.TLabelframe")
        low_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=(0, 6))
        low_frame.columnconfigure(0, weight=1)
        low_frame.rowconfigure(0, weight=1)

        low_columns = ("printer_id", "description", "item", "level", "source")
        self.low_tree = ttk.Treeview(low_frame, columns=low_columns, show="headings")
        for col, title, width in [
            ("printer_id", "Printer", 90),
            ("description", "Description", 260),
            ("item", "Item", 120),
            ("level", "Level", 90),
            ("source", "Source", 220),
        ]:
            self.low_tree.heading(col, text=title)
            self.low_tree.column(col, width=width, anchor="w")
        self.low_tree.grid(row=0, column=0, sticky="nsew")
        self._add_scrollbar(low_frame, self.low_tree, row=0, column=1)

        empty_frame = ttk.LabelFrame(dashboard, text="Current Empty Trays", style="Section.TLabelframe")
        empty_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0), pady=(0, 6))
        empty_frame.columnconfigure(0, weight=1)
        empty_frame.rowconfigure(0, weight=1)

        empty_columns = ("printer_id", "description", "tray", "since", "last_seen")
        self.empty_tree = ttk.Treeview(empty_frame, columns=empty_columns, show="headings", selectmode="browse")
        for col, title, width in [
            ("printer_id", "Printer", 90),
            ("description", "Description", 240),
            ("tray", "Tray", 110),
            ("since", "Empty Since", 160),
            ("last_seen", "Last Seen", 160),
        ]:
            self.empty_tree.heading(col, text=title)
            self.empty_tree.column(col, width=width, anchor="w")
        self.empty_tree.grid(row=0, column=0, sticky="nsew")
        self.empty_tree.bind("<<TreeviewSelect>>", lambda _event: self.generate_worknote())
        self._add_scrollbar(empty_frame, self.empty_tree, row=0, column=1)

        actions = ttk.Frame(empty_frame, style="Surface.TFrame")
        actions.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions, text="Generate Worknote", command=self.generate_worknote, style="GhostDark.TButton").pack(side="left")
        ttk.Button(actions, text="Copy Worknote", command=self.copy_worknote, style="GhostDark.TButton").pack(
            side="left", padx=(8, 0)
        )
        ttk.Button(actions, text="Mark Filled (Manual)", command=self.mark_selected_filled, style="GhostDark.TButton").pack(
            side="left", padx=(8, 0)
        )

        worknote_frame = ttk.LabelFrame(dashboard, text="ServiceNow Worknote Generator", style="Section.TLabelframe")
        worknote_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(6, 0))
        worknote_frame.columnconfigure(0, weight=1)
        worknote_frame.rowconfigure(1, weight=1)

        header = ttk.Frame(worknote_frame, style="Surface.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(header, text="Template", style="Toolbar.TLabel").pack(side="left")
        mode = ttk.Combobox(
            header,
            textvariable=self.worknote_mode_var,
            values=["Detected Empty", "Refilled and Tested"],
            width=24,
            state="readonly",
            style="Field.TCombobox",
        )
        mode.pack(side="left", padx=(6, 8))
        mode.bind("<<ComboboxSelected>>", lambda _event: self.generate_worknote())

        self.worknote_text = tk.Text(
            worknote_frame,
            wrap="word",
            height=10,
            font=(self.font_family, 12),
            bg=self.colors["surface_raised"],
            fg=self.colors["text"],
            insertbackground=self.colors["title"],
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["accent"],
            padx=12,
            pady=10,
        )
        self.worknote_text.grid(row=1, column=0, sticky="nsew")

        history_tab.columnconfigure(0, weight=1)
        history_tab.rowconfigure(1, weight=1)

        history_actions = ttk.Frame(history_tab, style="App.TFrame")
        history_actions.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(history_actions, text="Export History CSV", command=self.export_history_csv, style="GhostDark.TButton").pack(
            side="left"
        )

        history_columns = ("timestamp", "event_type", "printer_id", "description", "tray", "note")
        self.history_tree = ttk.Treeview(history_tab, columns=history_columns, show="headings", style="Data.Treeview")
        for col, title, width in [
            ("timestamp", "Time", 170),
            ("event_type", "Event", 90),
            ("printer_id", "Printer", 90),
            ("description", "Description", 320),
            ("tray", "Tray", 110),
            ("note", "Note", 260),
        ]:
            self.history_tree.heading(col, text=title)
            self.history_tree.column(col, width=width, anchor="w")
        self.history_tree.grid(row=1, column=0, sticky="nsew")
        self._add_scrollbar(history_tab, self.history_tree, row=1, column=1)

        for tree in (self.low_tree, self.empty_tree, self.history_tree):
            tree.configure(style="Data.Treeview")
            tree.tag_configure("evenrow", background=self.colors["surface"])
            tree.tag_configure("oddrow", background=self.colors["table_alt"])

        footer = ttk.Frame(container, style="Chrome.TFrame")
        footer.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        footer.columnconfigure(0, weight=1)

        ttk.Label(footer, textvariable=self.status_var, style="Status.TLabel", anchor="w").grid(
            row=0, column=0, sticky="ew"
        )
        ttk.Label(footer, text=APP_CREDITS, style="Credit.TLabel", anchor="e").grid(row=0, column=1, sticky="e")

    def _build_summary_card(self, parent: ttk.Frame, column: int, title: str, variable: tk.StringVar) -> None:
        card = ttk.Frame(parent, style="SummaryCard.TFrame", padding=(20, 16))
        card.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 8, 0))
        ttk.Label(card, text=title, style="SummaryTitle.TLabel").pack(anchor="w")
        ttk.Label(card, textvariable=variable, style="SummaryValue.TLabel").pack(anchor="w", pady=(6, 0))
        ttk.Label(card, text="Live", style="SummaryHint.TLabel").pack(anchor="w", pady=(2, 0))

    def _add_scrollbar(self, parent: ttk.Widget, tree: ttk.Treeview, row: int, column: int) -> None:
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=row, column=column, sticky="ns")

    def _apply_window_icon(self) -> None:
        search_dirs: list[Path] = []
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            search_dirs.append(Path(meipass))
        search_dirs.append(Path(__file__).resolve().parent)

        for base in search_dirs:
            icon_path = base / "assets" / "icons" / APP_ICON_FILE
            if not icon_path.exists():
                continue
            try:
                self._window_icon = tk.PhotoImage(file=str(icon_path))
                self.root.iconphoto(True, self._window_icon)
                return
            except tk.TclError:
                continue

    def trigger_refresh(self) -> None:
        if self.refresh_in_progress:
            return

        url = self.url_var.get().strip()
        if not url:
            self.status_var.set("Enter a monitor URL to start refreshing.")
            return

        self.refresh_in_progress = True
        self.status_var.set("Refreshing from monitor page...")

        toner_threshold = max(1, min(100, self.toner_threshold_var.get()))
        fuser_threshold = max(1, min(100, self.fuser_threshold_var.get()))

        worker = threading.Thread(
            target=self._refresh_worker,
            args=(url, toner_threshold, fuser_threshold),
            daemon=True,
        )
        worker.start()

    def _refresh_worker(self, url: str, toner_threshold: int, fuser_threshold: int) -> None:
        try:
            html = fetch_html(url)
            records = parse_monitor_page(html)
            if not records:
                raise ValueError("No printer rows were parsed. The page layout may have changed.")

            low_alerts = build_low_alerts(records, toner_threshold, fuser_threshold)
            current_empties = build_current_empties(records)
            scan_time = now_iso()

            with self.state_lock:
                changes = self.state_store.reconcile(current_empties, scan_time)
                open_empties = self.state_store.get_open_empties()
                events = self.state_store.get_events()

            payload = {
                "scan_time": scan_time,
                "records": records,
                "low_alerts": low_alerts,
                "open_empties": open_empties,
                "events": events,
                "changes": changes,
            }
            self.root.after(0, lambda: self._apply_refresh(payload))
        except urllib.error.URLError as exc:
            self.root.after(0, lambda: self._refresh_error(f"Network error: {exc}"))
        except Exception as exc:  # broad by design to keep GUI responsive
            self.root.after(0, lambda: self._refresh_error(str(exc)))

    def _apply_refresh(self, payload: dict[str, Any]) -> None:
        self.refresh_in_progress = False

        self.records = payload["records"]
        self.low_alerts = payload["low_alerts"]
        self.open_empties = payload["open_empties"]
        self.events = payload["events"]

        self.summary_printer_var.set(str(len(self.records)))
        self.summary_low_var.set(str(len(self.low_alerts)))
        self.summary_empty_var.set(str(len(self.open_empties)))

        self.last_refresh_var.set(display_time(payload["scan_time"]))

        self._refresh_low_tree()
        self._refresh_empty_tree()
        self._refresh_history_tree()

        changes = payload.get("changes", {})
        self.status_var.set(
            f"Refresh complete. New empties: {changes.get('new_empties', 0)} | "
            f"Marked filled: {changes.get('new_filled', 0)}"
        )

        interval = max(1, self.refresh_interval_var.get())
        self.next_auto_refresh_epoch = datetime.now().timestamp() + (interval * 60)

    def _refresh_error(self, message: str) -> None:
        self.refresh_in_progress = False
        self.status_var.set(f"Refresh failed: {message}")

        interval = max(1, self.refresh_interval_var.get())
        self.next_auto_refresh_epoch = datetime.now().timestamp() + (interval * 60)

    def _refresh_low_tree(self) -> None:
        for item in self.low_tree.get_children():
            self.low_tree.delete(item)

        for idx, alert in enumerate(self.low_alerts):
            row_tag = "evenrow" if idx % 2 == 0 else "oddrow"
            self.low_tree.insert(
                "",
                "end",
                values=(
                    alert.printer_id,
                    alert.description,
                    alert.item,
                    alert.level,
                    alert.source,
                ),
                tags=(row_tag,),
            )

    def _refresh_empty_tree(self) -> None:
        selected = self.empty_tree.selection()
        selected_key = selected[0] if selected else None

        for item in self.empty_tree.get_children():
            self.empty_tree.delete(item)

        rows = sorted(
            self.open_empties.values(),
            key=lambda row: (row.get("printer_id", ""), row.get("tray", "")),
        )

        for idx, row in enumerate(rows):
            row_tag = "evenrow" if idx % 2 == 0 else "oddrow"
            key = row.get("key", f"{row.get('printer_id', '')}::{row.get('tray', '')}")
            self.empty_tree.insert(
                "",
                "end",
                iid=key,
                values=(
                    row.get("printer_id", ""),
                    row.get("description", ""),
                    row.get("tray", ""),
                    display_time(row.get("since", "")),
                    display_time(row.get("last_seen", "")),
                ),
                tags=(row_tag,),
            )

        if selected_key and self.empty_tree.exists(selected_key):
            self.empty_tree.selection_set(selected_key)

    def _refresh_history_tree(self) -> None:
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)

        recent = sorted(
            self.events,
            key=lambda event: parse_iso(event.get("timestamp", "")) or datetime.min,
            reverse=True,
        )[:1000]

        for idx, event in enumerate(recent):
            row_tag = "evenrow" if idx % 2 == 0 else "oddrow"
            self.history_tree.insert(
                "",
                "end",
                values=(
                    display_time(event.get("timestamp", "")),
                    event.get("event_type", ""),
                    event.get("printer_id", ""),
                    event.get("description", ""),
                    event.get("tray", ""),
                    event.get("note", ""),
                ),
                tags=(row_tag,),
            )

    def _auto_refresh_tick(self) -> None:
        if self.auto_refresh_var.get() and not self.refresh_in_progress:
            now_epoch = datetime.now().timestamp()
            if self.next_auto_refresh_epoch == 0.0 or now_epoch >= self.next_auto_refresh_epoch:
                self.trigger_refresh()

        self.root.after(1000, self._auto_refresh_tick)

    def _selected_tray(self) -> dict[str, str] | None:
        selected = self.empty_tree.selection()
        if not selected:
            return None
        key = selected[0]
        return self.open_empties.get(key)

    def generate_worknote(self) -> None:
        tray = self._selected_tray()
        if tray is None:
            return

        generated = self._build_worknote_text(tray)
        self.worknote_text.delete("1.0", tk.END)
        self.worknote_text.insert("1.0", generated)

    def _build_worknote_text(self, tray: dict[str, str]) -> str:
        now_label = display_time(now_iso())
        since = display_time(tray.get("since", ""))
        last_seen = display_time(tray.get("last_seen", ""))
        printer_id = tray.get("printer_id", "")
        description = tray.get("description", "")
        tray_name = tray.get("tray", "")
        status = tray.get("status_message", "") or "None"
        printer_text = tray.get("printer_text", "") or "None"
        mode = self.worknote_mode_var.get()

        if mode == "Refilled and Tested":
            return (
                f"[{now_label}] Refill completed for printer {printer_id} ({description}).\n"
                f"Issue addressed: {tray_name} empty.\n"
                f"First detected empty: {since}.\n"
                f"Most recent empty detection: {last_seen}.\n"
                "Actions performed:\n"
                f"- Arrived on site and verified {tray_name} was empty.\n"
                f"- Refilled {tray_name} with paper.\n"
                "- Ran a test print and confirmed successful output.\n"
                "- No additional supply/tray faults observed after refill.\n"
                "Printer returned to service."
            )

        return (
            f"[{now_label}] Monitoring alert for printer {printer_id} ({description}).\n"
            f"Current issue: {tray_name} empty.\n"
            f"First detected empty: {since}.\n"
            f"Most recent detection: {last_seen}.\n"
            f"Status Message: {status}\n"
            f"Printer Text: {printer_text}\n"
            "Planned action:\n"
            f"- Refill {tray_name}.\n"
            "- Run test print.\n"
            "- Update ticket with verification results."
        )

    def copy_worknote(self) -> None:
        text = self.worknote_text.get("1.0", "end-1c").strip()
        if not text:
            self.generate_worknote()
            text = self.worknote_text.get("1.0", "end-1c").strip()

        if not text:
            messagebox.showinfo("No Worknote", "Select an empty tray first.")
            return

        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status_var.set("Worknote copied to clipboard.")

    def mark_selected_filled(self) -> None:
        tray = self._selected_tray()
        if tray is None:
            messagebox.showinfo("No Selection", "Select an empty tray to mark filled.")
            return

        description = f"{tray.get('printer_id', '')} - {tray.get('description', '')} ({tray.get('tray', '')})"
        if not messagebox.askyesno("Mark Filled", f"Mark this tray as filled?\n\n{description}"):
            return

        key = tray.get("key", "")
        timestamp = now_iso()
        with self.state_lock:
            updated = self.state_store.manual_mark_filled(key, timestamp)
            if updated:
                self.open_empties = self.state_store.get_open_empties()
                self.events = self.state_store.get_events()

        if not updated:
            self.status_var.set("Tray was already resolved.")
            return

        self.summary_empty_var.set(str(len(self.open_empties)))
        self._refresh_empty_tree()
        self._refresh_history_tree()
        self.status_var.set("Tray marked filled manually.")

    def export_history_csv(self) -> None:
        if not self.events:
            messagebox.showinfo("No Data", "No history events to export yet.")
            return

        default_name = f"tray_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        output_path = filedialog.asksaveasfilename(
            title="Export Tray History",
            defaultextension=".csv",
            initialfile=default_name,
            filetypes=[("CSV Files", "*.csv")],
        )
        if not output_path:
            return

        fields = [
            "timestamp",
            "event_type",
            "printer_id",
            "description",
            "tray",
            "empty_since",
            "last_seen",
            "status_message",
            "printer_text",
            "note",
        ]

        try:
            with open(output_path, "w", newline="", encoding="utf-8") as handle:
                writer = csv.DictWriter(handle, fieldnames=fields)
                writer.writeheader()
                for event in self.events:
                    writer.writerow({field: event.get(field, "") for field in fields})
            self.status_var.set(f"Exported history CSV: {output_path}")
        except OSError as exc:
            messagebox.showerror("Export Failed", str(exc))


def main() -> None:
    root = tk.Tk()
    PrinterMonitorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
