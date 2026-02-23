"""Tkinter desktop GUI for printer key checkout/return tracking."""

from __future__ import annotations

from pathlib import Path
import sys
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
from typing import Callable

from storage import (
    GoogleDriveStore,
    StorageError,
    append_event,
    create_google_drive_store,
    ensure_user_log_exists,
    export_aggregate,
    get_current_checkout,
    get_default_google_folder_url,
    get_env_google_client_id,
    get_env_google_client_secret,
    get_env_google_folder_id,
    get_env_google_folder_url,
    get_saved_google_client_id,
    get_saved_google_client_secret,
    get_saved_google_email_hint,
    get_saved_google_folder_id,
    get_saved_google_folder_url,
    get_whats_out,
    make_checkout_event,
    make_return_event,
    save_setup,
)
from utils import clean_text, display_timestamp

APP_NAME = "Printer Key Checkout Tracker"
APP_VERSION = "1.4.0"
CREDITS = "Jack Shetterly"


class KeyTrackerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_NAME)
        self.root.minsize(760, 500)

        self.store: GoogleDriveStore | None = None
        self.google_client_id = ""
        self.google_folder_url = ""
        self.user_id = ""
        self.status_var = tk.StringVar(value="Ready.")
        self.context_var = tk.StringVar(value="")
        self.colors = self._build_color_palette()
        self.font_family = self._select_font_family()
        self.ready = False

        self._configure_styles()
        self._apply_window_icon()

        if not self._initialize_context():
            return

        self._build_home_screen()
        self.ready = True

    def _build_color_palette(self) -> dict[str, str]:
        return {
            "window_bg": "#E9EEF5",
            "surface_bg": "#F4F7FC",
            "card_bg": "#FFFFFF",
            "border": "#D5DEEA",
            "text_primary": "#16243C",
            "text_secondary": "#5D6E88",
            "primary": "#2F5FAE",
            "primary_active": "#254F93",
            "secondary_bg": "#EAF0FA",
            "secondary_active": "#DDE7F7",
        }

    def _select_font_family(self) -> str:
        if sys.platform == "darwin":
            return "SF Pro Text"
        if sys.platform.startswith("win"):
            return "Segoe UI"
        return "Helvetica"

    def _configure_styles(self) -> None:
        self.root.geometry("920x620")
        self.root.configure(bg=self.colors["window_bg"])

        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        self.root.option_add("*Font", (self.font_family, 11))

        style.configure(
            "Root.TFrame",
            background=self.colors["window_bg"],
        )
        style.configure(
            "Surface.TFrame",
            background=self.colors["surface_bg"],
            borderwidth=1,
            relief="solid",
        )
        style.configure(
            "Card.TFrame",
            background=self.colors["card_bg"],
            borderwidth=1,
            relief="solid",
        )

        style.configure(
            "TLabel",
            background=self.colors["window_bg"],
            foreground=self.colors["text_primary"],
            font=(self.font_family, 11),
        )
        style.configure(
            "Title.TLabel",
            background=self.colors["card_bg"],
            foreground=self.colors["text_primary"],
            font=(self.font_family, 24, "bold"),
        )
        style.configure(
            "Subtitle.TLabel",
            background=self.colors["card_bg"],
            foreground=self.colors["text_secondary"],
            font=(self.font_family, 11),
        )
        style.configure(
            "Context.TLabel",
            background=self.colors["card_bg"],
            foreground=self.colors["text_primary"],
            font=(self.font_family, 10),
        )
        style.configure(
            "SectionTitle.TLabel",
            background=self.colors["window_bg"],
            foreground=self.colors["text_primary"],
            font=(self.font_family, 13, "bold"),
        )
        style.configure(
            "CardTitle.TLabel",
            background=self.colors["card_bg"],
            foreground=self.colors["text_primary"],
            font=(self.font_family, 12, "bold"),
        )
        style.configure(
            "CardBody.TLabel",
            background=self.colors["card_bg"],
            foreground=self.colors["text_secondary"],
            font=(self.font_family, 10),
        )
        style.configure(
            "Status.TLabel",
            background=self.colors["surface_bg"],
            foreground=self.colors["text_secondary"],
            font=(self.font_family, 10),
        )

        style.configure(
            "TButton",
            font=(self.font_family, 11),
            padding=(14, 10),
        )
        style.configure(
            "Primary.TButton",
            font=(self.font_family, 11, "bold"),
            background=self.colors["primary"],
            foreground="#FFFFFF",
            borderwidth=0,
            relief="flat",
            padding=(14, 10),
        )
        style.map(
            "Primary.TButton",
            background=[
                ("active", self.colors["primary_active"]),
                ("pressed", self.colors["primary_active"]),
            ],
            foreground=[("disabled", "#D7E2F4"), ("!disabled", "#FFFFFF")],
        )
        style.configure(
            "Secondary.TButton",
            font=(self.font_family, 11, "bold"),
            background=self.colors["secondary_bg"],
            foreground=self.colors["text_primary"],
            borderwidth=0,
            relief="flat",
            padding=(14, 10),
        )
        style.map(
            "Secondary.TButton",
            background=[
                ("active", self.colors["secondary_active"]),
                ("pressed", self.colors["secondary_active"]),
            ],
        )

        style.configure(
            "TEntry",
            fieldbackground="#FFFFFF",
            foreground=self.colors["text_primary"],
            padding=6,
        )

        style.configure(
            "Treeview",
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            foreground=self.colors["text_primary"],
            rowheight=28,
            borderwidth=1,
            relief="solid",
        )
        style.configure(
            "Treeview.Heading",
            background="#EAF0FA",
            foreground=self.colors["text_primary"],
            font=(self.font_family, 10, "bold"),
            relief="flat",
        )
        style.map(
            "Treeview",
            background=[("selected", "#DDE9FB")],
            foreground=[("selected", self.colors["text_primary"])],
        )

    def _resolve_resource_path(self, relative_path: str) -> Path:
        if hasattr(sys, "_MEIPASS"):
            return Path(getattr(sys, "_MEIPASS")) / relative_path
        return Path(__file__).resolve().parent / relative_path

    def _apply_window_icon(self) -> None:
        icon_path = self._resolve_resource_path("resources/AppIcon-256.png")
        if not icon_path.exists():
            return

        try:
            image = tk.PhotoImage(file=str(icon_path))
            self.root.iconphoto(True, image)
            self.root._icon_ref = image  # type: ignore[attr-defined]
        except Exception:
            pass

    def _initialize_context(self) -> bool:
        store = self._connect_google_drive_store_sso(force_prompt_credentials=False)
        if store is None:
            return False
        self.store = store

        user_id = self._prompt_user_id()
        if user_id is None:
            return False
        self.user_id = user_id

        try:
            ensure_user_log_exists(store, self.user_id)
        except StorageError as exc:
            messagebox.showerror(
                "Log Initialization Error",
                f"Could not initialize user event log.\n\n{exc}",
                parent=self.root,
            )
            return False

        return True

    def _prompt_google_client_id(self, initial: str = "") -> str | None:
        while True:
            value = simpledialog.askstring(
                "Google OAuth Client ID",
                (
                    "Enter Google OAuth Desktop App Client ID.\n\n"
                    "Users sign in with Google in a browser."
                ),
                parent=self.root,
                initialvalue=initial,
            )
            if value is None:
                return None

            client_id = clean_text(value)
            if client_id:
                return client_id

            messagebox.showerror(
                "Input Error",
                "Google Client ID is required.",
                parent=self.root,
            )

    def _prompt_google_client_secret(self, initial: str = "") -> str | None:
        while True:
            value = simpledialog.askstring(
                "Google OAuth Client Secret",
                "Enter Google OAuth Desktop App Client Secret:",
                parent=self.root,
                initialvalue=initial,
                show="*",
            )
            if value is None:
                return None

            client_secret = clean_text(value)
            if client_secret:
                return client_secret

            messagebox.showerror(
                "Input Error",
                "Google Client Secret is required.",
                parent=self.root,
            )

    def _prompt_google_folder(self, initial: str = "") -> str | None:
        while True:
            value = simpledialog.askstring(
                "Google Drive Folder",
                (
                    "Enter shared Google Drive folder URL (or folder ID).\n\n"
                    "Example: https://drive.google.com/drive/folders/<id>"
                ),
                parent=self.root,
                initialvalue=initial,
            )
            if value is None:
                return None

            folder_value = clean_text(value)
            if folder_value:
                return folder_value

            messagebox.showerror(
                "Input Error",
                "Google Drive folder URL or ID is required.",
                parent=self.root,
            )

    def _connect_google_drive_store_sso(self, force_prompt_credentials: bool) -> GoogleDriveStore | None:
        env_client_id = get_env_google_client_id()
        env_client_secret = get_env_google_client_secret()
        env_folder_url = get_env_google_folder_url()
        env_folder_id = get_env_google_folder_id()

        saved_client_id = ""
        saved_client_secret = ""
        saved_folder_url = ""
        saved_folder_id = ""
        saved_email_hint = ""
        try:
            saved_client_id = get_saved_google_client_id()
            saved_client_secret = get_saved_google_client_secret()
            saved_folder_url = get_saved_google_folder_url()
            saved_folder_id = get_saved_google_folder_id()
            saved_email_hint = get_saved_google_email_hint()
        except StorageError as exc:
            messagebox.showwarning(
                "Config Warning",
                f"Could not read saved settings.\n\n{exc}",
                parent=self.root,
            )

        client_id = env_client_id or saved_client_id
        client_secret = env_client_secret or saved_client_secret
        folder_url = env_folder_url or saved_folder_url or get_default_google_folder_url()
        folder_id = env_folder_id or saved_folder_id
        email_hint = saved_email_hint

        while True:
            if force_prompt_credentials or not client_id:
                prompted = self._prompt_google_client_id(initial=client_id)
                if prompted is None:
                    return None
                client_id = prompted
                force_prompt_credentials = False

            if not client_secret:
                prompted_secret = self._prompt_google_client_secret(initial="")
                if prompted_secret is None:
                    return None
                client_secret = prompted_secret

            if not folder_url and not folder_id:
                prompted_folder = self._prompt_google_folder(initial=folder_url)
                if prompted_folder is None:
                    return None
                folder_url = prompted_folder

            try:
                store = create_google_drive_store(
                    google_client_id=client_id,
                    google_client_secret=client_secret,
                    google_folder_url=folder_url,
                    google_folder_id=folder_id,
                    email_hint=email_hint,
                )
                self.google_client_id = client_id
                self.google_folder_url = store.folder_url
                save_setup(
                    google_client_id=client_id,
                    google_client_secret=client_secret,
                    google_folder_url=store.folder_url,
                    google_folder_id=store.root_folder_id,
                    email_hint=store.email_hint,
                )
                return store
            except StorageError as exc:
                action = messagebox.askyesnocancel(
                    "Unable to Connect to Google Drive",
                    (
                        "Google OAuth / Drive connection failed.\n\n"
                        "Yes: Retry sign-in\n"
                        "No: Re-enter Client ID, Client Secret, or folder\n"
                        "Cancel: Exit\n\n"
                        f"Details:\n{exc}"
                    ),
                    parent=self.root,
                )
                if action is None:
                    return None
                if action is False:
                    prompted_id = self._prompt_google_client_id(initial=client_id)
                    if prompted_id is None:
                        return None
                    client_id = prompted_id

                    prompted_secret = self._prompt_google_client_secret(initial=client_secret)
                    if prompted_secret is None:
                        return None
                    client_secret = prompted_secret

                    prompted_folder = self._prompt_google_folder(initial=folder_url or folder_id)
                    if prompted_folder is None:
                        return None
                    folder_url = prompted_folder
                    folder_id = ""

    def _prompt_user_id(self) -> str | None:
        while True:
            value = simpledialog.askstring(
                "User ID",
                "Enter your UserId for this session:",
                parent=self.root,
            )
            if value is None:
                return None

            user_id = clean_text(value)
            if user_id:
                return user_id

            messagebox.showerror("Input Error", "UserId is required.", parent=self.root)

    def _build_home_screen(self) -> None:
        container = ttk.Frame(self.root, style="Root.TFrame", padding=(22, 20))
        container.pack(fill="both", expand=True)

        hero = ttk.Frame(container, style="Card.TFrame", padding=(20, 18))
        hero.pack(fill="x")

        ttk.Label(hero, text=APP_NAME, style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            hero,
            text="Checkout and return tracking for shared sandbox printer keys.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        self._refresh_context_text()
        ttk.Label(
            hero,
            textvariable=self.context_var,
            style="Context.TLabel",
            justify="left",
            wraplength=840,
        ).pack(anchor="w", pady=(14, 0))

        ttk.Label(container, text="Actions", style="SectionTitle.TLabel").pack(
            anchor="w", pady=(18, 8)
        )

        action_grid = ttk.Frame(container, style="Root.TFrame")
        action_grid.pack(fill="both", expand=True)
        for index in range(2):
            action_grid.columnconfigure(index, weight=1, uniform="actions")

        actions = [
            ("Check Out", "Record a new key checkout.", self.open_checkout_form, "Primary.TButton"),
            ("Return", "Record a key return.", self.open_return_form, "Primary.TButton"),
            ("What's Out", "View keys currently checked out.", self.open_whats_out_window, "Secondary.TButton"),
            ("Export", "Generate the combined Excel workbook.", self.export_workbook, "Secondary.TButton"),
            ("Reconnect", "Refresh Google sign-in and storage link.", self.reconnect_google_drive, "Secondary.TButton"),
            ("About", "View version and project details.", self.show_about, "Secondary.TButton"),
        ]
        for index, (title, description, callback, button_style) in enumerate(actions):
            self._add_action_card(
                parent=action_grid,
                row=index // 2,
                column=index % 2,
                title=title,
                description=description,
                callback=callback,
                button_style=button_style,
            )

        status_frame = ttk.Frame(container, style="Surface.TFrame", padding=(14, 10))
        status_frame.pack(fill="x", side="bottom", pady=(14, 0))
        ttk.Label(status_frame, textvariable=self.status_var, style="Status.TLabel").pack(anchor="w")

    def _add_action_card(
        self,
        parent: ttk.Frame,
        row: int,
        column: int,
        title: str,
        description: str,
        callback: Callable[[], None],
        button_style: str,
    ) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=(14, 12))
        card.grid(row=row, column=column, padx=8, pady=8, sticky="nsew")
        parent.rowconfigure(row, weight=1)

        ttk.Label(card, text=title, style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            card,
            text=description,
            style="CardBody.TLabel",
            wraplength=360,
            justify="left",
        ).pack(anchor="w", pady=(4, 10))
        ttk.Button(
            card,
            text=title,
            style=button_style,
            command=callback,
        ).pack(fill="x")

    def _refresh_context_text(self) -> None:
        signed_in_as = "(unknown)"
        if self.store is not None and clean_text(self.store.email_hint):
            signed_in_as = self.store.email_hint
        folder_text = clean_text(self.google_folder_url) or "(not configured)"
        self.context_var.set(
            f"Session UserId: {self.user_id}\n"
            f"Drive Folder: {folder_text}\n"
            f"Signed In As: {signed_in_as}"
        )

    def reconnect_google_drive(self) -> None:
        store = self._connect_google_drive_store_sso(force_prompt_credentials=True)
        if store is None:
            return

        self.store = store
        try:
            ensure_user_log_exists(store, self.user_id)
        except StorageError as exc:
            messagebox.showerror("Storage Error", str(exc), parent=self.root)
            return

        self._refresh_context_text()
        self.status_var.set("Google Drive connection updated.")
        messagebox.showinfo(
            "Connected",
            "Google Drive OAuth connection is active.",
            parent=self.root,
        )

    def show_about(self) -> None:
        messagebox.showinfo(
            f"About {APP_NAME}",
            (
                f"{APP_NAME}\n"
                f"Version: {APP_VERSION}\n\n"
                "Tracks key checkouts/returns using per-user append-only CSV logs.\n"
                "Storage is in Google Drive via OAuth sign-in.\n\n"
                f"Credits: {CREDITS}"
            ),
            parent=self.root,
        )

    def open_checkout_form(self) -> None:
        window = tk.Toplevel(self.root)
        window.title("Check Out Key")
        window.resizable(False, False)
        window.configure(bg=self.colors["window_bg"])

        frame = ttk.Frame(window, style="Root.TFrame", padding=(16, 14))
        frame.pack(fill="both", expand=True)
        card = ttk.Frame(frame, style="Card.TFrame", padding=(16, 14))
        card.pack(fill="both", expand=True)

        ttk.Label(card, text="Check Out Key", style="CardTitle.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 2)
        )
        ttk.Label(
            card,
            text="Required fields are marked with *.",
            style="CardBody.TLabel",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 10))

        key_id_var = tk.StringVar()
        from_location_var = tk.StringVar()
        to_location_var = tk.StringVar()
        destination_var = tk.StringVar()

        ttk.Label(card, text="KeyId *").grid(row=2, column=0, sticky="w", pady=3)
        ttk.Entry(card, textvariable=key_id_var, width=42).grid(
            row=2, column=1, sticky="ew", pady=3
        )

        ttk.Label(card, text="FromLocation *").grid(row=3, column=0, sticky="w", pady=3)
        ttk.Entry(card, textvariable=from_location_var, width=42).grid(
            row=3, column=1, sticky="ew", pady=3
        )

        ttk.Label(card, text="ToLocation *").grid(row=4, column=0, sticky="w", pady=3)
        ttk.Entry(card, textvariable=to_location_var, width=42).grid(
            row=4, column=1, sticky="ew", pady=3
        )

        ttk.Label(card, text="PrinterOrDestination").grid(
            row=5, column=0, sticky="w", pady=3
        )
        ttk.Entry(card, textvariable=destination_var, width=42).grid(
            row=5, column=1, sticky="ew", pady=3
        )

        ttk.Label(card, text="Notes").grid(row=6, column=0, sticky="nw", pady=3)
        notes_widget = tk.Text(
            card,
            width=40,
            height=4,
            bg="#FFFFFF",
            fg=self.colors["text_primary"],
            relief="solid",
            bd=1,
            insertbackground=self.colors["text_primary"],
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["primary"],
            font=(self.font_family, 11),
        )
        notes_widget.grid(row=6, column=1, sticky="ew", pady=3)

        card.columnconfigure(1, weight=1)

        def submit() -> None:
            key_id = clean_text(key_id_var.get())
            from_location = clean_text(from_location_var.get())
            to_location = clean_text(to_location_var.get())
            destination = clean_text(destination_var.get())
            notes = clean_text(notes_widget.get("1.0", "end"))

            if not key_id or not from_location or not to_location:
                messagebox.showerror(
                    "Input Error",
                    "KeyId, FromLocation, and ToLocation are required.",
                    parent=window,
                )
                return

            assert self.store is not None
            try:
                currently_out, warnings = get_current_checkout(self.store, key_id)
                if currently_out is not None:
                    messagebox.showerror(
                        "Already Checked Out",
                        (
                            f"Key '{currently_out.KeyId}' is already OUT.\n\n"
                            f"Checked out by: {currently_out.UserId}\n"
                            f"Since: {display_timestamp(currently_out.Timestamp)}\n"
                            f"To location: {currently_out.ToLocation or '-'}\n"
                            f"Destination: {currently_out.PrinterOrDestination or '-'}"
                        ),
                        parent=window,
                    )
                    return

                event = make_checkout_event(
                    user_id=self.user_id,
                    key_id=key_id,
                    from_location=from_location,
                    to_location=to_location,
                    printer_or_destination=destination,
                    notes=notes,
                )
                append_event(self.store, self.user_id, event)

                self.status_var.set(
                    f"Checked out key '{event.KeyId}' at {display_timestamp(event.Timestamp)}."
                )
                messagebox.showinfo("Saved", "Checkout recorded.", parent=window)
                self._show_warnings_if_any(warnings, "Data Warnings")
                window.destroy()
            except StorageError as exc:
                messagebox.showerror("Storage Error", str(exc), parent=window)

        button_row = ttk.Frame(card, style="Card.TFrame")
        button_row.grid(row=7, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(button_row, text="Cancel", command=window.destroy).pack(
            side="right", padx=(6, 0)
        )
        ttk.Button(button_row, text="Submit", style="Primary.TButton", command=submit).pack(
            side="right"
        )

    def open_return_form(self) -> None:
        window = tk.Toplevel(self.root)
        window.title("Return Key")
        window.resizable(False, False)
        window.configure(bg=self.colors["window_bg"])

        frame = ttk.Frame(window, style="Root.TFrame", padding=(16, 14))
        frame.pack(fill="both", expand=True)
        card = ttk.Frame(frame, style="Card.TFrame", padding=(16, 14))
        card.pack(fill="both", expand=True)

        ttk.Label(card, text="Return Key", style="CardTitle.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 2)
        )
        ttk.Label(
            card,
            text="Required fields are marked with *.",
            style="CardBody.TLabel",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 10))

        key_id_var = tk.StringVar()
        returned_to_var = tk.StringVar()

        ttk.Label(card, text="KeyId *").grid(row=2, column=0, sticky="w", pady=3)
        ttk.Entry(card, textvariable=key_id_var, width=42).grid(
            row=2, column=1, sticky="ew", pady=3
        )

        ttk.Label(card, text="ReturnedToLocation *").grid(
            row=3, column=0, sticky="w", pady=3
        )
        ttk.Entry(card, textvariable=returned_to_var, width=42).grid(
            row=3, column=1, sticky="ew", pady=3
        )

        ttk.Label(card, text="Notes").grid(row=4, column=0, sticky="nw", pady=3)
        notes_widget = tk.Text(
            card,
            width=40,
            height=4,
            bg="#FFFFFF",
            fg=self.colors["text_primary"],
            relief="solid",
            bd=1,
            insertbackground=self.colors["text_primary"],
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["primary"],
            font=(self.font_family, 11),
        )
        notes_widget.grid(row=4, column=1, sticky="ew", pady=3)

        card.columnconfigure(1, weight=1)

        def submit() -> None:
            key_id = clean_text(key_id_var.get())
            returned_to = clean_text(returned_to_var.get())
            notes = clean_text(notes_widget.get("1.0", "end"))

            if not key_id or not returned_to:
                messagebox.showerror(
                    "Input Error",
                    "KeyId and ReturnedToLocation are required.",
                    parent=window,
                )
                return

            assert self.store is not None
            try:
                currently_out, warnings = get_current_checkout(self.store, key_id)
                if currently_out is None:
                    proceed = messagebox.askyesno(
                        "No Active Checkout Found",
                        (
                            f"Key '{key_id}' is not currently OUT.\n\n"
                            "Do you still want to record a return event?"
                        ),
                        parent=window,
                    )
                    if not proceed:
                        return

                event = make_return_event(
                    user_id=self.user_id,
                    key_id=key_id,
                    returned_to_location=returned_to,
                    notes=notes,
                )
                append_event(self.store, self.user_id, event)

                self.status_var.set(
                    f"Returned key '{event.KeyId}' at {display_timestamp(event.Timestamp)}."
                )
                messagebox.showinfo("Saved", "Return recorded.", parent=window)
                self._show_warnings_if_any(warnings, "Data Warnings")
                window.destroy()
            except StorageError as exc:
                messagebox.showerror("Storage Error", str(exc), parent=window)

        button_row = ttk.Frame(card, style="Card.TFrame")
        button_row.grid(row=5, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(button_row, text="Cancel", command=window.destroy).pack(
            side="right", padx=(6, 0)
        )
        ttk.Button(button_row, text="Submit", style="Primary.TButton", command=submit).pack(
            side="right"
        )

    def open_whats_out_window(self) -> None:
        window = tk.Toplevel(self.root)
        window.title("What's Out")
        window.geometry("860x420")
        window.configure(bg=self.colors["window_bg"])

        frame = ttk.Frame(window, style="Root.TFrame", padding=(14, 12))
        frame.pack(fill="both", expand=True)
        card = ttk.Frame(frame, style="Card.TFrame", padding=(14, 12))
        card.pack(fill="both", expand=True)

        ttk.Label(card, text="What's Out", style="CardTitle.TLabel").pack(anchor="w")

        count_var = tk.StringVar(value="Loading...")
        ttk.Label(card, textvariable=count_var, style="CardBody.TLabel").pack(
            anchor="w", pady=(0, 8)
        )

        columns = (
            "KeyId",
            "CheckedOutBy",
            "TimeOut",
            "ToLocation",
            "PrinterOrDestination",
        )
        tree = ttk.Treeview(card, columns=columns, show="headings", height=14)
        headings = {
            "KeyId": 110,
            "CheckedOutBy": 140,
            "TimeOut": 170,
            "ToLocation": 170,
            "PrinterOrDestination": 220,
        }
        for column, width in headings.items():
            tree.heading(column, text=column)
            tree.column(column, width=width, anchor="w")

        y_scroll = ttk.Scrollbar(card, orient="vertical", command=tree.yview)
        x_scroll = ttk.Scrollbar(card, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        tree.pack(fill="both", expand=True, side="left")
        y_scroll.pack(fill="y", side="right")
        x_scroll.pack(fill="x", side="bottom")

        def refresh() -> None:
            assert self.store is not None
            try:
                rows, warnings = get_whats_out(self.store)
            except StorageError as exc:
                messagebox.showerror("Storage Error", str(exc), parent=window)
                return

            tree.delete(*tree.get_children())
            for row in rows:
                tree.insert(
                    "",
                    "end",
                    values=(
                        row.KeyId,
                        row.CheckedOutBy,
                        display_timestamp(row.TimeOut),
                        row.ToLocation,
                        row.PrinterOrDestination,
                    ),
                )

            count_var.set(f"{len(rows)} key(s) currently OUT.")
            self.status_var.set(f"Loaded 'What's Out' view with {len(rows)} active key(s).")
            self._show_warnings_if_any(warnings, "Data Warnings")

        controls = ttk.Frame(card, style="Card.TFrame")
        controls.pack(fill="x", pady=(8, 0))
        ttk.Button(controls, text="Refresh", style="Secondary.TButton", command=refresh).pack(
            side="right"
        )

        refresh()

    def export_workbook(self) -> None:
        assert self.store is not None

        try:
            output_ref, warnings = export_aggregate(self.store)
        except StorageError as exc:
            messagebox.showerror("Export Error", str(exc), parent=self.root)
            return

        self.status_var.set("Exported workbook to Google Drive output folder.")
        messagebox.showinfo(
            "Export Complete",
            f"Workbook updated:\n{output_ref}",
            parent=self.root,
        )
        self._show_warnings_if_any(warnings, "Export Warnings")

    def _show_warnings_if_any(self, warnings: list[str], title: str) -> None:
        if not warnings:
            return

        preview_count = 8
        preview_lines = warnings[:preview_count]
        extra_count = len(warnings) - len(preview_lines)
        extra_text = f"\n...and {extra_count} more warning(s)." if extra_count > 0 else ""
        message = (
            f"{len(warnings)} warning(s) encountered while reading logs.\n\n"
            + "\n".join(preview_lines)
            + extra_text
        )
        messagebox.showwarning(title, message, parent=self.root)


def main() -> None:
    root = tk.Tk()
    app = KeyTrackerApp(root)
    if not app.ready:
        root.destroy()
        return
    root.mainloop()


if __name__ == "__main__":
    main()
