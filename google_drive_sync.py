"""
google_drive_sync.py — SortIt v2.0
Gestisce autenticazione OAuth2 e polling Google Drive.
"""

import os
import io
import sys
import json
import threading
import time
import winreg
from typing import Optional, Callable

# ── Costanti registro ────────────────────────────────────────────────────────
REG_BASE        = r"Software\SortIt"
TOKEN_REG_KEY   = "google_token"
DRIVE_FOLDER_KEY = "google_drive_folder_id"
DRIVE_FOLDER_NAME_KEY = "google_drive_folder_name"
POLL_INTERVAL_KEY = "google_poll_interval"

SCOPES = ["https://www.googleapis.com/auth/drive"]

# ── Helpers registro ─────────────────────────────────────────────────────────

def _reg_read(name: str, default=None):
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_BASE)
        value, _ = winreg.QueryValueEx(key, name)
        winreg.CloseKey(key)
        return value
    except Exception:
        return default


def _reg_write(name: str, value: str):
    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_BASE)
    winreg.SetValueEx(key, name, 0, winreg.REG_SZ, str(value))
    winreg.CloseKey(key)


def _reg_delete(name: str):
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_BASE,
                             0, winreg.KEY_SET_VALUE)
        winreg.DeleteValue(key, name)
        winreg.CloseKey(key)
    except Exception:
        pass


# ── Autenticazione ───────────────────────────────────────────────────────────

def get_credentials_path() -> str:
    """Restituisce il percorso di credentials.json accanto all'eseguibile/script."""
    if getattr(sys, "frozen", False):
        # PyInstaller estrae i file bundled in sys._MEIPASS
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "credentials.json")


def credentials_file_exists() -> bool:
    return os.path.exists(get_credentials_path())


def load_token() -> Optional[dict]:
    raw = _reg_read(TOKEN_REG_KEY)
    if raw:
        try:
            return json.loads(raw)
        except Exception:
            pass
    return None


def save_token(creds) -> None:
    _reg_write(TOKEN_REG_KEY, creds.to_json())


def delete_token() -> None:
    _reg_delete(TOKEN_REG_KEY)


def authenticate(on_done: Callable, on_error: Callable) -> None:
    """
    Lancia il flusso OAuth in un thread separato.
    Chiama on_done(creds, email) oppure on_error(msg) al termine.
    """
    def _run():
        try:
            from google.oauth2.credentials import Credentials
            from google_auth_oauthlib.flow import InstalledAppFlow
            from googleapiclient.discovery import build

            creds = None
            token_data = load_token()
            if token_data:
                creds = Credentials.from_authorized_user_info(token_data, SCOPES)

            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    from google.auth.transport.requests import Request
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        get_credentials_path(), SCOPES)
                    creds = flow.run_local_server(port=0)
                save_token(creds)

            service = build("drive", "v3", credentials=creds)
            profile = service.about().get(fields="user").execute()
            email = profile.get("user", {}).get("emailAddress", "")
            on_done(creds, email)

        except Exception as e:
            on_error(str(e))

    threading.Thread(target=_run, daemon=True).start()


def get_user_email(creds) -> str:
    try:
        from googleapiclient.discovery import build
        service = build("drive", "v3", credentials=creds)
        profile = service.about().get(fields="user").execute()
        return profile.get("user", {}).get("emailAddress", "")
    except Exception:
        return ""


def list_drive_folders(creds) -> list:
    """Restituisce lista di dict {id, name} delle cartelle nel My Drive."""
    try:
        from googleapiclient.discovery import build
        service = build("drive", "v3", credentials=creds)
        results = service.files().list(
            q="mimeType='application/vnd.google-apps.folder' and trashed=false and 'root' in parents",
            fields="files(id, name)",
            orderBy="name"
        ).execute()
        return results.get("files", [])
    except Exception:
        return []


# ── Polling ──────────────────────────────────────────────────────────────────

class DrivePoller:
    """
    Controlla periodicamente una cartella Drive.
    Scarica i nuovi file nella cartella locale e li elimina da Drive.
    """

    def __init__(self,
                 creds,
                 drive_folder_id: str,
                 local_folder: str,
                 interval_minutes: int,
                 on_file_downloaded: Callable,   # callback(local_path) → None
                 log_callback: Callable):         # callback(msg, level) → None
        self.creds            = creds
        self.drive_folder_id  = drive_folder_id
        self.local_folder     = local_folder
        self.interval         = interval_minutes * 60
        self.on_file_downloaded = on_file_downloaded
        self.log              = log_callback
        self._stop_event      = threading.Event()
        self._thread          = None
        self._seen_ids: set   = set()

    def start(self):
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def stop(self):
        self._stop_event.set()
        if self._thread:
            self._thread.join(timeout=5)

    @property
    def is_running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def _loop(self):
        self.log("🔄 Polling Google Drive avviato.", "success")
        while not self._stop_event.is_set():
            try:
                self._poll()
            except Exception as e:
                self.log(f"[ERRORE] Polling Drive: {e}", "error")
            self._stop_event.wait(self.interval)

    def _poll(self):
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseDownload

        # Rinnova credenziali se scadute
        if self.creds.expired and self.creds.refresh_token:
            from google.auth.transport.requests import Request
            self.creds.refresh(Request())
            save_token(self.creds)

        service = build("drive", "v3", credentials=self.creds)

        # Cerca file (non cartelle, non nel cestino) nella cartella Drive
        q = (f"'{self.drive_folder_id}' in parents "
             f"and mimeType != 'application/vnd.google-apps.folder' "
             f"and trashed = false")
        results = service.files().list(
            q=q,
            fields="files(id, name, mimeType, size)"
        ).execute()
        files = results.get("files", [])

        new_files = [f for f in files if f["id"] not in self._seen_ids]
        if not new_files:
            return

        os.makedirs(self.local_folder, exist_ok=True)

        for file in new_files:
            fid   = file["id"]
            fname = file["name"]
            mime  = file.get("mimeType", "")

            try:
                # Google Docs → esporta come PDF
                if mime.startswith("application/vnd.google-apps"):
                    export_map = {
                        "application/vnd.google-apps.document":
                            ("application/pdf", ".pdf"),
                        "application/vnd.google-apps.spreadsheet":
                            ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx"),
                        "application/vnd.google-apps.presentation":
                            ("application/vnd.openxmlformats-officedocument.presentationml.presentation", ".pptx"),
                    }
                    if mime not in export_map:
                        self.log(f"[SKIP] {fname} — tipo Google non supportato.", "warning")
                        self._seen_ids.add(fid)
                        continue
                    export_mime, ext = export_map[mime]
                    if not fname.endswith(ext):
                        fname += ext
                    request = service.files().export_media(
                        fileId=fid, mimeType=export_mime)
                else:
                    request = service.files().get_media(fileId=fid)

                dest_path = os.path.join(self.local_folder, fname)
                # Gestione conflitti nome
                base, ext_ = os.path.splitext(dest_path)
                counter = 1
                while os.path.exists(dest_path):
                    dest_path = f"{base}_{counter}{ext_}"
                    counter += 1

                buf = io.BytesIO()
                downloader = MediaIoBaseDownload(buf, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()

                with open(dest_path, "wb") as f:
                    f.write(buf.getvalue())

                self.log(f"☁ Scaricato da Drive: {fname}", "info")

                # Elimina da Drive
                service.files().delete(fileId=fid).execute()
                self.log(f"🗑 Eliminato da Drive: {fname}", "info")

                self._seen_ids.add(fid)

                # Notifica al FolderWatcher / SortIt
                self.on_file_downloaded(dest_path)

            except Exception as e:
                self.log(f"[ERRORE] Download {fname}: {e}", "error")
                self._seen_ids.add(fid)  # evita loop infinito
