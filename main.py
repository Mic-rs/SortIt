import os
import sys
import json
import shutil
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
from rules_manager import RulesManager
from watcher import FolderWatcher
import google_drive_sync as gds

# Percorso base: funziona sia da script che da exe PyInstaller
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")  # non più usato, mantenuto per compatibilità
RULES_FILE  = os.path.join(BASE_DIR, "rules.json")   # non più usato
ICON_FILE   = os.path.join(BASE_DIR, "icon.ico")

ctk.set_default_color_theme("blue")


# ─────────────────────────────────────────────
#  Finestra regole
# ─────────────────────────────────────────────

class RulesWindow(ctk.CTkToplevel):
    def __init__(self, parent, rules_manager):
        super().__init__(parent)
        self.withdraw()
        self.rm = rules_manager
        self.title("Gestione Regole — SortIt")
        self.geometry("740x520")
        self.resizable(False, False)
        self._build()
        self._refresh_list()
        self.after(120, self.deiconify)
        self.after(150, self.lift)

    def _build(self):
        ctk.CTkLabel(self, text="Categorie configurate",
                     font=ctk.CTkFont(size=15, weight="bold")).pack(
                     anchor="w", padx=20, pady=(16, 6))

        style = ttk.Style()
        style.theme_use("clam")

        is_dark = ctk.get_appearance_mode().lower() == "dark"
        tree_bg     = "#2b2b2b" if is_dark else "#f0f0f0"
        tree_fg     = "#e0e0e0" if is_dark else "#1a1a1a"
        tree_sel    = "#3a7bd5"
        heading_bg  = "#1e1e1e" if is_dark else "#dde3ea"
        heading_fg  = "#00d2ff" if is_dark else "#0077cc"
        field_bg    = tree_bg

        style.configure("Rules.Treeview",
                        rowheight=28, font=("Segoe UI", 9),
                        borderwidth=0,
                        background=tree_bg,
                        foreground=tree_fg,
                        fieldbackground=field_bg)
        style.configure("Rules.Treeview.Heading",
                        font=("Segoe UI", 9, "bold"), relief="flat",
                        background=heading_bg, foreground=heading_fg)
        style.map("Rules.Treeview",
                  background=[("selected", tree_sel)],
                  foreground=[("selected", "white")])

        frame = ctk.CTkFrame(self, corner_radius=10)
        frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        cols = ("Nome", "Cartella", "Estensioni", "Parole chiave")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings",
                                 height=14, style="Rules.Treeview")
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=155 if col != "Parole chiave" else 235)

        sb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        sb.pack(side="right", fill="y")

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(pady=(0, 16))
        for label, cmd in [("➕  Aggiungi", self._add),
                            ("✏  Modifica",  self._edit),
                            ("🗑  Elimina",   self._delete)]:
            ctk.CTkButton(btn_row, text=label, width=130,
                          corner_radius=8, command=cmd).pack(side="left", padx=6)

    def _refresh_list(self):
        self.tree.delete(*self.tree.get_children())
        for cat in self.rm.categories:
            kws = cat.get("keywords", [])
            self.tree.insert("", "end", values=(
                cat["name"], cat["folder"],
                ", ".join(cat.get("extensions", [])),
                ", ".join(kws[:4]) + ("…" if len(kws) > 4 else "")
            ))

    def _add(self):
        CategoryDialog(self, self.rm, None, self._refresh_list)

    def _edit(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Selezione", "Seleziona una categoria.", parent=self)
            return
        name = self.tree.item(sel[0])["values"][0]
        CategoryDialog(self, self.rm, self.rm.get_category(name), self._refresh_list)

    def _delete(self):
        sel = self.tree.selection()
        if not sel:
            return
        name = self.tree.item(sel[0])["values"][0]
        if messagebox.askyesno("Elimina", f"Eliminare '{name}'?", parent=self):
            self.rm.remove_category(name)
            self._refresh_list()


class CategoryDialog(ctk.CTkToplevel):
    def __init__(self, parent, rules_manager, category, on_save):
        super().__init__(parent)
        self.withdraw()
        self.rm  = rules_manager
        self.cat = category
        self.on_save = on_save
        self.title("Modifica categoria" if category else "Nuova categoria")
        self.geometry("560x500")
        self.resizable(False, False)
        self._build()
        self.after(120, self.deiconify)
        self.after(150, self.lift)

    def _build(self):
        ctk.CTkLabel(self,
                     text="Modifica categoria" if self.cat else "Nuova categoria",
                     font=ctk.CTkFont(size=14, weight="bold")).grid(
                     row=0, column=0, columnspan=2, padx=20, pady=(16, 10), sticky="w")

        fields = [
            ("Nome categoria",          "name"),
            ("Cartella destinazione",   "folder"),
            ("Estensioni (virgola)",    "extensions"),
            ("Parole chiave (virgola)", "keywords"),
            ("Template rinomina",       "rename_template"),
        ]
        self.entries = {}
        grid_row = 1
        for lbl, key in fields:
            ctk.CTkLabel(self, text=lbl,
                         font=ctk.CTkFont(size=12)).grid(
                         row=grid_row, column=0, sticky="w", padx=20, pady=5)
            e = ctk.CTkEntry(self, width=260, corner_radius=8)
            e.grid(row=grid_row, column=1, sticky="ew", padx=20, pady=5)
            self.entries[key] = e
            grid_row += 1

        if self.cat:
            self.entries["name"].insert(0, self.cat.get("name", ""))
            self.entries["folder"].insert(0, self.cat.get("folder", ""))
            self.entries["extensions"].insert(0, ", ".join(self.cat.get("extensions", [])))
            self.entries["keywords"].insert(0, ", ".join(self.cat.get("keywords", [])))
            self.entries["rename_template"].insert(0, self.cat.get("rename_template") or "")

        ctk.CTkLabel(self,
                     text="Variabili: {date} = data   {type} = categoria\n"
                          "{sender} = mittente   {number} = numero documento",
                     font=ctk.CTkFont(size=11),
                     text_color="gray").grid(row=grid_row, column=1, sticky="w", padx=20, pady=(0, 4))

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.grid(row=grid_row+1, column=0, columnspan=2, pady=14)
        ctk.CTkButton(btn_row, text="Salva", width=120, corner_radius=8,
                      command=self._save).pack(side="left", padx=8)
        ctk.CTkButton(btn_row, text="Annulla", width=120, corner_radius=8,
                      fg_color="transparent", border_width=1,
                      text_color=("gray20", "gray80"),
                      command=self.destroy).pack(side="left", padx=8)

    def _save(self):
        name     = self.entries["name"].get().strip()
        folder   = self.entries["folder"].get().strip()
        exts     = [e.strip() for e in self.entries["extensions"].get().split(",") if e.strip()]
        keywords = [k.strip() for k in self.entries["keywords"].get().split(",") if k.strip()]
        template = self.entries["rename_template"].get().strip() or None
        if not name or not folder:
            messagebox.showerror("Errore", "Nome e cartella sono obbligatori.", parent=self)
            return
        new_cat = {"name": name, "folder": folder, "extensions": exts,
                   "keywords": keywords, "rename_template": template}
        if self.cat:
            self.rm.update_category(self.cat["name"], new_cat)
        else:
            self.rm.add_category(name, folder, keywords, exts, template)
        self.on_save()
        self.destroy()


# ─────────────────────────────────────────────
#  Finestra statistiche
# ─────────────────────────────────────────────

class SettingsWindow(ctk.CTkToplevel):
    def __init__(self, parent, rm, autostart_var, automonitor_var,
                 toggle_autostart_cb, save_config_cb, confidence_cb):
        super().__init__(parent)
        self.withdraw()
        self.title("Impostazioni — SortIt")
        self.geometry("480x420")
        self.resizable(False, False)
        self._rm = rm
        self._save_config_cb   = save_config_cb
        self._confidence_cb    = confidence_cb
        self._autostart_var    = autostart_var
        self._automonitor_var  = automonitor_var
        self._toggle_autostart = toggle_autostart_cb
        self._build()
        self.after(120, self.deiconify)
        self.after(150, self.lift)

    def _row(self, parent, title, subtitle=None):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=5)
        left = ctk.CTkFrame(row, fg_color="transparent")
        left.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(left, text=title,
                     font=ctk.CTkFont(size=13, weight="bold"),
                     anchor="w").pack(anchor="w")
        if subtitle:
            ctk.CTkLabel(left, text=subtitle,
                         font=ctk.CTkFont(size=11),
                         text_color="gray", anchor="w").pack(anchor="w")
        return row

    def _divider(self, parent):
        ctk.CTkFrame(parent, height=1,
                     fg_color=("gray80", "gray25")).pack(fill="x", pady=4)

    def _build(self):
        ctk.CTkLabel(self, text="⚙️  Impostazioni",
                     font=ctk.CTkFont(size=15, weight="bold")).pack(
                     anchor="w", padx=20, pady=(18, 8))

        card = ctk.CTkFrame(self, corner_radius=12)
        card.pack(fill="x", padx=20, pady=(0, 12))
        ci = ctk.CTkFrame(card, fg_color="transparent")
        ci.pack(fill="x", padx=16, pady=12)

        # Tema
        row = self._row(ci, "🌙  Tema scuro", "Alterna tra tema chiaro e scuro")
        self._theme_var = ctk.BooleanVar(
            value=ctk.get_appearance_mode().lower() == "dark")
        ctk.CTkSwitch(row, text="", variable=self._theme_var,
                      width=46, command=self._toggle_theme).pack(side="right")

        self._divider(ci)

        # Avvio con Windows
        row = self._row(ci, "🚀  Avvio con Windows",
                        "Apre SortIt automaticamente all'accensione del PC")
        ctk.CTkSwitch(row, text="", variable=self._autostart_var,
                      width=46, command=self._toggle_autostart).pack(side="right")

        self._divider(ci)

        # Avvio monitoraggio automatico
        row = self._row(ci, "▶  Avvia monitoraggio all'apertura",
                        "Inizia il monitoraggio automaticamente all'avvio dell'app")
        ctk.CTkSwitch(row, text="", variable=self._automonitor_var,
                      width=46, command=self._save_config_cb).pack(side="right")

        self._divider(ci)

        # Slider confidenza
        conf_header = ctk.CTkFrame(ci, fg_color="transparent")
        conf_header.pack(fill="x", pady=(4, 0))
        ctk.CTkLabel(conf_header, text="🎯  Soglia di confidenza",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
        self._conf_label = ctk.CTkLabel(
            conf_header,
            text=f"{int(self._rm.confidence_threshold * 100)}%",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=("#0077cc", "#00d2ff"))
        self._conf_label.pack(side="right")
        ctk.CTkLabel(ci, text="Percentuale minima di parole chiave per classificare un file",
                     font=ctk.CTkFont(size=11), text_color="gray",
                     anchor="w").pack(anchor="w")
        slider = ctk.CTkSlider(ci, from_=0, to=100, number_of_steps=20,
                               command=self._on_conf)
        slider.set(int(self._rm.confidence_threshold * 100))
        slider.pack(fill="x", pady=(6, 4))

        ctk.CTkButton(self, text="Chiudi", corner_radius=8,
                      command=self.destroy).pack(pady=(4, 16))

    def _toggle_theme(self):
        mode = "dark" if self._theme_var.get() else "light"
        ctk.set_appearance_mode(mode)
        self.master._theme = mode
        self.master._save_config()

    def _on_conf(self, value):
        self._confidence_cb(value)
        self._conf_label.configure(text=f"{round(value)}%")



# ─────────────────────────────────────────────
#  App principale
# ─────────────────────────────────────────────

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("SortIt")
        self.geometry("680x620")
        self.minsize(580, 520)

        if os.path.exists(ICON_FILE):
            try:
                self.iconbitmap(ICON_FILE)
            except Exception:
                pass
        # Fix icona taskbar Windows via ctypes
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("mic-rs.sortit.1.0")
        except Exception:
            pass
        png_icon = os.path.join(BASE_DIR, "icon.png")
        if os.path.exists(png_icon):
            try:
                from PIL import Image, ImageTk
                img = Image.open(png_icon).resize((64, 64), Image.LANCZOS)
                self._icon_photo = ImageTk.PhotoImage(img)
                self.iconphoto(True, self._icon_photo)
            except Exception:
                pass

        self.rm          = RulesManager()
        self.watcher     = None
        self._drive_poller: gds.DrivePoller = None
        self._drive_creds  = None

        (self.base_folder, saved_theme, saved_autostart, saved_automonitor,
         saved_mode, saved_drive_folder_id, saved_drive_folder_name,
         saved_poll_interval) = self._load_config()

        self._stats = {"total": 0, "unclassified": 0, "by_category": {}}

        self._theme = saved_theme if saved_theme in ("dark", "light") else "dark"
        ctk.set_appearance_mode(self._theme)

        self._autostart   = ctk.BooleanVar(value=saved_autostart)
        self._automonitor = ctk.BooleanVar(value=saved_automonitor)
        self._source_mode = ctk.StringVar(value=saved_mode)   # "local" | "drive"
        self._drive_folder_id   = saved_drive_folder_id
        self._drive_folder_name = saved_drive_folder_name
        self._poll_interval     = ctk.IntVar(value=saved_poll_interval)

        # Tenta di ripristinare il token Google salvato
        token_data = gds.load_token()
        if token_data:
            try:
                from google.oauth2.credentials import Credentials
                self._drive_creds = Credentials.from_authorized_user_info(token_data, gds.SCOPES)
            except Exception:
                self._drive_creds = None

        self._build_ui()
        self._update_status()

        # Avvia monitoraggio automatico se impostato
        if saved_automonitor and self.base_folder and os.path.isdir(self.base_folder):
            self.after(500, self._start)

    # ── Config ──────────────────────────────────────────────────────────────

    def _load_config(self):
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\SortIt")
            def rget(name, default):
                try: return winreg.QueryValueEx(key, name)[0]
                except: return default
            folder           = rget("last_folder", "")
            theme            = rget("theme", "dark")
            autostart        = rget("autostart", "0") == "1"
            automonitor      = rget("automonitor", "0") == "1"
            mode             = rget("source_mode", "local")   # "local" | "drive"
            drive_folder_id  = rget(gds.DRIVE_FOLDER_KEY, "")
            drive_folder_name = rget(gds.DRIVE_FOLDER_NAME_KEY, "")
            poll_interval    = int(rget(gds.POLL_INTERVAL_KEY, "5"))
            winreg.CloseKey(key)
            return folder, theme, autostart, automonitor, mode, drive_folder_id, drive_folder_name, poll_interval
        except Exception:
            return "", "dark", False, False, "local", "", "", 5

    def _save_config(self):
        try:
            import winreg
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\SortIt")
            winreg.SetValueEx(key, "last_folder",  0, winreg.REG_SZ, self.base_folder)
            winreg.SetValueEx(key, "theme",        0, winreg.REG_SZ, self._theme)
            winreg.SetValueEx(key, "autostart",    0, winreg.REG_SZ, "1" if self._autostart.get() else "0")
            winreg.SetValueEx(key, "automonitor",  0, winreg.REG_SZ, "1" if self._automonitor.get() else "0")
            winreg.SetValueEx(key, "source_mode",  0, winreg.REG_SZ, self._source_mode.get())
            winreg.SetValueEx(key, gds.DRIVE_FOLDER_KEY,       0, winreg.REG_SZ, self._drive_folder_id)
            winreg.SetValueEx(key, gds.DRIVE_FOLDER_NAME_KEY,  0, winreg.REG_SZ, self._drive_folder_name)
            winreg.SetValueEx(key, gds.POLL_INTERVAL_KEY,      0, winreg.REG_SZ, str(self._poll_interval.get()))
            winreg.CloseKey(key)
        except Exception:
            pass

    # ── UI ──────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Header
        header = ctk.CTkFrame(self, corner_radius=0, height=72)
        header.pack(fill="x")
        header.pack_propagate(False)

        ctk.CTkLabel(header, text="S",
                     font=ctk.CTkFont(family="Georgia", size=26, weight="bold"),
                     text_color=("#0077cc", "#00d2ff")).pack(side="left", padx=(20, 0), pady=12)

        title_col = ctk.CTkFrame(header, fg_color="transparent")
        title_col.pack(side="left", pady=(20, 0))

        lbl_ortit = ctk.CTkLabel(title_col, text="ortIt",
                     font=ctk.CTkFont(size=18, weight="bold"))
        lbl_ortit.pack(anchor="w", pady=(0, 0))

        ctk.CTkLabel(title_col, text="v2.0",
                     font=ctk.CTkFont(size=9),
                     text_color=("#0077cc", "#00d2ff"),
                     height=0,
                     fg_color="transparent").pack(anchor="w", padx=(19, 0))

        # "Created by Mic" a destra con link GitHub
        credit_frame = ctk.CTkFrame(header, fg_color="transparent")
        credit_frame.pack(side="right", padx=16)
        ctk.CTkLabel(credit_frame, text="Created by ",
                     font=ctk.CTkFont(size=11), text_color="gray").pack(side="left")
        mic_lbl = ctk.CTkLabel(credit_frame, text="Mic",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color=("#0077cc", "#00d2ff"), cursor="hand2")
        mic_lbl.pack(side="left")
        mic_lbl.bind("<Button-1>", lambda e: __import__("webbrowser").open("https://github.com/mic-rs"))

        # Corpo scrollabile
        body = ctk.CTkScrollableFrame(self, corner_radius=0, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=16, pady=(12, 0))

        # ── Sorgente (switch Local / Drive) ──
        self._section_label(body, "SORGENTE")
        src_card = ctk.CTkFrame(body, corner_radius=12)
        src_card.pack(fill="x", pady=(0, 14))

        si = ctk.CTkFrame(src_card, fg_color="transparent")
        si.pack(fill="x", padx=16, pady=14)

        # Label sinistra (Locale)
        self._lbl_local = ctk.CTkLabel(
            si, text="💻  Locale",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=("gray10", "gray90") if self._source_mode.get() == "local" else ("gray50", "gray50"))
        self._lbl_local.pack(side="left", padx=(0, 10))

        self._mode_switch = ctk.CTkSwitch(
            si, text="", variable=self._source_mode,
            onvalue="drive", offvalue="local",
            width=52, command=self._on_mode_change)
        self._mode_switch.pack(side="left")

        # Label destra (Drive)
        self._lbl_drive = ctk.CTkLabel(
            si, text="☁  Google Drive",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=("gray10", "gray90") if self._source_mode.get() == "drive" else ("gray50", "gray50"))
        self._lbl_drive.pack(side="left", padx=(10, 0))

        # ── Google Drive ──
        self._section_label(body, "GOOGLE DRIVE")
        self._drive_card = ctk.CTkFrame(body, corner_radius=12)
        self._drive_card.pack(fill="x", pady=(0, 14))

        di = ctk.CTkFrame(self._drive_card, fg_color="transparent")
        di.pack(fill="x", padx=16, pady=12)

        # Riga account
        acc_row = ctk.CTkFrame(di, fg_color="transparent")
        acc_row.pack(fill="x", pady=(0, 8))

        self._drive_account_label = ctk.CTkLabel(
            acc_row,
            text=self._get_drive_account_text(),
            font=ctk.CTkFont(size=12),
            text_color="gray")
        self._drive_account_label.pack(side="left", fill="x", expand=True)

        self._btn_connect = ctk.CTkButton(
            acc_row, text="🔗  Connetti Google", width=160, corner_radius=8,
            command=self._connect_google)
        self._btn_connect.pack(side="right")

        self._btn_disconnect = ctk.CTkButton(
            acc_row, text="Disconnetti", width=120, corner_radius=8,
            fg_color="transparent", border_width=1,
            text_color=("gray20", "gray80"),
            command=self._disconnect_google)
        self._btn_disconnect.pack(side="right", padx=(0, 6))

        # Riga cartella Drive
        folder_drive_row = ctk.CTkFrame(di, fg_color="transparent")
        folder_drive_row.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(folder_drive_row, text="Cartella Drive:",
                     font=ctk.CTkFont(size=12)).pack(side="left")

        self._drive_folder_label = ctk.CTkLabel(
            folder_drive_row,
            text=self._drive_folder_name or "Nessuna selezionata",
            font=ctk.CTkFont(size=12),
            text_color="gray")
        self._drive_folder_label.pack(side="left", padx=8, fill="x", expand=True)

        ctk.CTkButton(folder_drive_row, text="📂  Scegli", width=100,
                      corner_radius=8, command=self._pick_drive_folder).pack(side="right")

        # Riga intervallo polling
        poll_row = ctk.CTkFrame(di, fg_color="transparent")
        poll_row.pack(fill="x")

        ctk.CTkLabel(poll_row, text="Polling ogni",
                     font=ctk.CTkFont(size=12)).pack(side="left")

        self._poll_value_label = ctk.CTkLabel(
            poll_row,
            text=f"{self._poll_interval.get()} min",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=("gray10", "gray90"),
            width=50)
        self._poll_value_label.pack(side="right")

        self._poll_slider = ctk.CTkSlider(
            di, from_=1, to=60, number_of_steps=59,
            command=self._on_poll_slider)
        self._poll_slider.set(self._poll_interval.get())
        self._poll_slider.pack(fill="x", pady=(6, 4))

        # Se credentials.json mancante, avvisa
        if not gds.credentials_file_exists():
            ctk.CTkLabel(di,
                         text="⚠  credentials.json non trovato — necessario per il login Google.\n"
                              "Scaricalo da Google Cloud Console e mettilo accanto all'eseguibile.",
                         font=ctk.CTkFont(size=11),
                         text_color=("#d97706", "#fb923c"),
                         justify="left").pack(anchor="w", pady=(8, 0))

        self._refresh_drive_ui()

        # ── Cartella locale ──
        self._section_label(body, "CARTELLA LOCALE")
        folder_card = ctk.CTkFrame(body, corner_radius=12)
        folder_card.pack(fill="x", pady=(0, 14))

        fi = ctk.CTkFrame(folder_card, fg_color="transparent")
        fi.pack(fill="x", padx=16, pady=12)
        self.folder_label = ctk.CTkLabel(
            fi,
            text=self.base_folder or "Nessuna cartella selezionata",
            font=ctk.CTkFont(size=12),
            text_color="gray" if not self.base_folder else None,
            anchor="w")
        self.folder_label.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(fi, text="📂  Sfoglia", width=110, corner_radius=8,
                      command=self._select_folder).pack(side="right")

        # ── Opzioni ──
        self._section_label(body, "OPZIONI")
        opts_card = ctk.CTkFrame(body, corner_radius=12)
        opts_card.pack(fill="x", pady=(0, 14))

        oi = ctk.CTkFrame(opts_card, fg_color="transparent")
        oi.pack(fill="x", padx=16, pady=12)

        self.dry_run_var = ctk.BooleanVar(value=self.rm.dry_run)
        self.rename_var  = ctk.BooleanVar(value=self.rm.rename_enabled)

        dry_row = ctk.CTkFrame(oi, fg_color="transparent")
        dry_row.pack(fill="x", pady=4)
        ctk.CTkLabel(dry_row, text="🔍  Dry Run",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
        ctk.CTkLabel(dry_row, text="  — mostra anteprima senza spostare i file",
                     font=ctk.CTkFont(size=12), text_color="gray").pack(side="left")
        ctk.CTkSwitch(dry_row, text="", variable=self.dry_run_var,
                      width=46, command=self._toggle_dry_run).pack(side="right")

        ctk.CTkFrame(oi, height=1, fg_color=("gray80", "gray25")).pack(fill="x", pady=6)

        rename_row = ctk.CTkFrame(oi, fg_color="transparent")
        rename_row.pack(fill="x", pady=4)
        ctk.CTkLabel(rename_row, text="✎  Rinomina automatica",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
        ctk.CTkSwitch(rename_row, text="", variable=self.rename_var,
                      width=46, command=self._toggle_rename).pack(side="right")

        # ── Azioni ──
        self._section_label(body, "AZIONI")
        actions_card = ctk.CTkFrame(body, corner_radius=12)
        actions_card.pack(fill="x", pady=(0, 14))

        ai = ctk.CTkFrame(actions_card, fg_color="transparent")
        ai.pack(fill="x", padx=16, pady=12)

        self.btn_start = ctk.CTkButton(
            ai, text="▶  Avvia", width=110, corner_radius=8,
            fg_color="#15803d", hover_color="#166534", command=self._start)
        self.btn_start.pack(side="left", padx=(0, 8))

        self.btn_stop = ctk.CTkButton(
            ai, text="⏹  Ferma", width=100, corner_radius=8,
            fg_color="#b91c1c", hover_color="#991b1b",
            command=self._stop, state="disabled")
        self.btn_stop.pack(side="left", padx=(0, 8))

        ctk.CTkButton(ai, text="⚡  Ordina Esistenti", width=160,
                      corner_radius=8, command=self._sort_existing).pack(side="left", padx=(0, 8))

        right_btns = ctk.CTkFrame(ai, fg_color="transparent")
        right_btns.pack(side="right")
        ctk.CTkButton(right_btns, text="📋", width=40, corner_radius=8,
                      fg_color="transparent", border_width=1,
                      text_color=("gray20", "gray80"),
                      command=self._open_rules).pack(side="left", padx=(0, 6))
        ctk.CTkButton(right_btns, text="⚙️", width=40, corner_radius=8,
                      fg_color="transparent", border_width=1,
                      text_color=("gray20", "gray80"),
                      command=self._open_settings).pack(side="left")

        # Status
        self.status_label = ctk.CTkLabel(
            body, text="",
            font=ctk.CTkFont(size=12, weight="bold"), anchor="w")
        self.status_label.pack(anchor="w", pady=(0, 6))

        # ── Log ──
        self._section_label(body, "LOG OPERAZIONI")
        is_dark  = ctk.get_appearance_mode().lower() == "dark"
        bg_color = "#1e1e1e" if is_dark else "#f0f0f0"

        log_card = ctk.CTkFrame(body, corner_radius=12,
                                fg_color=("#e8e8e8", "#1e1e1e"))
        log_card.pack(fill="x", pady=(0, 6))

        log_fg     = "#e0e0e0" if is_dark else "#1a1a1a"
        log_sel_bg = "#3a7bd5" if is_dark else "#0077cc"
        log_cursor = "#00d2ff" if is_dark else "#0077cc"

        self.log_text = tk.Text(
            log_card, height=10, state="disabled",
            font=("Consolas", 9), relief="flat",
            padx=12, pady=10, wrap="word",
            bg=bg_color, fg=log_fg,
            insertbackground=log_cursor,
            selectbackground=log_sel_bg,
            borderwidth=0)
        self.log_text.pack(fill="x", padx=4, pady=4)

        info_fg = "#1a1a1a" if not is_dark else "#e0e0e0"
        self.log_text.tag_config("info",    foreground=info_fg)
        self.log_text.tag_config("success", foreground="#16a34a" if not is_dark else "#4ade80")
        self.log_text.tag_config("warning", foreground="#d97706" if not is_dark else "#fb923c")
        self.log_text.tag_config("error",   foreground="#dc2626" if not is_dark else "#f87171")
        self.log_text.tag_config("undo",    foreground="#2563eb" if not is_dark else "#60a5fa")

        ctk.CTkButton(body, text="🗑  Pulisci log", width=130,
                      corner_radius=8, fg_color="transparent", border_width=1,
                      text_color=("gray20", "gray80"),
                      command=self._clear_log).pack(anchor="e", pady=(4, 16))

    def _section_label(self, parent, text: str):
        ctk.CTkLabel(parent, text=text,
                     font=ctk.CTkFont(size=10, weight="bold"),
                     text_color="gray").pack(anchor="w", pady=(4, 4))

    # ── Handlers ────────────────────────────────────────────────────────────

    def _on_confidence_change(self, value):
        threshold = round(value) / 100
        self.rm.rules["confidence_threshold"] = threshold
        self.rm.save()

    def _toggle_autostart(self):
        import winreg
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        app_name = "SortIt"
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path,
                                 0, winreg.KEY_SET_VALUE)
            if self._autostart.get():
                exe_path = os.path.abspath(sys.executable if getattr(sys, "frozen", False)
                                           else __file__)
                winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, f'"{exe_path}"')
                self._log("Avvio automatico con Windows attivato.", "success")
            else:
                try:
                    winreg.DeleteValue(key, app_name)
                except FileNotFoundError:
                    pass
                self._log("Avvio automatico con Windows disattivato.", "info")
            winreg.CloseKey(key)
        except Exception as e:
            self._log(f"[ERRORE] Impossibile modificare avvio automatico: {e}", "error")
            self._autostart.set(not self._autostart.get())  # ripristina
        self._save_config()

    def _toggle_theme(self):
        self._theme = "light" if self._theme == "dark" else "dark"
        ctk.set_appearance_mode(self._theme)
        self._save_config()

    def _select_folder(self):
        folder = filedialog.askdirectory(title="Seleziona cartella da monitorare")
        if folder:
            self.base_folder = folder
            self.folder_label.configure(text=folder, text_color=("gray10", "gray90"))
            self._save_config()

    def _make_log_callback(self):
        def callback(msg: str, level: str = "info"):
            self._log(msg, level)
            if not self.rm.dry_run and "→" in msg \
                    and "[DRY RUN]" not in msg and "[ERRORE]" not in msg:
                self._stats["total"] += 1
                if "[NON CLASSIFICATO]" in msg:
                    self._stats["unclassified"] += 1
                else:
                    try:
                        cat_name = msg.split("→")[1].strip().split("/")[0].strip()
                        self._stats["by_category"][cat_name] = \
                            self._stats["by_category"].get(cat_name, 0) + 1
                    except Exception:
                        pass
        return callback

    def _start(self):
        mode = self._source_mode.get()

        if mode == "drive":
            if not self._drive_creds:
                messagebox.showwarning("Attenzione", "Prima connetti un account Google.")
                return
            if not self._drive_folder_id:
                messagebox.showwarning("Attenzione", "Prima seleziona una cartella Drive.")
                return
            if not self.base_folder or not os.path.isdir(self.base_folder):
                messagebox.showwarning("Attenzione", "Seleziona una cartella locale dove salvare i file scaricati.")
                return

            interval = self._get_poll_interval()
            self._poll_interval.set(interval)
            self._save_config()

            def on_file_downloaded(local_path):
                # SortIt classifica il file scaricato nella cartella locale
                if not self.watcher:
                    self.watcher = FolderWatcher(self.base_folder, self.rm,
                                                 self._make_log_callback())
                self.watcher._handler = self.watcher._make_handler() if not self.watcher._handler else self.watcher._handler
                self.watcher._handler._handle_file(local_path)

            self._drive_poller = gds.DrivePoller(
                creds=self._drive_creds,
                drive_folder_id=self._drive_folder_id,
                local_folder=self.base_folder,
                interval_minutes=interval,
                on_file_downloaded=on_file_downloaded,
                log_callback=self._make_log_callback()
            )
            self._drive_poller.start()
            self._log(f"☁ Polling Drive avviato (ogni {interval} min) — cartella: {self._drive_folder_name}", "success")

        else:
            # Modalità locale
            if not self.base_folder or not os.path.isdir(self.base_folder):
                messagebox.showwarning("Attenzione", "Prima seleziona una cartella locale valida.")
                return
            self.watcher = FolderWatcher(self.base_folder, self.rm,
                                         self._make_log_callback())
            self.watcher.start()
            self._log("Monitoraggio cartella locale avviato.", "success")

        self._update_status()

    def _stop(self):
        if self.watcher:
            self.watcher.stop()
        if self._drive_poller:
            self._drive_poller.stop()
            self._drive_poller = None
        self._update_status()
        self._log("Monitoraggio fermato.", "info")

    def _sort_existing(self):
        if not self.base_folder or not os.path.isdir(self.base_folder):
            messagebox.showwarning("Attenzione", "Prima seleziona una cartella valida.")
            return
        if not self.watcher:
            self.watcher = FolderWatcher(self.base_folder, self.rm,
                                         self._make_log_callback())
        self._log("Ordinamento file esistenti in corso…", "info")
        self.watcher.sort_existing()
        self._log("Ordinamento completato.", "success")

    def _open_rules(self):    RulesWindow(self, self.rm)
    def _open_settings(self): SettingsWindow(self, self.rm, self._autostart,
                                              self._automonitor,
                                              self._toggle_autostart,
                                              self._save_config,
                                              self._on_confidence_change)

    def _toggle_dry_run(self):
        self.rm.dry_run = self.dry_run_var.get()
        self.rm.save()

    def _toggle_rename(self):
        self.rm.rename_enabled = self.rename_var.get()
        self.rm.save()

    def _update_status(self):
        local_running = self.watcher and self.watcher.is_running
        drive_running = self._drive_poller and self._drive_poller.is_running
        running = local_running or drive_running
        if running:
            mode = self._source_mode.get()
            label = "●  Polling Drive attivo" if mode == "drive" else "●  Monitoraggio attivo"
            self.status_label.configure(text=label, text_color="#4ade80")
            self.btn_start.configure(state="disabled")
            self.btn_stop.configure(state="normal")
        else:
            self.status_label.configure(text="○  In attesa", text_color="gray")
            self.btn_start.configure(state="normal")
            self.btn_stop.configure(state="disabled")

    def _log(self, message: str, level: str = "info"):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{ts}] {message}\n", level)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _get_drive_account_text(self) -> str:
        if self._drive_creds:
            try:
                email = gds.get_user_email(self._drive_creds)
                if email:
                    return f"✅  Connesso come {email}"
            except Exception:
                pass
        return "Non connesso"

    def _refresh_drive_ui(self):
        connected = self._drive_creds is not None
        self._btn_connect.configure(state="disabled" if connected else "normal")
        self._btn_disconnect.configure(state="normal" if connected else "disabled")
        self._drive_account_label.configure(text=self._get_drive_account_text())

    def _on_mode_change(self):
        # Blocca switch se monitoraggio attivo
        local_running = self.watcher and self.watcher.is_running
        drive_running = self._drive_poller and self._drive_poller.is_running
        if local_running or drive_running:
            # Ripristina il valore precedente
            prev = "local" if self._source_mode.get() == "drive" else "drive"
            self._source_mode.set(prev)
            messagebox.showwarning(
                "Monitoraggio attivo",
                "Ferma il monitoraggio prima di cambiare sorgente.")
            return

        mode = self._source_mode.get()
        active_color   = ("gray10", "gray90")
        inactive_color = ("gray50", "gray50")
        if mode == "drive":
            self._lbl_local.configure(text_color=inactive_color)
            self._lbl_drive.configure(text_color=active_color)
        else:
            self._lbl_local.configure(text_color=active_color)
            self._lbl_drive.configure(text_color=inactive_color)
        self._save_config()

    def _connect_google(self):
        if not gds.credentials_file_exists():
            messagebox.showerror(
                "File mancante",
                "credentials.json non trovato.\n"
                "Scaricalo da Google Cloud Console e mettilo accanto all'eseguibile.")
            return
        self._log("Apertura browser per autenticazione Google…", "info")
        self._btn_connect.configure(state="disabled", text="Connessione…")

        def on_done(creds, email):
            self._drive_creds = creds
            self.after(0, lambda: (
                self._log(f"✅ Connesso come {email}", "success"),
                self._btn_connect.configure(text="🔗  Connetti Google"),
                self._refresh_drive_ui()
            ))

        def on_error(msg):
            self.after(0, lambda: (
                self._log(f"[ERRORE] Login Google fallito: {msg}", "error"),
                self._btn_connect.configure(
                    state="normal", text="🔗  Connetti Google")
            ))

        gds.authenticate(on_done, on_error)

    def _disconnect_google(self):
        if messagebox.askyesno("Disconnetti", "Disconnettere l'account Google?"):
            self._drive_creds = None
            gds.delete_token()
            self._refresh_drive_ui()
            self._log("Account Google disconnesso.", "info")

    def _pick_drive_folder(self):
        if not self._drive_creds:
            messagebox.showwarning("Attenzione", "Connetti prima un account Google.")
            return

        folders = gds.list_drive_folders(self._drive_creds)
        if not folders:
            messagebox.showinfo("Cartelle Drive",
                                "Nessuna cartella trovata nel tuo Drive.")
            return

        # Finestra di selezione
        picker = ctk.CTkToplevel(self)
        picker.withdraw()
        picker.title("Seleziona cartella Drive")
        picker.geometry("400x380")
        picker.resizable(False, False)

        ctk.CTkLabel(picker, text="Cartelle nel tuo Google Drive",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(
                     anchor="w", padx=20, pady=(16, 8))

        listbox = tk.Listbox(picker, font=("Segoe UI", 11),
                             selectmode="single", activestyle="none",
                             relief="flat", borderwidth=0)
        listbox.pack(fill="both", expand=True, padx=20, pady=(0, 8))

        is_dark = ctk.get_appearance_mode().lower() == "dark"
        listbox.configure(
            bg="#2b2b2b" if is_dark else "#f0f0f0",
            fg="#e0e0e0" if is_dark else "#1a1a1a",
            selectbackground="#3a7bd5", selectforeground="white")

        for f in folders:
            listbox.insert("end", f["name"])

        def confirm():
            sel = listbox.curselection()
            if not sel:
                return
            idx = sel[0]
            self._drive_folder_id   = folders[idx]["id"]
            self._drive_folder_name = folders[idx]["name"]
            self._drive_folder_label.configure(
                text=self._drive_folder_name,
                text_color=("gray10", "gray90"))
            self._save_config()
            picker.destroy()

        ctk.CTkButton(picker, text="Seleziona", corner_radius=8,
                      command=confirm).pack(pady=(0, 16))
        picker.after(120, picker.deiconify)
        picker.after(150, picker.lift)

    def _on_poll_slider(self, value):
        val = max(1, min(60, round(value)))
        self._poll_interval.set(val)
        self._poll_value_label.configure(text=f"{val} min")
        self._save_config()

    def _get_poll_interval(self) -> int:
        return max(1, min(60, self._poll_interval.get()))

    def on_close(self):
        if self.watcher:
            self.watcher.stop()
        if self._drive_poller:
            self._drive_poller.stop()
        self.destroy()


if __name__ == "__main__":
    import tkinter as _tk
    import math as _math
    import time as _time

    bg_color = "#0f0f1a"
    accent   = "#00d2ff"

    splash = _tk.Tk()
    splash.title("SortIt")
    splash.resizable(False, False)
    splash.configure(bg=bg_color)

    sw = splash.winfo_screenwidth()
    sh = splash.winfo_screenheight()
    splash.geometry(f"320x320+{(sw-320)//2}+{(sh-320)//2}")

    canvas = _tk.Canvas(splash, width=320, height=320,
                        bg=bg_color, highlightthickness=0)
    canvas.pack(fill="both", expand=True)

    canvas.create_text(160, 262, text="SortIt", fill="#e0e0e0",
                       font=("Segoe UI", 18, "bold"))
    canvas.create_text(160, 290, text="Ordinamento Intelligente",
                       fill="#6b7280", font=("Segoe UI", 10))

    # Punti path della S — bezier cubiche interpolate manualmente
    def s_points(cx, cy, size):
        import math
        s = size / 80

        # Punti di controllo bezier cubica per una S
        # Ogni segmento: P0, P1, P2, P3
        def bezier(p0, p1, p2, p3, steps=20):
            pts = []
            for i in range(steps + 1):
                t = i / steps
                x = ((1-t)**3 * p0[0] + 3*(1-t)**2*t * p1[0] +
                     3*(1-t)*t**2 * p2[0] + t**3 * p3[0])
                y = ((1-t)**3 * p0[1] + 3*(1-t)**2*t * p1[1] +
                     3*(1-t)*t**2 * p2[1] + t**3 * p3[1])
                pts.append((cx + x*s, cy + y*s))
            return pts

        pts = []
        # Arco superiore della S (da destra-alto verso sinistra-centro)
        pts += bezier(( 35, -60), ( 35, -90), (-35, -90), (-35, -45))
        # Raccordo centrale (da sinistra-alto verso destra-basso)
        pts += bezier((-35, -45), (-35,  -5), ( 35,   5), ( 35,  45))
        # Arco inferiore della S (da destra-centro verso sinistra-basso)
        pts += bezier(( 35,  45), ( 35,  90), (-35,  90), (-35,  60))

        return pts

    pts   = s_points(160, 140, 90)
    total = len(pts)
    delay_ms = max(1, int(1800 / total))

    # Disegna la S sincrono con update() invece di after ricorsivo
    for n in range(total - 1):
        x1, y1 = pts[n]
        x2, y2 = pts[n + 1]
        width = 2 + (n / total) * 4
        canvas.create_line(x1, y1, x2, y2,
                           fill=accent, width=width,
                           capstyle=_tk.ROUND)
        canvas.update()
        _time.sleep(delay_ms / 1000)

    _time.sleep(0.3)
    splash.destroy()

    app = App()
    app.protocol("WM_DELETE_WINDOW", app.on_close)
    app.mainloop()
