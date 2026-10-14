#!/usr/bin/env python3
"""DocuBuilder – Buttons, die hinterlegte Excel-Listen öffnen.

Läuft unter Windows und Linux, benötigt nur Python 3 (tkinter ist enthalten).
Die Einstellungen liegen in config.json neben dem Programm (portabel).
"""
import json
import os
import platform
import re
import subprocess
import sys
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont

APP_NAME = "DocuBuilder"
APP_VERSION = "1.1.0"

DEFAULT_SETTINGS = {"theme": "system", "max_items": 20}
MAX_ITEMS_LIMIT = 200

PALETTES = {
    "light": dict(bg="#f3f4f2", fg="#1d2a24", panel="#ffffff", accent="#1f6b4a",
                  accent_fg="#ffffff", line="#c9d0cb", mute="#66736c",
                  btn="#e9ede9", btn_hover="#dbe6df"),
    "dark": dict(bg="#171c19", fg="#e6ece8", panel="#212925", accent="#4fb283",
                 accent_fg="#0e1612", line="#35413b", mute="#93a299",
                 btn="#2a342f", btn_hover="#34423b"),
}


# ---------------------------------------------------------------- Konfiguration
def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def config_path():
    """config.json neben dem Programm; falls nicht beschreibbar im Benutzerordner."""
    local = os.path.join(app_dir(), "config.json")
    try:
        with open(local, "a", encoding="utf-8"):
            pass
        return local
    except OSError:
        fallback = os.path.join(os.path.expanduser("~"), ".docubuilder")
        os.makedirs(fallback, exist_ok=True)
        return os.path.join(fallback, "config.json")


CONFIG = config_path()


def clean_settings(raw):
    s = dict(DEFAULT_SETTINGS)
    if isinstance(raw, dict):
        if raw.get("theme") in ("system", "light", "dark"):
            s["theme"] = raw["theme"]
        try:
            s["max_items"] = min(MAX_ITEMS_LIMIT, max(1, int(raw.get("max_items"))))
        except (TypeError, ValueError):
            pass
    return s


def load_config():
    """Liest Einstellungen und Einträge (alte Dateien mit reiner Liste werden übernommen)."""
    try:
        with open(CONFIG, encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError):
        return dict(DEFAULT_SETTINGS), []
    raw_items = data if isinstance(data, list) else data.get("items", []) if isinstance(data, dict) else []
    raw_settings = data.get("settings") if isinstance(data, dict) else None
    items = [{"name": str(x.get("name", "")), "path": str(x.get("path", ""))}
             for x in raw_items if isinstance(x, dict)]
    return clean_settings(raw_settings), items


def save_config(settings, items):
    try:
        with open(CONFIG, "w", encoding="utf-8") as f:
            json.dump({"settings": settings, "items": items}, f, ensure_ascii=False, indent=2)
    except OSError as e:
        messagebox.showerror(APP_NAME, f"Einstellungen konnten nicht gespeichert werden:\n{e}")


# ---------------------------------------------------------------- Öffnen
def open_target(path):
    path = path.strip()
    if not path:
        messagebox.showinfo(APP_NAME, "Für diesen Button ist kein Pfad hinterlegt.")
        return
    if re.match(r"^https?://", path, re.I):
        webbrowser.open(path)
        return
    full = os.path.expandvars(os.path.expanduser(path))
    if not os.path.exists(full):
        messagebox.showerror(APP_NAME, f"Datei nicht gefunden:\n{full}")
        return
    try:
        if sys.platform.startswith("win"):
            os.startfile(full)  # öffnet mit dem Standardprogramm (Excel)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", full])
        else:
            subprocess.Popen(["xdg-open", full], stdout=subprocess.DEVNULL,
                             stderr=subprocess.DEVNULL)
    except Exception as e:  # noqa: BLE001
        messagebox.showerror(APP_NAME, f"Datei konnte nicht geöffnet werden:\n{e}")


# ---------------------------------------------------------------- Design
def detect_system_theme():
    """Liefert 'dark' oder 'light' passend zum Betriebssystem."""
    try:
        if sys.platform.startswith("win"):
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as k:
                return "light" if winreg.QueryValueEx(k, "AppsUseLightTheme")[0] else "dark"
        for key in ("color-scheme", "gtk-theme"):
            out = subprocess.run(["gsettings", "get", "org.gnome.desktop.interface", key],
                                 capture_output=True, text=True, timeout=2).stdout.lower()
            if "dark" in out:
                return "dark"
        if "dark" in os.environ.get("GTK_THEME", "").lower():
            return "dark"
    except Exception:  # noqa: BLE001
        pass
    return "light"


def set_dark_titlebar(win, dark):
    """Dunkle Titelleiste unter Windows 10/11 (rein kosmetisch, Fehler werden ignoriert)."""
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
        win.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(win.winfo_id())
        value = ctypes.c_int(1 if dark else 0)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(value), 4)
    except Exception:  # noqa: BLE001
        pass


def setup_style(root, p):
    s = ttk.Style(root)
    s.theme_use("clam")
    root.configure(bg=p["bg"])
    s.configure(".", background=p["bg"], foreground=p["fg"], fieldbackground=p["panel"],
                bordercolor=p["line"], lightcolor=p["line"], darkcolor=p["line"],
                troughcolor=p["bg"], focuscolor=p["accent"])
    s.configure("TFrame", background=p["bg"])
    s.configure("TLabel", background=p["bg"], foreground=p["fg"])
    s.configure("Mute.TLabel", foreground=p["mute"])
    s.configure("Title.TLabel", font=(None, 15, "bold"))
    s.configure("TLabelframe", background=p["bg"], bordercolor=p["line"])
    s.configure("TLabelframe.Label", background=p["bg"], foreground=p["fg"])
    s.configure("TButton", background=p["btn"], foreground=p["fg"], bordercolor=p["line"],
                padding=(10, 5))
    s.map("TButton", background=[("active", p["btn_hover"])])
    s.configure("Accent.TButton", background=p["accent"], foreground=p["accent_fg"])
    s.map("Accent.TButton", background=[("active", p["accent"])])
    s.configure("TEntry", fieldbackground=p["panel"], foreground=p["fg"], insertcolor=p["fg"])
    s.configure("TSpinbox", fieldbackground=p["panel"], foreground=p["fg"], background=p["btn"],
                arrowcolor=p["fg"], insertcolor=p["fg"])
    s.configure("TRadiobutton", background=p["bg"], foreground=p["fg"], indicatorcolor=p["panel"])
    s.map("TRadiobutton", background=[("active", p["bg"])],
          indicatorcolor=[("selected", p["accent"])])
    s.configure("TNotebook", background=p["bg"], bordercolor=p["line"])
    s.configure("TNotebook.Tab", background=p["btn"], foreground=p["fg"], padding=(14, 6))
    s.map("TNotebook.Tab", background=[("selected", p["panel"])])
    s.configure("Treeview", background=p["panel"], fieldbackground=p["panel"],
                foreground=p["fg"], rowheight=26)
    s.map("Treeview", background=[("selected", p["accent"])],
          foreground=[("selected", p["accent_fg"])])
    s.configure("Treeview.Heading", background=p["btn"], foreground=p["fg"])
    s.configure("TScrollbar", background=p["btn"], troughcolor=p["bg"], arrowcolor=p["fg"],
                bordercolor=p["line"])


# ---------------------------------------------------------------- Einstellungen
class Settings(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Einstellungen")
        self.geometry("760x540")
        self.minsize(620, 440)
        self.configure(bg=app.pal["bg"])
        self.transient(app)
        self.grab_set()
        set_dark_titlebar(self, app.is_dark)

        outer = ttk.Frame(self, padding=12)
        outer.pack(fill="both", expand=True)
        outer.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)
        nb = ttk.Notebook(outer)
        nb.grid(row=0, column=0, sticky="nsew")
        self.tab_items = ttk.Frame(nb, padding=12)
        self.tab_general = ttk.Frame(nb, padding=16)
        self.tab_info = ttk.Frame(nb, padding=16)
        nb.add(self.tab_items, text="Einträge")
        nb.add(self.tab_general, text="Allgemein")
        nb.add(self.tab_info, text="Info")
        ttk.Button(outer, text="Schließen", command=self.destroy).grid(
            row=1, column=0, sticky="e", pady=(12, 0))

        self.build_items_tab()
        self.build_general_tab()
        self.build_info_tab()

    # ---- Tab Einträge
    def build_items_tab(self):
        f = self.tab_items
        f.columnconfigure(0, weight=1)
        f.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(f, columns=("name", "path"), show="headings",
                                 selectmode="browse", height=8)
        self.tree.heading("name", text="Name")
        self.tree.heading("path", text="Pfad / Link")
        self.tree.column("name", width=170)
        self.tree.column("path", width=360)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(f, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        side = ttk.Frame(f)
        side.grid(row=0, column=2, sticky="n", padx=(10, 0))
        for text, cmd in (("Nach oben", lambda: self.move(-1)),
                          ("Nach unten", lambda: self.move(1)),
                          ("Löschen", self.delete)):
            ttk.Button(side, text=text, command=cmd, width=12).pack(pady=2)
        self.counter = ttk.Label(side, style="Mute.TLabel")
        self.counter.pack(pady=(14, 0))

        form = ttk.LabelFrame(f, text="Eintrag", padding=10)
        form.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(12, 0))
        form.columnconfigure(1, weight=1)
        self.name_var = tk.StringVar()
        self.path_var = tk.StringVar()
        ttk.Label(form, text="Name:").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Entry(form, textvariable=self.name_var).grid(row=0, column=1, columnspan=2,
                                                         sticky="ew", pady=3)
        ttk.Label(form, text="Pfad / Link:").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(form, textvariable=self.path_var).grid(row=1, column=1, sticky="ew", pady=3)
        ttk.Button(form, text="Durchsuchen …", command=self.browse).grid(row=1, column=2,
                                                                          padx=(6, 0))
        btns = ttk.Frame(form)
        btns.grid(row=2, column=0, columnspan=3, sticky="e", pady=(8, 0))
        ttk.Button(btns, text="Neuer Eintrag", command=self.new).pack(side="left", padx=4)
        ttk.Button(btns, text="Speichern", style="Accent.TButton",
                   command=self.save_entry).pack(side="left")
        self.refresh()

    def refresh(self, select=None):
        items = self.app.items
        self.tree.delete(*self.tree.get_children())
        for i, it in enumerate(items):
            self.tree.insert("", "end", iid=str(i), values=(it["name"], it["path"]))
        if select is not None and 0 <= select < len(items):
            self.tree.selection_set(str(select))
            self.tree.see(str(select))
        self.counter.configure(text=f"{len(items)} von {self.app.settings['max_items']}")

    def selected(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    def on_select(self, _event=None):
        i = self.selected()
        if i is not None:
            self.name_var.set(self.app.items[i]["name"])
            self.path_var.set(self.app.items[i]["path"])

    def new(self):
        self.tree.selection_remove(self.tree.selection())
        self.name_var.set("")
        self.path_var.set("")

    def browse(self):
        p = filedialog.askopenfilename(
            parent=self, title="Excel-Datei wählen",
            filetypes=[("Excel-Dateien", "*.xlsx *.xlsm *.xls *.xlsb *.csv *.ods"),
                       ("Alle Dateien", "*.*")])
        if p:
            self.path_var.set(os.path.normpath(p))
            if not self.name_var.get().strip():
                self.name_var.set(os.path.splitext(os.path.basename(p))[0])

    def save_entry(self):
        items = self.app.items
        name, path = self.name_var.get().strip(), self.path_var.get().strip()
        if not name or not path:
            messagebox.showinfo(APP_NAME, "Bitte Name und Pfad eintragen.", parent=self)
            return
        i = self.selected()
        if i is None:
            if len(items) >= self.app.settings["max_items"]:
                messagebox.showinfo(
                    APP_NAME, "Die maximale Anzahl an Einträgen ist erreicht.\n"
                              "Sie lässt sich unter „Allgemein“ erhöhen.", parent=self)
                return
            items.append({"name": name, "path": path})
            i = len(items) - 1
        else:
            items[i] = {"name": name, "path": path}
        self.commit(i)

    def delete(self):
        i = self.selected()
        if i is None:
            return
        if messagebox.askyesno(APP_NAME, f"„{self.app.items[i]['name']}“ löschen?", parent=self):
            del self.app.items[i]
            self.new()
            self.commit(None)

    def move(self, d):
        items = self.app.items
        i = self.selected()
        if i is None or not 0 <= i + d < len(items):
            return
        items[i], items[i + d] = items[i + d], items[i]
        self.commit(i + d)

    def commit(self, select):
        self.app.save()
        self.app.build(force=True)
        self.refresh(select)

    # ---- Tab Allgemein
    def build_general_tab(self):
        f = self.tab_general
        f.columnconfigure(0, weight=1)

        view = ttk.LabelFrame(f, text="Ansicht", padding=12)
        view.grid(row=0, column=0, sticky="ew")
        self.theme_var = tk.StringVar(value=self.app.settings["theme"])
        for value, text in (("system", "An das System anpassen"),
                            ("light", "Hell"), ("dark", "Dunkel")):
            ttk.Radiobutton(view, text=text, value=value, variable=self.theme_var,
                            command=self.apply_theme_choice).pack(anchor="w", pady=2)

        lim = ttk.LabelFrame(f, text="Auswahlmöglichkeiten", padding=12)
        lim.grid(row=1, column=0, sticky="ew", pady=(14, 0))
        row = ttk.Frame(lim)
        row.pack(anchor="w")
        ttk.Label(row, text="Maximale Anzahl:").pack(side="left")
        self.max_var = tk.StringVar(value=str(self.app.settings["max_items"]))
        self.spin = ttk.Spinbox(row, from_=1, to=MAX_ITEMS_LIMIT, width=6,
                                textvariable=self.max_var, command=self.apply_max)
        self.spin.pack(side="left", padx=8)
        self.spin.bind("<Return>", lambda e: self.apply_max())
        self.spin.bind("<FocusOut>", lambda e: self.apply_max())
        self.max_hint = ttk.Label(lim, style="Mute.TLabel",
                                  text=f"Erlaubt: 1 bis {MAX_ITEMS_LIMIT}. "
                                       "Nicht niedriger als die Zahl vorhandener Einträge.")
        self.max_hint.pack(anchor="w", pady=(6, 0))

        ttk.Button(f, text="Auf Standard zurücksetzen", command=self.reset_defaults).grid(
            row=2, column=0, sticky="w", pady=(18, 0))

    def apply_theme_choice(self):
        self.app.settings["theme"] = self.theme_var.get()
        self.app.save()
        self.app.apply_theme()
        self.app.build(force=True)

    def apply_max(self):
        n_items = len(self.app.items)
        try:
            value = int(self.max_var.get())
        except ValueError:
            value = self.app.settings["max_items"]
        low = max(1, n_items)
        clamped = min(MAX_ITEMS_LIMIT, max(low, value))
        if clamped != value:
            self.max_hint.configure(
                text=f"Wert angepasst auf {clamped} (erlaubt: {low} bis {MAX_ITEMS_LIMIT}).")
        self.max_var.set(str(clamped))
        if clamped != self.app.settings["max_items"]:
            self.app.settings["max_items"] = clamped
            self.app.save()
            self.refresh(self.selected())

    def reset_defaults(self):
        if not messagebox.askyesno(APP_NAME, "Allgemeine Einstellungen auf Standard zurücksetzen?\n"
                                             "Deine Einträge bleiben erhalten.", parent=self):
            return
        self.app.settings.update(DEFAULT_SETTINGS)
        self.app.settings["max_items"] = max(DEFAULT_SETTINGS["max_items"], len(self.app.items))
        self.theme_var.set(self.app.settings["theme"])
        self.max_var.set(str(self.app.settings["max_items"]))
        self.max_hint.configure(text=f"Erlaubt: 1 bis {MAX_ITEMS_LIMIT}. "
                                     "Nicht niedriger als die Zahl vorhandener Einträge.")
        self.app.save()
        self.app.apply_theme()
        self.app.build(force=True)
        self.refresh(self.selected())

    # ---- Tab Info
    def build_info_tab(self):
        f = self.tab_info
        f.columnconfigure(1, weight=1)
        ttk.Label(f, text=APP_NAME, style="Title.TLabel").grid(row=0, column=0, columnspan=2,
                                                               sticky="w")
        ttk.Label(f, text="Buttons, die hinterlegte Excel-Listen öffnen.",
                  style="Mute.TLabel").grid(row=1, column=0, columnspan=2, sticky="w",
                                            pady=(0, 14))
        rows = (("Version", APP_VERSION),
                ("Lizenz", "MIT"),
                ("Konfiguration", CONFIG),
                ("System", platform.platform()),
                ("Python", platform.python_version()))
        for i, (k, v) in enumerate(rows, start=2):
            ttk.Label(f, text=k + ":", style="Mute.TLabel").grid(row=i, column=0, sticky="nw",
                                                                 padx=(0, 12), pady=2)
            ttk.Label(f, text=v, wraplength=480, justify="left").grid(row=i, column=1,
                                                                       sticky="w", pady=2)
        ttk.Button(f, text="Konfigurationsordner öffnen",
                   command=lambda: open_target(os.path.dirname(CONFIG))).grid(
            row=len(rows) + 2, column=0, columnspan=2, sticky="w", pady=(18, 0))


# ---------------------------------------------------------------- Hauptfenster
class App(tk.Tk):
    TILE_W = 220

    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("720x420")
        self.minsize(360, 240)
        self.settings, self.items = load_config()
        if not self.items and self._is_empty_file():
            self.items = [{"name": "Beispiel: Kundenliste",
                           "path": os.path.join(os.path.expanduser("~"), "Kundenliste.xlsx")}]
            self.save()
        self.ui_font = tkfont.nametofont("TkDefaultFont").actual("family")
        self.pal = PALETTES["light"]
        self.is_dark = False
        self.apply_theme()

        top = ttk.Frame(self, padding=(14, 12))
        top.pack(fill="x")
        ttk.Label(top, text=APP_NAME, style="Title.TLabel").pack(side="left")
        ttk.Button(top, text="Einstellungen", command=lambda: Settings(self)).pack(side="right")

        self.area = ttk.Frame(self, padding=(14, 0, 14, 14))
        self.area.pack(fill="both", expand=True)
        self.cols = 0
        self.area.bind("<Configure>", lambda e: self.build())
        self.build(force=True)

    @staticmethod
    def _is_empty_file():
        try:
            return os.path.getsize(CONFIG) == 0
        except OSError:
            return True

    def save(self):
        save_config(self.settings, self.items)

    def apply_theme(self):
        mode = self.settings["theme"]
        effective = detect_system_theme() if mode == "system" else mode
        self.is_dark = effective == "dark"
        self.pal = PALETTES[effective]
        setup_style(self, self.pal)
        set_dark_titlebar(self, self.is_dark)
        for w in self.winfo_children():
            if isinstance(w, tk.Toplevel):
                w.configure(bg=self.pal["bg"])
                set_dark_titlebar(w, self.is_dark)

    def build(self, force=False):
        width = max(self.area.winfo_width(), 1)
        cols = max(1, width // (self.TILE_W + 10))
        if cols == self.cols and not force:
            return
        self.cols = cols
        p = self.pal
        for w in self.area.winfo_children():
            w.destroy()
        if not self.items:
            ttk.Label(self.area, text="Noch keine Einträge. Lege sie unter „Einstellungen“ an.",
                      style="Mute.TLabel").pack(pady=30)
            return
        for c in range(cols):
            self.area.columnconfigure(c, weight=1, uniform="t")
        for i, it in enumerate(self.items):
            b = tk.Button(self.area, text=it["name"] or "Ohne Namen", wraplength=self.TILE_W - 24,
                          height=3, relief="flat", bd=0, cursor="hand2",
                          font=(self.ui_font, 11), bg=p["panel"], fg=p["fg"],
                          activebackground=p["accent"], activeforeground=p["accent_fg"],
                          highlightthickness=1, highlightbackground=p["line"],
                          highlightcolor=p["accent"],
                          command=lambda path=it["path"]: open_target(path))
            b.bind("<Enter>", lambda e, w=b: w.configure(bg=p["btn_hover"]))
            b.bind("<Leave>", lambda e, w=b: w.configure(bg=p["panel"]))
            b.grid(row=i // cols, column=i % cols, padx=5, pady=5, sticky="nsew", ipady=6)


if __name__ == "__main__":
    App().mainloop()
