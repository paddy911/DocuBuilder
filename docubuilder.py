#!/usr/bin/env python3
"""DocuBuilder – Buttons, die hinterlegte Excel-Listen öffnen.

Läuft unter Windows und Linux, benötigt nur Python 3 (tkinter ist enthalten).
Anmeldung mit Benutzerkonten (Admin fest); Einstellungen verschlüsselt und versteckt (siehe README).
"""
import hashlib
import hmac
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter import font as tkfont

APP_NAME = "DocuBuilder"
# Version und Kanal werden von tools/version.py gepflegt (Nightly-Builds setzen sie nur beim Bauen).
APP_VERSION = "1.0.1.0"
APP_CHANNEL = "stable"
APP_TITLE = APP_NAME + (" Nightly" if APP_CHANNEL == "nightly" else "")

DEFAULT_SETTINGS = {"theme": "system", "max_items": 20, "files_dir": "", "setup_done": False,
                    "users": []}
ADMIN_NAME = "Admin"                       # festes Konto, kann nicht gelöscht werden
MIN_PW_LEN = 4
# Tastenkombinationen, die das Default-Fenster nach der ersten Einrichtung öffnen.
# Die zweite gilt als Ersatz, falls das System ALT+F11 selbst belegt.
DEFAULT_SHORTCUTS = ("<Alt-F11>", "<Control-Shift-F11>")
EXCEL_EXT = (".xlsx", ".xlsm", ".xls", ".xlsb", ".csv", ".ods")
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


WIN = sys.platform.startswith("win")
MAGIC = b"DBC1"
# Eingebetteter Schlüsselanteil: verhindert Lesen/Bearbeiten im Editor, ist aber kein Schutz
# gegen Entschlossene, die den Quelltext kennen (siehe README).
_SEED = b"DocuBuilder/settings/v1/7c2e9b41"
PORTABLE_FLAG = "portable.txt"
LINUX_SHARED = "/etc/docubuilder"
NO_RIGHTS = ("Nur der Admin oder Benutzer mit dem Recht „Darf ändern“ dürfen Einträge und "
             "Default-Einstellungen ändern.")
# Gemeinsame Einstellungen für alle Benutzer; die Ansicht ("theme") speichert jeder Benutzer selbst.
MACHINE_KEYS = ("max_items", "files_dir", "setup_done", "users")
STATE = {"legacy": None, "error": False}


def _set_win_attr(path, attr):
    try:
        import ctypes
        ctypes.windll.kernel32.SetFileAttributesW(str(path), attr)
    except Exception:  # noqa: BLE001
        pass


def protect_file(path, shared=False):
    """Datei schreibgeschützt und (unter Windows) versteckt ablegen."""
    try:
        if WIN:
            _set_win_attr(path, 0x1 | 0x2)   # READONLY | HIDDEN
        else:
            os.chmod(path, 0o666 if shared else 0o400)
    except OSError:
        pass


def unprotect_file(path):
    try:
        if os.path.exists(path):
            if WIN:
                _set_win_attr(path, 0x80)    # NORMAL
            else:
                os.chmod(path, 0o600)
    except OSError:
        pass


def is_portable():
    return os.path.exists(os.path.join(app_dir(), PORTABLE_FLAG))


def is_admin():
    try:
        if WIN:
            import ctypes
            return bool(ctypes.windll.shell32.IsUserAnAdmin())
        return os.geteuid() == 0
    except Exception:  # noqa: BLE001
        return False


def _run(cmd):
    try:
        flags = 0x08000000 if WIN else 0     # CREATE_NO_WINDOW
        return subprocess.run(cmd, capture_output=True, creationflags=flags).returncode == 0
    except OSError:
        return False


def prepare_machine_dir(d):
    """Legt den gemeinsamen Ordner an und setzt die Rechte: Alle Benutzer dürfen die
    Einstellungsdatei lesen und überschreiben (wer was ändern darf, prüft die Anmeldung im
    Programm), aber nicht löschen oder ersetzen. Sonst ließe sich durch Löschen eine neue
    Ersteinrichtung mit eigenem Admin erzwingen."""
    if is_portable() or (not WIN and d != LINUX_SHARED):
        return
    created = not os.path.isdir(d)
    try:
        os.makedirs(d, exist_ok=True)
    except OSError:
        return
    if not (created or is_admin()):
        return
    if WIN:
        _run(["icacls", d, "/setowner", "*S-1-5-32-544", "/T", "/C", "/Q"])   # nur als Admin
        # SIDs statt Namen, damit es auch auf deutschen Systemen (Administratoren/Benutzer) passt.
        _run(["icacls", d, "/inheritance:r", "/grant:r", "*S-1-5-18:(OI)(CI)F",
              "*S-1-5-32-544:(OI)(CI)F", "*S-1-5-32-545:(RX,WD)", "/Q"])
        _run(["icacls", d, "/grant", "*S-1-5-32-545:(OI)(IO)(R,W)", "/Q"])
        _run(["icacls", os.path.join(d, "*"), "/reset", "/C", "/Q"])
        _set_win_attr(d, 0x2)                # HIDDEN
    else:
        try:
            os.chmod(d, 0o1777)              # Sticky-Bit: nur der Eigentümer darf löschen
        except OSError:
            pass


def user_dir():
    """Versteckter Ordner im Benutzerprofil für die persönliche Ansicht."""
    try:
        if WIN:
            base = os.environ.get("APPDATA") or os.path.expanduser("~")
            d = os.path.join(base, "DocuBuilder")
        else:
            base = os.environ.get("XDG_CONFIG_HOME") or os.path.join(os.path.expanduser("~"), ".config")
            d = os.path.join(base, "docubuilder")
        os.makedirs(d, exist_ok=True)
        if WIN:
            _set_win_attr(d, 0x2)            # HIDDEN
        else:
            os.chmod(d, 0o700)
        return d
    except OSError:
        return app_dir()


def machine_dir():
    """Gemeinsamer Ordner für alle Benutzer. Mit portable.txt neben dem Programm: Programmordner."""
    if is_portable():
        return app_dir()
    if WIN:
        return os.path.join(os.environ.get("ProgramData") or r"C:\ProgramData", "DocuBuilder")
    # Unter Linux kann nur root /etc/docubuilder anlegen; bis dahin pro Benutzer.
    return LINUX_SHARED if os.path.isdir(LINUX_SHARED) or is_admin() else user_dir()


MACHINE_DIR = machine_dir()
CONFIG = os.path.join(MACHINE_DIR, ".docubuilder.dat")
PREFS = os.path.join(user_dir(), ".docubuilder-user.dat")
LEGACY_FILES = [os.path.join(user_dir(), ".docubuilder.dat"),     # 1.3: pro Benutzer
                os.path.join(app_dir(), "config.json"),
                os.path.join(os.path.expanduser("~"), ".docubuilder", "config.json")]


def _keys(salt):
    dk = hashlib.pbkdf2_hmac("sha256", _SEED, salt, 50_000, dklen=64)
    return dk[:32], dk[32:]


def _stream(key, nonce, n):
    out, counter = bytearray(), 0
    while len(out) < n:
        out += hmac.new(key, nonce + counter.to_bytes(8, "big"), hashlib.sha256).digest()
        counter += 1
    return bytes(out[:n])


def _xor(data, ks):
    return bytes(a ^ b for a, b in zip(data, ks))


def encrypt_blob(data):
    salt, nonce = os.urandom(16), os.urandom(16)
    ek, mk = _keys(salt)
    ct = _xor(data, _stream(ek, nonce, len(data)))
    mac = hmac.new(mk, MAGIC + salt + nonce + ct, hashlib.sha256).digest()
    return MAGIC + salt + nonce + mac + ct


def decrypt_blob(blob):
    if blob[:4] != MAGIC or len(blob) < 68:
        raise ValueError("Ungültiges Dateiformat")
    salt, nonce, mac, ct = blob[4:20], blob[20:36], blob[36:68], blob[68:]
    ek, mk = _keys(salt)
    if not hmac.compare_digest(mac, hmac.new(mk, MAGIC + salt + nonce + ct, hashlib.sha256).digest()):
        raise ValueError("Datei verändert oder beschädigt")
    return _xor(ct, _stream(ek, nonce, len(ct)))


def hash_password(pw):
    salt, iters = os.urandom(16), 200_000
    dk = hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), salt, iters)
    return f"pbkdf2${iters}${salt.hex()}${dk.hex()}"


def verify_password(pw, stored):
    try:
        _, iters, salt, dk = stored.split("$")
        calc = hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), bytes.fromhex(salt), int(iters))
        return hmac.compare_digest(calc.hex(), dk)
    except ValueError:
        return False


def clean_settings(raw):
    s = dict(DEFAULT_SETTINGS, users=[])
    if isinstance(raw, dict):
        if raw.get("theme") in ("system", "light", "dark"):
            s["theme"] = raw["theme"]
        try:
            s["max_items"] = min(MAX_ITEMS_LIMIT, max(1, int(raw.get("max_items"))))
        except (TypeError, ValueError):
            pass
        s["files_dir"] = str(raw.get("files_dir") or "")
        s["setup_done"] = bool(raw.get("setup_done", False))
        s["users"] = clean_users(raw.get("users"))
        h = raw.get("admin_hash")              # bis 1.3: Passwort des Default-Fensters
        if not s["users"] and isinstance(h, str) and h.startswith("pbkdf2$"):
            s["users"] = [new_user(ADMIN_NAME, h, admin=True)]
    return s


def new_user(name, pw_hash, admin=False, can_edit=False):
    return {"name": name, "hash": pw_hash, "admin": admin, "can_edit": admin or can_edit}


def clean_users(raw):
    """Gültige Konten; nur das Konto ADMIN_NAME ist Admin, Namen sind eindeutig."""
    users, seen = [], set()
    for u in raw if isinstance(raw, list) else []:
        if not isinstance(u, dict):
            continue
        name, h = str(u.get("name", "")).strip(), u.get("hash")
        if not name or name.lower() in seen or not (isinstance(h, str) and h.startswith("pbkdf2$")):
            continue
        seen.add(name.lower())
        is_adm = name.lower() == ADMIN_NAME.lower()
        users.append(new_user(ADMIN_NAME if is_adm else name, h, is_adm, bool(u.get("can_edit"))))
    users.sort(key=lambda u: not u["admin"])  # Admin immer zuerst
    return users


def find_user(users, name):
    return next((u for u in users if u["name"].lower() == name.strip().lower()), None)


def has_admin(settings):
    return any(u["admin"] for u in settings["users"])


def parse_config(data):
    raw_items = data if isinstance(data, list) else data.get("items", []) if isinstance(data, dict) else []
    raw_settings = data.get("settings") if isinstance(data, dict) else None
    items = [{"name": str(x.get("name", "")), "path": str(x.get("path", ""))}
             for x in raw_items if isinstance(x, dict)]
    settings = clean_settings(raw_settings)
    if not (isinstance(raw_settings, dict) and "setup_done" in raw_settings):
        settings["setup_done"] = bool(items)   # ältere Dateien: Einrichtung gilt als erledigt
    return settings, items


def read_store(path, allow_plain=False):
    with open(path, "rb") as f:
        raw = f.read()
    if allow_plain and raw[:4] != MAGIC:
        return json.loads(raw.decode("utf-8-sig"))
    return json.loads(decrypt_blob(raw).decode("utf-8"))


def write_store(path, obj, shared=False):
    data = encrypt_blob(json.dumps(obj, ensure_ascii=False).encode("utf-8"))
    if shared:
        # Gemeinsame Datei direkt überschreiben: Benutzer dürfen sie ändern, aber nicht
        # löschen oder ersetzen (siehe prepare_machine_dir).
        unprotect_file(path)
        try:
            with open(path, "r+b" if os.path.exists(path) else "xb") as f:
                f.write(data)
                f.truncate()
        finally:
            protect_file(path, shared)
        return
    tmp = path + ".tmp"
    with open(tmp, "wb") as f:
        f.write(data)
    unprotect_file(path)
    try:
        os.replace(tmp, path)
    finally:
        protect_file(path, shared)


def load_config():
    """Liest die gemeinsame Konfiguration. Ältere Dateien (1.3, config.json) werden übernommen."""
    if os.path.exists(CONFIG):
        try:
            return parse_config(read_store(CONFIG))
        except (OSError, ValueError):
            STATE["error"] = True
            return clean_settings(None), []
    for path in LEGACY_FILES:
        if os.path.isfile(path):
            try:
                data = read_store(path, allow_plain=True)
                STATE["legacy"] = path
                return parse_config(data)
            except (OSError, ValueError):
                continue
    return clean_settings(None), []


def load_prefs(settings):
    """Übernimmt die persönliche Ansicht des Benutzers in settings."""
    try:
        theme = read_store(PREFS).get("theme")
    except (OSError, ValueError, AttributeError):
        return
    if theme in ("system", "light", "dark"):
        settings["theme"] = theme


def save_config(settings, items):
    try:
        write_store(CONFIG, {"settings": {k: settings[k] for k in MACHINE_KEYS}, "items": items},
                    shared=True)
        return True
    except PermissionError:
        messagebox.showerror(APP_NAME, "Windows erlaubt das Speichern nicht:\n" + CONFIG +
                             "\n\nBitte DocuBuilder einmal als Administrator starten "
                             "(Rechtsklick → Als Administrator ausführen), damit die Rechte "
                             "gesetzt werden.")
    except OSError as e:
        messagebox.showerror(APP_NAME, f"Einstellungen konnten nicht gespeichert werden:\n{e}")
    return False


def _remove(path):
    """Datei oder Ordner löschen (Schreibschutz vorher aufheben). True, wenn danach weg."""
    if not os.path.exists(path):
        return True
    try:
        if os.path.isdir(path):
            for root, _, files in os.walk(path):
                for f in files:
                    unprotect_file(os.path.join(root, f))
            shutil.rmtree(path)
        else:
            unprotect_file(path)
            os.remove(path)
    except OSError:
        pass
    return not os.path.exists(path)


def _own_dir(d):
    """Nur eigene DocuBuilder-Ordner ganz löschen, nie z. B. den Programmordner."""
    return os.path.basename(os.path.normpath(d)).lower() == "docubuilder" and \
        os.path.normcase(os.path.abspath(d)) != os.path.normcase(app_dir())


def _run_elevated(params):
    """cmd.exe mit Administratorrechten (UAC-Abfrage) ausführen und warten. False bei Ablehnung."""
    import ctypes
    from ctypes import wintypes

    class ShellExecuteInfo(ctypes.Structure):
        _fields_ = [("cbSize", wintypes.DWORD), ("fMask", ctypes.c_ulong),
                    ("hwnd", wintypes.HWND), ("lpVerb", wintypes.LPCWSTR),
                    ("lpFile", wintypes.LPCWSTR), ("lpParameters", wintypes.LPCWSTR),
                    ("lpDirectory", wintypes.LPCWSTR), ("nShow", ctypes.c_int),
                    ("hInstApp", wintypes.HINSTANCE), ("lpIDList", ctypes.c_void_p),
                    ("lpClass", wintypes.LPCWSTR), ("hkeyClass", wintypes.HKEY),
                    ("dwHotKey", wintypes.DWORD), ("hIcon", wintypes.HANDLE),
                    ("hProcess", wintypes.HANDLE)]

    info = ShellExecuteInfo(cbSize=ctypes.sizeof(ShellExecuteInfo), fMask=0x40,  # NOCLOSEPROCESS
                            lpVerb="runas", lpFile="cmd.exe", lpParameters=params, nShow=0)
    if not ctypes.windll.shell32.ShellExecuteExW(ctypes.byref(info)):
        return False
    if info.hProcess:
        ctypes.windll.kernel32.WaitForSingleObject(info.hProcess, 60_000)
        ctypes.windll.kernel32.CloseHandle(info.hProcess)
    return True


def remove_all_data():
    """Löscht die gemeinsame Konfiguration (Einträge, Benutzer) und die eigene Ansicht.
    Liefert None bei Erfolg, sonst einen Fehlertext."""
    user = os.path.dirname(PREFS)
    for p in [user] if _own_dir(user) else [PREFS, PREFS + ".tmp"] + LEGACY_FILES:
        _remove(p)
    target = MACHINE_DIR if _own_dir(MACHINE_DIR) else CONFIG
    if _remove(target):
        return None
    if WIN:
        if os.path.isdir(target):
            cmd = f'/c attrib -R -H "{target}\\*" & rmdir /s /q "{target}"'
        else:
            cmd = f'/c attrib -R -H "{target}" & del /f /q "{target}"'
        if _run_elevated(cmd) and not os.path.exists(target):
            return None
        if os.path.exists(CONFIG):
            protect_file(CONFIG, shared=True)
        return ("Die gemeinsame Konfiguration konnte nicht gelöscht werden, weil keine "
                f"Administratorrechte erteilt wurden:\n{target}")
    return f"Bitte mit Administratorrechten löschen:\nsudo rm -rf {target}"


def save_prefs(settings):
    try:
        write_store(PREFS, {"theme": settings["theme"]})
    except OSError as e:
        messagebox.showerror(APP_NAME, f"Ansicht konnte nicht gespeichert werden:\n{e}")


# ---------------------------------------------------------------- Öffnen
def open_target(path, base=""):
    path = path.strip()
    if not path:
        messagebox.showinfo(APP_NAME, "Für diesen Button ist kein Pfad hinterlegt.")
        return
    if re.match(r"^https?://", path, re.I):
        webbrowser.open(path)
        return
    full = os.path.expandvars(os.path.expanduser(path))
    if base and not os.path.isabs(full):
        full = os.path.join(base, full)   # relative Pfade ab dem Dateiordner (Default)
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


def list_excel_files(folder):
    try:
        names = os.listdir(folder)
    except OSError:
        return []
    files = [os.path.join(folder, n) for n in names
             if os.path.splitext(n)[1].lower() in EXCEL_EXT and not n.startswith("~$")
             and os.path.isfile(os.path.join(folder, n))]
    return sorted(files, key=lambda f: os.path.basename(f).lower())


def fill_from_folder(items, folder, n):
    """Trägt Excel-Dateien aus dem Ordner in leere Plätze bzw. am Ende ein (max. n Einträge)."""
    existing = {os.path.normcase(os.path.normpath(i["path"])) for i in items if i["path"]}
    added = 0
    for path in list_excel_files(folder):
        key = os.path.normcase(os.path.normpath(path))
        if key in existing:
            continue
        entry = {"name": os.path.splitext(os.path.basename(path))[0], "path": path}
        slot = next((i for i, it in enumerate(items) if not it["path"]), None)
        if slot is not None:
            items[slot] = entry
        elif len(items) < n:
            items.append(entry)
        else:
            break
        existing.add(key)
        added += 1
    return added


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
    s.configure("TCheckbutton", background=p["bg"], foreground=p["fg"], indicatorcolor=p["panel"])
    s.map("TCheckbutton", background=[("active", p["bg"])],
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
        self.geometry("760x620")
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

        edit = []                                 # wird ohne Schreibrecht gesperrt
        side = ttk.Frame(f)
        side.grid(row=0, column=2, sticky="n", padx=(10, 0))
        for text, cmd in (("Nach oben", lambda: self.move(-1)),
                          ("Nach unten", lambda: self.move(1)),
                          ("Löschen", self.delete)):
            edit.append(ttk.Button(side, text=text, command=cmd, width=12))
            edit[-1].pack(pady=2)
        self.counter = ttk.Label(side, style="Mute.TLabel")
        self.counter.pack(pady=(14, 0))

        form = ttk.LabelFrame(f, text="Eintrag", padding=10)
        form.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(12, 0))
        form.columnconfigure(1, weight=1)
        self.name_var = tk.StringVar()
        self.path_var = tk.StringVar()
        ttk.Label(form, text="Name:").grid(row=0, column=0, sticky="w", pady=3)
        edit.append(ttk.Entry(form, textvariable=self.name_var))
        edit[-1].grid(row=0, column=1, columnspan=2, sticky="ew", pady=3)
        ttk.Label(form, text="Pfad / Link:").grid(row=1, column=0, sticky="w", pady=3)
        edit.append(ttk.Entry(form, textvariable=self.path_var))
        edit[-1].grid(row=1, column=1, sticky="ew", pady=3)
        edit.append(ttk.Button(form, text="Durchsuchen …", command=self.browse))
        edit[-1].grid(row=1, column=2, padx=(6, 0))
        btns = ttk.Frame(form)
        btns.grid(row=2, column=0, columnspan=3, sticky="e", pady=(8, 0))
        edit.append(ttk.Button(btns, text="Neuer Eintrag", command=self.new))
        edit[-1].pack(side="left", padx=4)
        edit.append(ttk.Button(btns, text="Speichern", style="Accent.TButton",
                               command=self.save_entry))
        edit[-1].pack(side="left")
        if not self.app.can_edit:
            for w in edit:
                w.state(["disabled"])
            ttk.Label(f, text=NO_RIGHTS, style="Mute.TLabel", wraplength=600,
                      justify="left").grid(row=2, column=0, columnspan=3, sticky="w",
                                           pady=(10, 0))
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
            initialdir=self.app.settings["files_dir"] or None,
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
                              "Sie ist in den Default-Einstellungen festgelegt.", parent=self)
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

        ttk.Button(f, text="Ansicht auf Standard zurücksetzen", command=self.reset_defaults).grid(
            row=1, column=0, sticky="w", pady=(18, 0))

    def apply_theme_choice(self):
        self.app.settings["theme"] = self.theme_var.get()
        save_prefs(self.app.settings)
        self.app.apply_theme()
        self.app.build(force=True)

    def reset_defaults(self):
        if not messagebox.askyesno(APP_NAME, "Ansicht auf Standard zurücksetzen?", parent=self):
            return
        self.app.settings["theme"] = DEFAULT_SETTINGS["theme"]
        self.theme_var.set(self.app.settings["theme"])
        save_prefs(self.app.settings)
        self.app.apply_theme()
        self.app.build(force=True)

    # ---- Tab Info
    def build_info_tab(self):
        f = self.tab_info
        f.columnconfigure(1, weight=1)
        ttk.Label(f, text=APP_NAME, style="Title.TLabel").grid(row=0, column=0, columnspan=2,
                                                               sticky="w")
        ttk.Label(f, text="Buttons, die hinterlegte Excel-Listen öffnen.",
                  style="Mute.TLabel").grid(row=1, column=0, columnspan=2, sticky="w",
                                            pady=(0, 14))
        rows = (("Version", APP_VERSION + (" (Nightly)" if APP_CHANNEL == "nightly" else "")),
                ("Lizenz", "MIT"),
                ("Konfiguration", CONFIG),
                ("Angemeldet als", self.app.user["name"] + (
                    " (Admin)" if self.app.user["admin"] else
                    " (darf ändern)" if self.app.can_edit else " (nur ansehen)")),
                ("System", platform.platform()),
                ("Python", platform.python_version()))
        for i, (k, v) in enumerate(rows, start=2):
            ttk.Label(f, text=k + ":", style="Mute.TLabel").grid(row=i, column=0, sticky="nw",
                                                                 padx=(0, 12), pady=2)
            ttk.Label(f, text=v, wraplength=480, justify="left").grid(row=i, column=1,
                                                                       sticky="w", pady=2)
        if self.app.user["admin"]:
            box = ttk.LabelFrame(f, text="Daten entfernen", padding=10)
            box.grid(row=len(rows) + 2, column=0, columnspan=2, sticky="ew", pady=(18, 0))
            ttk.Label(box, style="Mute.TLabel", wraplength=560, justify="left",
                      text="Löscht alle Einträge, Benutzer und Passwörter dieses Rechners sowie "
                           "deine Ansicht. Danach wird DocuBuilder beendet und kann gelöscht "
                           "werden. Windows fragt nach Administratorrechten.").pack(anchor="w")
            ttk.Button(box, text="DocuBuilder-Daten entfernen …",
                       command=self.remove_data).pack(anchor="w", pady=(8, 0))

    def remove_data(self):
        if not messagebox.askyesno(
                APP_NAME, "Wirklich ALLE DocuBuilder-Daten auf diesem Rechner löschen?\n\n"
                          "Einträge, Benutzer und Passwörter gehen verloren. Das kann nicht "
                          "rückgängig gemacht werden.", icon="warning", default="no", parent=self):
            return
        pw = simpledialog.askstring(APP_NAME, f"Zur Bestätigung das {ADMIN_NAME}-Passwort eingeben:",
                                    show="•", parent=self)
        if pw is None:
            return
        admin = next((u for u in self.app.settings["users"] if u["admin"]), None)
        if not (admin and verify_password(pw, admin["hash"])):
            messagebox.showerror(APP_NAME, "Falsches Passwort. Es wurde nichts gelöscht.", parent=self)
            return
        err = remove_all_data()
        if err:
            messagebox.showerror(APP_NAME, err, parent=self)
            return
        messagebox.showinfo(
            APP_NAME, "Alle DocuBuilder-Daten wurden entfernt.\n\nDocuBuilder wird jetzt beendet. "
                      "Danach kannst du die Programmdatei löschen.\n\nAndere Windows-Benutzer "
                      "können ihre Ansicht unter %APPDATA%\\DocuBuilder selbst löschen.", parent=self)
        self.app.destroy()


# ---------------------------------------------------------------- Default
def check_password(p1, p2):
    """Fehlertext oder None."""
    if p1 != p2:
        return "Die Passwörter stimmen nicht überein."
    if len(p1) < MIN_PW_LEN:
        return f"Das Passwort braucht mindestens {MIN_PW_LEN} Zeichen."
    return None


class UserDialog(tk.Toplevel):
    """Benutzer anlegen (user=None) oder Passwort eines Benutzers neu setzen."""

    def __init__(self, parent, app, taken, user=None):
        super().__init__(parent)
        self.result, self.taken, self.user = None, taken, user
        self.title("Passwort setzen" if user else "Benutzer anlegen")
        self.configure(bg=app.pal["bg"])
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        set_dark_titlebar(self, app.is_dark)
        f = ttk.Frame(self, padding=18)
        f.pack(fill="both", expand=True)
        f.columnconfigure(1, weight=1)
        self.name_var, self.pw1, self.pw2 = tk.StringVar(), tk.StringVar(), tk.StringVar()
        self.edit_var = tk.BooleanVar(value=False)
        ttk.Label(f, text="Benutzername:").grid(row=0, column=0, sticky="w", pady=4)
        if user:
            ttk.Label(f, text=user["name"]).grid(row=0, column=1, sticky="w", padx=(8, 0))
        else:
            e = ttk.Entry(f, textvariable=self.name_var, width=30)
            e.grid(row=0, column=1, sticky="ew", pady=4, padx=(8, 0))
            e.focus_set()
        ttk.Label(f, text="Passwort:").grid(row=1, column=0, sticky="w", pady=4)
        pw = ttk.Entry(f, textvariable=self.pw1, show="•", width=30)
        pw.grid(row=1, column=1, sticky="ew", pady=4, padx=(8, 0))
        if user:
            pw.focus_set()
        ttk.Label(f, text="Wiederholen:").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(f, textvariable=self.pw2, show="•", width=30).grid(row=2, column=1, sticky="ew",
                                                                     pady=4, padx=(8, 0))
        if not user:
            ttk.Checkbutton(f, text="Darf ändern (Einträge und Default-Fenster)",
                            variable=self.edit_var).grid(row=3, column=0, columnspan=2,
                                                         sticky="w", pady=(6, 0))
        self.msg = ttk.Label(f, wraplength=320, justify="left")
        self.msg.grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))
        row = ttk.Frame(f)
        row.grid(row=5, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(row, text="Abbrechen", command=self.destroy).pack(side="left", padx=6)
        ttk.Button(row, text="OK", style="Accent.TButton", command=self.ok).pack(side="left")
        self.bind("<Return>", lambda ev: self.ok())
        self.bind("<Escape>", lambda ev: self.destroy())

    def ok(self):
        name = self.user["name"] if self.user else self.name_var.get().strip()
        if not self.user:
            if not name or len(name) > 40:
                return self.msg.configure(text="Bitte einen Benutzernamen (max. 40 Zeichen) eingeben.")
            if name.lower() in self.taken:
                return self.msg.configure(text="Diesen Benutzernamen gibt es schon.")
        err = check_password(self.pw1.get(), self.pw2.get())
        if err:
            return self.msg.configure(text=err)
        self.result = new_user(name, hash_password(self.pw1.get()),
                               can_edit=self.user["can_edit"] if self.user else self.edit_var.get())
        self.destroy()


class DefaultDialog(tk.Toplevel):
    """Anzahl der Buttons, Dateiordner, Benutzer. Erststart automatisch, danach per ALT+F11."""

    def __init__(self, app, first_run=False):
        super().__init__(app)
        self.app = app
        self.first_run = first_run
        self.manage_users = first_run or app.user["admin"]
        self.title("Default – Erste Einrichtung" if first_run else "Default")
        self.geometry("680x760" if self.manage_users else "640x480")
        self.minsize(580, 640 if self.manage_users else 420)
        self.configure(bg=app.pal["bg"])
        self.transient(app)
        self.grab_set()
        set_dark_titlebar(self, app.is_dark)
        self.protocol("WM_DELETE_WINDOW", self.cancel)

        f = ttk.Frame(self, padding=18)
        f.pack(fill="both", expand=True)
        f.columnconfigure(1, weight=1)
        ttk.Label(f, text="Default", style="Title.TLabel").grid(row=0, column=0, columnspan=3,
                                                                sticky="w")
        ttk.Label(f, style="Mute.TLabel", wraplength=580, justify="left",
                  text="Lege fest, wie viele Buttons es gibt und in welchem Ordner deine Dateien "
                       "liegen. Dieses Fenster erscheint nur beim ersten Start. Später öffnest du "
                       "es mit ALT+F11 (alternativ STRG+UMSCHALT+F11).").grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(4, 16))

        ttk.Label(f, text="Anzahl der Buttons:").grid(row=2, column=0, sticky="w", pady=6)
        self.count_var = tk.StringVar(value=str(app.settings["max_items"]))
        ttk.Spinbox(f, from_=1, to=MAX_ITEMS_LIMIT, width=6, textvariable=self.count_var).grid(
            row=2, column=1, sticky="w", pady=6)

        ttk.Label(f, text="Ordner mit den Dateien:").grid(row=3, column=0, sticky="w", pady=6)
        self.dir_var = tk.StringVar(value=app.settings["files_dir"])
        ttk.Entry(f, textvariable=self.dir_var).grid(row=3, column=1, sticky="ew", pady=6)
        ttk.Button(f, text="Durchsuchen …", command=self.browse).grid(row=3, column=2,
                                                                       padx=(6, 0))

        self.import_var = tk.BooleanVar(value=first_run)
        ttk.Checkbutton(f, text="Excel-Dateien aus dem Ordner als Buttons übernehmen",
                        variable=self.import_var).grid(row=4, column=1, columnspan=2,
                                                       sticky="w", pady=(2, 0))

        self.pw1, self.pw2 = tk.StringVar(), tk.StringVar()
        admin = next((u for u in app.settings["users"] if u["admin"]), None)
        self.admin_hash = admin["hash"] if admin else ""
        self.users = [dict(u) for u in app.settings["users"] if not u["admin"]]
        if self.manage_users:
            self.build_users(f)
        else:
            ttk.Label(f, style="Mute.TLabel", text="Benutzer verwaltet nur der Admin.").grid(
                row=5, column=0, columnspan=3, sticky="w", pady=(16, 0))

        ttk.Label(f, style="Mute.TLabel", wraplength=580, justify="left",
                  text="Nicht belegte Buttons heißen „Auswahl 1, 2, …“ und können unter "
                       "Einstellungen → Einträge zugeordnet werden. Relative Pfade werden ab dem "
                       "Ordner gesucht.").grid(row=6, column=0, columnspan=3, sticky="w",
                                               pady=(14, 0))
        self.msg = ttk.Label(f, wraplength=580, justify="left")
        self.msg.grid(row=7, column=0, columnspan=3, sticky="w", pady=(10, 0))

        btns = ttk.Frame(f)
        btns.grid(row=8, column=0, columnspan=3, sticky="se", pady=(16, 0))
        f.rowconfigure(8, weight=1)
        ttk.Button(btns, text="Abbrechen", command=self.cancel).pack(side="left", padx=6)
        ttk.Button(btns, text="Übernehmen", style="Accent.TButton",
                   command=self.apply).pack(side="left")

    def build_users(self, f):
        box = ttk.LabelFrame(f, text="Benutzer", padding=10)
        box.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(16, 0))
        f.rowconfigure(5, weight=1)
        box.columnconfigure(1, weight=1)
        box.rowconfigure(3, weight=1)
        ttk.Label(box, style="Mute.TLabel", wraplength=560, justify="left", text=(
            f"Das Konto „{ADMIN_NAME}“ besteht immer und darf alles. Bitte ein Passwort festlegen."
            if self.first_run else
            f"Passwort für „{ADMIN_NAME}“ ändern (leer lassen = unverändert):")).grid(
            row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(box, text=f"{ADMIN_NAME}-Passwort:").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(box, textvariable=self.pw1, show="•").grid(row=1, column=1, columnspan=2,
                                                             sticky="ew", pady=3, padx=(8, 0))
        ttk.Label(box, text="Wiederholen:").grid(row=2, column=0, sticky="w", pady=3)
        ttk.Entry(box, textvariable=self.pw2, show="•").grid(row=2, column=1, columnspan=2,
                                                             sticky="ew", pady=3, padx=(8, 0))

        self.utree = ttk.Treeview(box, columns=("name", "edit"), show="headings",
                                  selectmode="browse", height=5)
        self.utree.heading("name", text="Benutzer")
        self.utree.heading("edit", text="Darf ändern")
        self.utree.column("name", width=220)
        self.utree.column("edit", width=120, anchor="center")
        self.utree.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(10, 0))
        side = ttk.Frame(box)
        side.grid(row=3, column=2, sticky="n", padx=(10, 0), pady=(10, 0))
        for text, cmd in (("Hinzufügen …", self.add_user),
                          ("Passwort …", self.reset_password),
                          ("Recht an/aus", self.toggle_edit),
                          ("Löschen", self.delete_user)):
            ttk.Button(side, text=text, command=cmd, width=14).pack(pady=2)
        self.refresh_users()

    def refresh_users(self, select=None):
        self.utree.delete(*self.utree.get_children())
        self.utree.insert("", "end", iid="admin", values=(f"{ADMIN_NAME} (fest)", "ja"))
        for i, u in enumerate(self.users):
            self.utree.insert("", "end", iid=str(i),
                              values=(u["name"], "ja" if u["can_edit"] else "nein"))
        if select is not None:
            self.utree.selection_set(select)

    def selected_user(self):
        sel = self.utree.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "Bitte einen Benutzer auswählen.", parent=self)
            return None
        if sel[0] == "admin":
            messagebox.showinfo(APP_NAME, f"Das Konto „{ADMIN_NAME}“ kann nicht gelöscht oder "
                                          "eingeschränkt werden. Sein Passwort änderst du oben.",
                                parent=self)
            return None
        return int(sel[0])

    def add_user(self):
        taken = {ADMIN_NAME.lower()} | {u["name"].lower() for u in self.users}
        dlg = UserDialog(self, self.app, taken)
        self.wait_window(dlg)
        if dlg.result:
            self.users.append(dlg.result)
            self.refresh_users(str(len(self.users) - 1))

    def reset_password(self):
        i = self.selected_user()
        if i is None:
            return
        dlg = UserDialog(self, self.app, set(), user=self.users[i])
        self.wait_window(dlg)
        if dlg.result:
            self.users[i] = dlg.result
            self.refresh_users(str(i))

    def toggle_edit(self):
        i = self.selected_user()
        if i is not None:
            self.users[i]["can_edit"] = not self.users[i]["can_edit"]
            self.refresh_users(str(i))

    def delete_user(self):
        i = self.selected_user()
        if i is not None and messagebox.askyesno(
                APP_NAME, f"Benutzer „{self.users[i]['name']}“ löschen?", parent=self):
            del self.users[i]
            self.refresh_users()

    def browse(self):
        d = filedialog.askdirectory(parent=self, title="Ordner mit den Dateien wählen",
                                    initialdir=self.dir_var.get() or None)
        if d:
            self.dir_var.set(os.path.normpath(d))

    def cancel(self):
        if self.first_run and not messagebox.askyesno(
                APP_NAME, "Ohne Einrichtung kann DocuBuilder nicht benutzt werden.\n"
                          "Programm beenden?", parent=self):
            return
        self.grab_release()
        self.destroy()
        if self.first_run:
            self.app.destroy()

    def apply(self):
        try:
            n = int(self.count_var.get())
        except ValueError:
            n = 0
        if not 1 <= n <= MAX_ITEMS_LIMIT:
            self.msg.configure(text=f"Bitte eine Zahl von 1 bis {MAX_ITEMS_LIMIT} eingeben.")
            return
        folder = self.dir_var.get().strip()
        if folder:
            folder = os.path.normpath(os.path.expandvars(os.path.expanduser(folder)))
            if not os.path.isdir(folder):
                self.msg.configure(text="Der Ordner wurde nicht gefunden.")
                return
        elif self.import_var.get():
            self.msg.configure(text="Zum Einlesen bitte einen Ordner angeben.")
            return

        users = self.app.settings["users"]
        if self.manage_users:
            p1, p2 = self.pw1.get(), self.pw2.get()
            admin_hash = self.admin_hash
            if p1 or p2 or not admin_hash:
                err = check_password(p1, p2)
                if err:
                    self.msg.configure(text=f"{ADMIN_NAME}-Passwort: {err}")
                    return
                admin_hash = hash_password(p1)
            users = [new_user(ADMIN_NAME, admin_hash, admin=True)] + self.users

        items = self.app.items
        if len(items) > n:
            if not messagebox.askyesno(
                    APP_NAME, f"Es gibt {len(items)} Einträge, erlaubt sind {n}.\n"
                              "Die überzähligen Einträge am Ende werden entfernt. Fortfahren?",
                    parent=self):
                return
            del items[n:]
        imported = fill_from_folder(items, folder, n) if self.import_var.get() else 0
        total = len(list_excel_files(folder)) if self.import_var.get() else 0
        while len(items) < n:
            items.append({"name": f"Auswahl {len(items) + 1}", "path": ""})

        self.app.settings.update(max_items=n, files_dir=folder, setup_done=True, users=users)
        if self.first_run:
            self.app.user = users[0]             # nach der Einrichtung als Admin angemeldet
        if not self.app.save():
            if self.first_run:
                self.app.user = None
            return
        self.grab_release()
        self.destroy()
        if self.first_run:
            self.app.show_main()
        else:
            self.app.build(force=True)
        if self.import_var.get():
            messagebox.showinfo(APP_NAME, f"{imported} Datei(en) übernommen "
                                          f"(im Ordner liegen {total}).", parent=self.app)


# ---------------------------------------------------------------- Hauptfenster
class App(tk.Tk):
    TILE_W = 220

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("720x420")
        self.minsize(360, 240)
        prepare_machine_dir(MACHINE_DIR)
        self.settings, self.items = load_config()
        load_prefs(self.settings)
        self.user = None                         # angemeldeter Benutzer (dict aus settings["users"])
        self.settings_win = None
        self.default_win = None
        self.ui_font = tkfont.nametofont("TkDefaultFont").actual("family")
        self.pal = PALETTES["light"]
        self.is_dark = False
        self.apply_theme()

        top = ttk.Frame(self, padding=(14, 12))
        top.pack(fill="x")
        ttk.Label(top, text=APP_NAME, style="Title.TLabel").pack(side="left")
        self.session = ttk.Frame(top)            # nur sichtbar, wenn jemand angemeldet ist
        ttk.Button(self.session, text="Einstellungen", command=self.open_settings).pack(side="right")
        ttk.Button(self.session, text="Abmelden", command=self.logout).pack(side="right", padx=6)
        self.who = ttk.Label(self.session, style="Mute.TLabel")
        self.who.pack(side="right", padx=6)

        self.area = ttk.Frame(self, padding=(14, 0, 14, 14))
        self.cols = 0
        self.area.bind("<Configure>", lambda e: self.build())
        self.login = self.build_login()

        for key in DEFAULT_SHORTCUTS:
            self.bind_all(key, lambda e: self.open_default() or "break")
        self.after(200, self.startup)

    @property
    def can_edit(self):
        return bool(self.user and self.user["can_edit"])

    @staticmethod
    def _alive(win):
        return win is not None and win.winfo_exists()

    # ---- Anmeldung
    def build_login(self):
        outer = ttk.Frame(self, padding=14)
        f = ttk.Frame(outer)
        f.place(relx=0.5, rely=0.4, anchor="center")
        ttk.Label(f, text="Anmelden", style="Title.TLabel").grid(row=0, column=0, columnspan=2,
                                                                 sticky="w", pady=(0, 10))
        self.login_name, self.login_pw = tk.StringVar(), tk.StringVar()
        ttk.Label(f, text="Benutzer:").grid(row=1, column=0, sticky="w", pady=4)
        self.login_entry = ttk.Entry(f, textvariable=self.login_name, width=28)
        self.login_entry.grid(row=1, column=1, pady=4, padx=(8, 0))
        ttk.Label(f, text="Passwort:").grid(row=2, column=0, sticky="w", pady=4)
        pw = ttk.Entry(f, textvariable=self.login_pw, show="•", width=28)
        pw.grid(row=2, column=1, pady=4, padx=(8, 0))
        self.login_btn = ttk.Button(f, text="Anmelden", style="Accent.TButton",
                                    command=self.do_login)
        self.login_btn.grid(row=3, column=1, sticky="e", pady=(10, 0))
        self.login_msg = ttk.Label(f, style="Mute.TLabel", wraplength=320, justify="left")
        self.login_msg.grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 0))
        for w in (self.login_entry, pw):
            w.bind("<Return>", lambda e: self.do_login())
        return outer

    def show_login(self, msg=""):
        for win in (self.settings_win, self.default_win):
            if self._alive(win):
                win.destroy()
        self.user = None
        self.session.pack_forget()
        self.area.pack_forget()
        self.login.pack(fill="both", expand=True)
        self.login_pw.set("")
        self.login_msg.configure(text=msg)
        self.login_entry.focus_set()

    def do_login(self):
        if self.login_btn.instate(["disabled"]):
            return
        # frisch einlesen, damit Änderungen anderer Benutzer gelten
        self.settings, self.items = load_config()
        load_prefs(self.settings)
        if STATE["error"]:
            return self.locked()
        user = find_user(self.settings["users"], self.login_name.get())
        if not (user and verify_password(self.login_pw.get(), user["hash"])):
            self.login_pw.set("")
            self.login_msg.configure(text="Benutzername oder Passwort ist falsch.")
            self.login_btn.state(["disabled"])   # kurze Sperre gegen schnelles Durchprobieren
            self.after(1500, lambda: self.login_btn.state(["!disabled"]))
            return
        self.user = user
        if STATE["legacy"] and self.can_edit:   # ältere Einstellungen dauerhaft übernehmen
            self.save()
        self.show_main()

    def show_main(self):
        self.login.pack_forget()
        self.who.configure(text=f"Angemeldet: {self.user['name']}")
        self.session.pack(side="right")
        self.area.pack(fill="both", expand=True)
        self.apply_theme()
        self.build(force=True)

    def logout(self):
        self.login_name.set("")
        self.show_login()

    def locked(self):
        messagebox.showerror(
            APP_NAME, "Die Einstellungsdatei ist beschädigt oder wurde außerhalb von DocuBuilder "
                      f"verändert:\n{CONFIG}\n\nEin Windows-Administrator muss die Datei löschen. "
                      "Danach startet DocuBuilder mit der Ersteinrichtung.", parent=self)
        self.destroy()

    # ---- Fenster
    def open_settings(self):
        if not self.user or self._alive(self.default_win):
            return
        if self._alive(self.settings_win):
            self.settings_win.lift()
            return
        self.settings_win = Settings(self)

    def open_default(self, first_run=False):
        if self._alive(self.settings_win):
            return
        if self._alive(self.default_win):
            self.default_win.lift()
            return
        if not first_run:
            if not self.user:
                return
            if not self.can_edit:
                messagebox.showinfo(APP_NAME, NO_RIGHTS, parent=self)
                return
        self.default_win = DefaultDialog(self, first_run=first_run)

    def startup(self):
        if STATE["error"]:
            return self.locked()
        if has_admin(self.settings):
            self.show_login()
        else:                                    # Ersteinrichtung: Admin-Passwort und Benutzer
            self.show_login("Ersteinrichtung läuft …")
            self.open_default(first_run=True)

    def save(self):
        if not self.can_edit:
            messagebox.showerror(APP_NAME, NO_RIGHTS, parent=self)
            return False
        if not save_config(self.settings, self.items):
            return False
        if STATE["legacy"]:                      # ältere Datei übernommen: entfernen
            save_prefs(self.settings)
            unprotect_file(STATE["legacy"])
            try:
                os.remove(STATE["legacy"])
            except OSError:
                pass
            STATE["legacy"] = None
        return True

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
        if not self.user:
            return
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
                          command=lambda path=it["path"]: open_target(path, self.settings["files_dir"]))
            b.bind("<Enter>", lambda e, w=b: w.configure(bg=p["btn_hover"]))
            b.bind("<Leave>", lambda e, w=b: w.configure(bg=p["panel"]))
            b.grid(row=i // cols, column=i % cols, padx=5, pady=5, sticky="nsew", ipady=6)


if __name__ == "__main__":
    if "--setup-machine" in sys.argv[1:]:
        # vom Installationsprogramm (mit Adminrechten) aufgerufen: gemeinsamen Ordner einrichten
        prepare_machine_dir(MACHINE_DIR)
    else:
        App().mainloop()
