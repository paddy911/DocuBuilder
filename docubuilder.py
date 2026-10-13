#!/usr/bin/env python3
"""DocuBuilder – Buttons, die hinterlegte Excel-Listen öffnen.

Läuft unter Windows und Linux, benötigt nur Python 3 (tkinter ist enthalten).
Die Einstellungen liegen in config.json neben dem Programm (portabel).
"""
import json
import os
import re
import subprocess
import sys
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_NAME = "DocuBuilder"


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


def load_items():
    try:
        with open(CONFIG, encoding="utf-8") as f:
            data = json.load(f)
        return [{"name": str(x.get("name", "")), "path": str(x.get("path", ""))}
                for x in data if isinstance(x, dict)]
    except (OSError, ValueError):
        return []


def save_items(items):
    try:
        with open(CONFIG, "w", encoding="utf-8") as f:
            json.dump(items, f, ensure_ascii=False, indent=2)
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


# ---------------------------------------------------------------- Einstellungen
class Settings(tk.Toplevel):
    def __init__(self, master, items, on_change):
        super().__init__(master)
        self.items = items
        self.on_change = on_change
        self.title("Einstellungen")
        self.geometry("720x460")
        self.minsize(560, 380)
        self.transient(master)
        self.grab_set()

        frame = ttk.Frame(self, padding=12)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(frame, columns=("name", "path"), show="headings",
                                 selectmode="browse")
        self.tree.heading("name", text="Name")
        self.tree.heading("path", text="Pfad / Link")
        self.tree.column("name", width=180)
        self.tree.column("path", width=380)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        side = ttk.Frame(frame)
        side.grid(row=0, column=2, sticky="n", padx=(10, 0))
        for text, cmd in (("Nach oben", lambda: self.move(-1)),
                          ("Nach unten", lambda: self.move(1)),
                          ("Löschen", self.delete)):
            ttk.Button(side, text=text, command=cmd, width=12).pack(pady=2)

        form = ttk.LabelFrame(frame, text="Eintrag", padding=10)
        form.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(12, 0))
        form.columnconfigure(1, weight=1)
        self.name_var = tk.StringVar()
        self.path_var = tk.StringVar()
        ttk.Label(form, text="Name:").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Entry(form, textvariable=self.name_var).grid(row=0, column=1, columnspan=2,
                                                         sticky="ew", pady=3)
        ttk.Label(form, text="Pfad / Link:").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(form, textvariable=self.path_var).grid(row=1, column=1, sticky="ew",
                                                         pady=3)
        ttk.Button(form, text="Durchsuchen …", command=self.browse).grid(row=1, column=2,
                                                                          padx=(6, 0))
        btns = ttk.Frame(form)
        btns.grid(row=2, column=0, columnspan=3, sticky="e", pady=(8, 0))
        ttk.Button(btns, text="Neuer Eintrag", command=self.new).pack(side="left", padx=4)
        ttk.Button(btns, text="Speichern", command=self.save_entry).pack(side="left")

        ttk.Button(frame, text="Schließen", command=self.destroy).grid(
            row=2, column=0, columnspan=3, sticky="e", pady=(12, 0))
        self.refresh()

    def refresh(self, select=None):
        self.tree.delete(*self.tree.get_children())
        for i, it in enumerate(self.items):
            self.tree.insert("", "end", iid=str(i), values=(it["name"], it["path"]))
        if select is not None and 0 <= select < len(self.items):
            self.tree.selection_set(str(select))
            self.tree.see(str(select))

    def selected(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    def on_select(self, _event=None):
        i = self.selected()
        if i is not None:
            self.name_var.set(self.items[i]["name"])
            self.path_var.set(self.items[i]["path"])

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
        name, path = self.name_var.get().strip(), self.path_var.get().strip()
        if not name or not path:
            messagebox.showinfo(APP_NAME, "Bitte Name und Pfad eintragen.", parent=self)
            return
        i = self.selected()
        if i is None:
            self.items.append({"name": name, "path": path})
            i = len(self.items) - 1
        else:
            self.items[i] = {"name": name, "path": path}
        self.commit(i)

    def delete(self):
        i = self.selected()
        if i is None:
            return
        if messagebox.askyesno(APP_NAME, f"„{self.items[i]['name']}“ löschen?", parent=self):
            del self.items[i]
            self.new()
            self.commit(None)

    def move(self, d):
        i = self.selected()
        if i is None or not 0 <= i + d < len(self.items):
            return
        self.items[i], self.items[i + d] = self.items[i + d], self.items[i]
        self.commit(i + d)

    def commit(self, select):
        save_items(self.items)
        self.on_change()
        self.refresh(select)


# ---------------------------------------------------------------- Hauptfenster
class App(tk.Tk):
    TILE_W = 220

    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("720x420")
        self.minsize(360, 240)
        self.items = load_items()
        if not self.items:
            self.items = [{"name": "Beispiel: Kundenliste",
                           "path": os.path.join(os.path.expanduser("~"), "Kundenliste.xlsx")}]
            save_items(self.items)

        top = ttk.Frame(self, padding=(14, 12))
        top.pack(fill="x")
        ttk.Label(top, text=APP_NAME, font=("Segoe UI", 15, "bold")).pack(side="left")
        ttk.Button(top, text="Einstellungen", command=self.open_settings).pack(side="right")

        self.area = ttk.Frame(self, padding=(14, 0, 14, 14))
        self.area.pack(fill="both", expand=True)
        self.cols = 0
        self.area.bind("<Configure>", lambda e: self.build())
        self.build(force=True)

    def build(self, force=False):
        width = max(self.area.winfo_width(), 1)
        cols = max(1, width // (self.TILE_W + 10))
        if cols == self.cols and not force:
            return
        self.cols = cols
        for w in self.area.winfo_children():
            w.destroy()
        if not self.items:
            ttk.Label(self.area, text="Noch keine Einträge. Lege sie unter „Einstellungen“ an."
                      ).pack(pady=30)
            return
        for c in range(cols):
            self.area.columnconfigure(c, weight=1, uniform="t")
        for i, it in enumerate(self.items):
            b = tk.Button(self.area, text=it["name"] or "Ohne Namen", wraplength=self.TILE_W - 20,
                          height=3, relief="groove", bd=2, font=("Segoe UI", 11),
                          cursor="hand2", command=lambda p=it["path"]: open_target(p))
            b.grid(row=i // cols, column=i % cols, padx=5, pady=5, sticky="nsew")

    def open_settings(self):
        Settings(self, self.items, lambda: self.build(force=True))


if __name__ == "__main__":
    App().mainloop()
