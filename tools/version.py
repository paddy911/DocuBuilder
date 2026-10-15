#!/usr/bin/env python3
"""Versionsnummern von DocuBuilder (Major.Minor.Patch.Build).

Stable (Branch main):   1.0.0.0 -> 1.1.0.0 -> 1.2.0.0 ...   (je Merge/Push auf main)
Nightly (Branch nightly): Stable-Version + Build-Nummer, z. B. 1.0.0.1, 1.0.0.2 ...
                          Die Build-Nummer zählt die eigenen Commits auf nightly seit dem
                          letzten Abgleich mit main und wird nur beim Bauen gesetzt, nie committet.

  python tools/version.py show                        aktuelle Version
  python tools/version.py set 1.2.0.0 [--channel nightly]
  python tools/version.py nightly                     Nightly-Version berechnen (ändert nichts)
  python tools/version.py release                     nächste Stable-Version setzen: ist die aktuelle
                                                      Version schon veröffentlicht (Tag v<Version>),
                                                      wird Minor erhöht, sonst bleibt sie
"""
import argparse
import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
APP = os.path.join(ROOT, "docubuilder.py")
INFO = os.path.join(ROOT, "version_info.txt")
STABLE_BRANCH = os.environ.get("STABLE_BRANCH", "main")
VERSION_RE = re.compile(r'^APP_VERSION = "[^"]*"', re.M)
CHANNEL_RE = re.compile(r'^APP_CHANNEL = "[^"]*"', re.M)


def read(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


def write(path, text):
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write(text)


def fmt(v):
    return ".".join(map(str, v))


def parse(s):
    parts = s.strip().split(".")
    if len(parts) != 4 or not all(p.isdigit() and int(p) <= 65535 for p in parts):
        sys.exit(f"Ungültige Version: {s!r} (erwartet z. B. 1.2.0.0)")
    return tuple(int(p) for p in parts)


def current():
    m = re.search(r'^APP_VERSION = "([^"]*)"', read(APP), re.M)
    if not m:
        sys.exit("APP_VERSION nicht in docubuilder.py gefunden")
    return parse(m.group(1))


def set_version(v, channel="stable"):
    s = fmt(v)
    app = VERSION_RE.sub(f'APP_VERSION = "{s}"', read(APP), count=1)
    write(APP, CHANNEL_RE.sub(f'APP_CHANNEL = "{channel}"', app, count=1))
    tup = ", ".join(map(str, v))
    info = re.sub(r"(filevers|prodvers)=\([\d, ]*\)", lambda m: f"{m.group(1)}=({tup})", read(INFO))
    info = re.sub(r"StringStruct\('(FileVersion|ProductVersion)', '[^']*'\)",
                  lambda m: f"StringStruct('{m.group(1)}', '{s}')", info)
    write(INFO, info)


def git(*args):
    return subprocess.run(["git", *args], cwd=ROOT, capture_output=True, text=True,
                          check=True).stdout.strip()


def nightly():
    base = current()
    try:
        n = int(git("rev-list", "--count", "--no-merges", f"origin/{STABLE_BRANCH}..HEAD"))
    except (OSError, subprocess.CalledProcessError, ValueError):
        sys.exit(f"Build-Nummer nicht ermittelbar (git und origin/{STABLE_BRANCH} nötig)")
    return base[:3] + (max(1, n),)


def release():
    v = current()
    if git("tag", "--list", f"v{fmt(v)}"):
        v = (v[0], v[1] + 1, 0, 0)
    set_version(v)
    return v


def main():
    ap = argparse.ArgumentParser(description="Versionsnummern von DocuBuilder")
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("show")
    p = sub.add_parser("set")
    p.add_argument("version")
    p.add_argument("--channel", choices=("stable", "nightly"), default="stable")
    sub.add_parser("nightly")
    sub.add_parser("release")
    a = ap.parse_args()
    if a.cmd == "show":
        print(fmt(current()))
    elif a.cmd == "set":
        set_version(parse(a.version), a.channel)
        print(a.version)
    elif a.cmd == "nightly":
        print(fmt(nightly()))
    else:
        print(fmt(release()))


if __name__ == "__main__":
    main()
