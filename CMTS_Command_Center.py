"""
ARRIS E6000 CMTS Command Center  -  v0.3.0

Operations console for ARRIS / CommScope E6000 CER chassis, reached directly or through
a bastion (jump) host with a one-time VIP token.

What changed in 0.3.0 (rewrite of the 0.2.x Gemini build)
  * Read-only guardrail removed - any CLI command can be sent (disruptive commands ask
    for confirmation, which can be switched off in Settings).
  * New SSH engine: waits for the real CMTS prompt instead of the first '>' or '#',
    answers --More-- pagers, handles [y/n] confirmations, wide PTY so tables don't wrap,
    one persistent bastion session (VIP token is asked once, not per command / per node),
    CMTS channels are cached so hopping between nodes is instant, Stop button sends Ctrl-C.
  * All SSH work runs on one background worker, parsing happens off the UI thread,
    UI updates are marshalled through a queue (no more cross-thread Tk calls / freezes).
  * Column-aware output parser: every SNR / MER / power / codeword / utilisation /
    temperature / state cell is coloured against configurable thresholds; port rows in
    summary tables are coloured by online %.
  * New Table view (sortable, filterable, "issues only"), redesigned charts, colour-coded
    fleet leaderboard, colour-coded triage + maintenance diff, history replay.
  * Fleet audit shares the bastion session (parallel channels, one VIP token).
  * Excel export uses the real column headers with health colouring; HTML export is escaped.
  * 0.3.1: "Output -> Excel" saves the on-screen output (colours kept) to a location you choose.
  * Config + history live in ~/.e6000_command_center (old files are migrated).
"""

from __future__ import annotations

import base64
import codecs
import difflib
import html
import io
import json
import queue
import re
import shutil
import socket
import sqlite3
import sys
import threading
import time
import tkinter as tk
import webbrowser
from bisect import bisect_right
from collections import Counter, deque
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field, replace
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Callable, Dict, List, Optional, Tuple

import customtkinter as ctk
import paramiko

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:  # export is disabled gracefully
    openpyxl = None

VERSION = "0.3.2"
APP_NAME = "E6000 Command Center"
UI_FONT = "Segoe UI"
MONO_FONT = "Consolas"

# =============================================================================
# Paths & configuration
# =============================================================================
APP_DIR = Path.home() / ".e6000_command_center"
CONFIG_FILE = APP_DIR / "config.json"
DB_FILE = APP_DIR / "flight_recorder.db"


def resource_path(name: str) -> Path:
    """Locate a bundled file (e.g. assets/app.ico) both from source and from a PyInstaller build."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / name


def _migrate_legacy_files() -> None:
    try:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        for legacy, new in ((Path("config.json"), CONFIG_FILE), (Path("flight_recorder.db"), DB_FILE)):
            if legacy.exists() and not new.exists():
                shutil.copy2(legacy, new)
    except Exception:
        pass


DEFAULT_SETTINGS: Dict[str, object] = {
    "theme": "Midnight",
    "hop_cmd": "ssh {host}",
    "paging_cmds": ["terminal length 0", "page-length 0"],
    "otp_mode": "append",          # append = password+token in one prompt, separate = two prompts
    "login_timeout": 40,
    "cmd_timeout": 120,
    "fleet_workers": 6,
    "confirm_disruptive": True,
    "auto_charts": False,
    "snapshot_cmds": ["show cable modem summary port mac | include mac", "show cable flap-list"],
    "term_max_lines": 60000,
    "export_dir": "",              # last folder used by a Save dialog (remembered between runs)
}

# kind: min = higher is better, max = lower is better, window = inside [lo, hi] is best
DEFAULT_THRESHOLDS: Dict[str, dict] = {
    "us_snr":     {"label": "US SNR / MER (dB)",        "kind": "min",    "good": 30.0,         "warn": 27.0,          "range": [0.1, 60]},
    "ds_snr":     {"label": "DS SNR / MER (dB)",        "kind": "min",    "good": 35.0,         "warn": 33.0,          "range": [0.1, 60]},
    "us_tx":      {"label": "Modem US Tx power (dBmV)", "kind": "window", "good": [35.0, 49.0], "warn": [30.0, 52.0],  "range": [1, 70]},
    "ds_rx":      {"label": "Modem DS Rx power (dBmV)", "kind": "window", "good": [-7.0, 7.0],  "warn": [-10.0, 10.0], "range": [-40, 40]},
    "cmts_rx":    {"label": "CMTS US Rx level (dBmV)",  "kind": "window", "good": [-3.0, 3.0],  "warn": [-6.0, 6.0],   "range": [-30, 30]},
    "uncorr_pct": {"label": "Uncorrectable CW (%)",     "kind": "max",    "good": 0.1,          "warn": 1.0,           "range": [0, 100]},
    "corr_pct":   {"label": "Corrected CW (%)",         "kind": "max",    "good": 5.0,          "warn": 15.0,          "range": [0, 100]},
    "uncorr":     {"label": "Uncorrectable CW (count)", "kind": "max",    "good": 0,            "warn": 1000,          "range": [0, 1e15]},
    "util":       {"label": "Utilisation (%)",          "kind": "max",    "good": 70.0,         "warn": 85.0,          "range": [0, 100]},
    "temp":       {"label": "Temperature (C)",          "kind": "max",    "good": 55.0,         "warn": 70.0,          "range": [-30, 150]},
    "flap":       {"label": "Flap count",               "kind": "max",    "good": 9,            "warn": 49,            "range": [0, 1e9]},
    "online_pct": {"label": "Port online (%)",          "kind": "min",    "good": 98.0,         "warn": 90.0,          "range": [0, 100]},
    "opt_rx":     {"label": "Optical Rx (dBm)",         "kind": "window", "good": [-12.0, 0.0], "warn": [-16.0, 2.0],  "range": [-40, 10]},
}
METRIC_SHORT = {
    "us_snr": "US SNR", "ds_snr": "DS SNR", "us_tx": "US Tx", "ds_rx": "DS Rx", "cmts_rx": "US Rx",
    "uncorr_pct": "Uncorr%", "corr_pct": "Corr%", "uncorr": "Uncorr", "util": "Util", "temp": "Temp",
    "flap": "Flaps", "online_pct": "Online%", "opt_rx": "Opt Rx",
}


def load_config() -> dict:
    try:
        if CONFIG_FILE.exists():
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def save_config(data: dict) -> None:
    try:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        CONFIG_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception:
        pass


# =============================================================================
# Themes
# =============================================================================
THEMES: Dict[str, dict] = {
    "Midnight": {
        "mode": "Dark", "bg": "#0E1116", "panel": "#131821", "card": "#181E29", "input": "#0D1118",
        "border": "#273042", "hover": "#222B3A", "accent": "#4C8DFF", "accent_hover": "#3B76E0",
        "on_accent": "#FFFFFF", "text": "#E6EAF2", "muted": "#8C96A8", "term_bg": "#0B0E13",
        "term_fg": "#C8D0DC", "ok": "#3FB950", "warn": "#E3B341", "crit": "#FF6B61", "mac": "#C9A2FF",
        "ip": "#7CC4FF", "iface": "#3CCBD6", "th": "#9DB4D9", "dim": "#5C6678", "sel": "#264F78",
        "stripe": "#10141B", "hit": "#5A4A12", "hit_cur": "#E3B341",
    },
    "Graphite": {
        "mode": "Dark", "bg": "#161718", "panel": "#1C1D1F", "card": "#222326", "input": "#151618",
        "border": "#34363B", "hover": "#2C2E33", "accent": "#2BB3A3", "accent_hover": "#229A8C",
        "on_accent": "#06110F", "text": "#ECECEC", "muted": "#9A9DA3", "term_bg": "#121314",
        "term_fg": "#D4D4D4", "ok": "#6CCB5F", "warn": "#F2C14E", "crit": "#FF6E6E", "mac": "#D7A6FF",
        "ip": "#8CC8FF", "iface": "#4FD1C5", "th": "#B5BAC4", "dim": "#666A72", "sel": "#2F4F4A",
        "stripe": "#17181A", "hit": "#4D4317", "hit_cur": "#F2C14E",
    },
    "Daylight": {
        "mode": "Light", "bg": "#F3F5F9", "panel": "#FFFFFF", "card": "#FFFFFF", "input": "#F6F8FB",
        "border": "#D9DFE9", "hover": "#EDF1F7", "accent": "#2F6FEB", "accent_hover": "#245BC4",
        "on_accent": "#FFFFFF", "text": "#1B2230", "muted": "#5D6B82", "term_bg": "#FFFFFF",
        "term_fg": "#1F2328", "ok": "#1A7F37", "warn": "#9A6700", "crit": "#CF222E", "mac": "#8250DF",
        "ip": "#0550AE", "iface": "#0B7A83", "th": "#3D4F6E", "dim": "#8C959F", "sel": "#CFE2FF",
        "stripe": "#F7F9FC", "hit": "#FFF1B8", "hit_cur": "#F2C14E",
    },
}

# =============================================================================
# Command catalog
# =============================================================================
@dataclass
class CommandSpec:
    title: str
    desc: str
    template: str
    params: List[str] = field(default_factory=list)

    def format(self, **kw) -> str:
        cmd = self.template
        for key in self.params:
            cmd = cmd.replace("{" + key + "}", (kw.get(key) or "").strip())
        return re.sub(r"\s+", " ", cmd).strip()


@dataclass
class PlaybookSpec:
    title: str
    desc: str
    templates: List[str]
    params: List[str] = field(default_factory=list)

    def format(self, **kw) -> List[str]:
        return [CommandSpec(self.title, "", t, self.params).format(**kw) for t in self.templates]


COMMAND_CATALOG: Dict[str, CommandSpec] = {c.title: c for c in [
    CommandSpec("Modem Counts (Port Summary)", "Port-by-port registered / offline / total counts.",
                "show cable modem summary port mac | include mac"),
    CommandSpec("Modem Summary All", "Chassis-wide modem state breakdown.", "show cable modem summary"),
    CommandSpec("Modem Summary by Cable-Mac", "Sign-on and operational distribution for one cable-mac.",
                "show cable modem summary mac cable-mac {cmac}", ["cmac"]),
    CommandSpec("Modems on Cable-Mac", "Every modem on a cable-mac with state and levels.",
                "show cable modem cable-mac {cmac}", ["cmac"]),
    CommandSpec("Cable Noise & SNR", "Upstream SNR / MER and corrected / uncorrectable codewords.",
                "show cable noise cable-mac {cmac}", ["cmac"]),
    CommandSpec("OFDM CM Count Profile", "DOCSIS 3.1 OFDM profile distribution.",
                "show interface cable-mac {cmac} ofdm cm-count-profile", ["cmac"]),
    CommandSpec("Heavy-hitters Downstream", "Top downstream consumers on a cable-mac.",
                "show cable modem heavyhitter downstream cable-mac {cmac}", ["cmac"]),
    CommandSpec("Modem Verbose", "Full PHY / timing / power / RxMER detail for one modem.",
                "show cable modem {mac_address} verbose", ["mac_address"]),
    CommandSpec("Modem PHY Levels", "Upstream / downstream power and SNR for one modem.",
                "show cable modem {mac_address} phy", ["mac_address"]),
    CommandSpec("Modem CPE Devices", "CPE IP / MAC addresses behind a modem.",
                "show cable modem {mac_address} cpe", ["mac_address"]),
    CommandSpec("Modem Service Flows", "QoS service flows and rates for one modem.",
                "show cable modem {mac_address} service-flow", ["mac_address"]),
    CommandSpec("Flap List", "Unstable modems, ranging aborts, power adjustments.", "show cable flap-list"),
    CommandSpec("Rogue Modems", "Modems failing BPI+, cloned MACs.", "show cable modem rogue"),
    CommandSpec("Upstream Capacity", "Utilisation and state of all upstream channels.",
                "show interface cable-upstream summary"),
    CommandSpec("Downstream Capacity", "Utilisation and state of all downstream channels.",
                "show interface cable-downstream summary"),
    CommandSpec("Optical Transceiver", "SFP / QSFP light levels and temperature.",
                "show interface {interface} transceiver", ["interface"]),
    CommandSpec("Video Jitter", "RSM jitter buffer, underflows and overflows.",
                "show video jitter-status cable-downstream {interface}", ["interface"]),
    CommandSpec("Chassis Environment", "Power, fans, temperatures.", "show environment"),
    CommandSpec("Linecard Status", "Linecard operational state.", "show linecard status"),
    CommandSpec("Redundancy", "Active / standby RP state.", "show redundancy"),
    CommandSpec("Recent Log", "Most recent log entries first.", "show logging reverse"),
    CommandSpec("Version", "Software version and uptime.", "show version"),
]}

PLAYBOOK_CATALOG: Dict[str, PlaybookSpec] = {p.title: p for p in [
    PlaybookSpec("Modem 360 + Truck-Roll Classifier",
                 "Verbose, PHY, service flows, CPE and flap history for one modem, then Premise vs Plant verdict.",
                 ["show cable modem {mac_address} verbose", "show cable modem {mac_address} phy",
                  "show cable modem {mac_address} service-flow", "show cable modem {mac_address} cpe",
                  "show cable flap-list | include {mac_address}"], ["mac_address"]),
    PlaybookSpec("Cable-Mac Blast-Radius Sweep",
                 "Summary, noise, OFDM profiles and flap list for one cable-mac to isolate shared impairments.",
                 ["show cable modem summary mac cable-mac {cmac}", "show cable noise cable-mac {cmac}",
                  "show interface cable-mac {cmac} ofdm cm-count-profile", "show cable flap-list"], ["cmac"]),
    PlaybookSpec("Node Capacity & Core Health",
                 "Port summary, US/DS capacity, redundancy and environment.",
                 ["show cable modem summary port mac | include mac", "show interface cable-upstream summary",
                  "show interface cable-downstream summary", "show redundancy", "show environment"]),
]}

AUTOCOMPLETE_BASE = sorted(set(
    [c.template.split("{")[0].strip() + (" " if "{" in c.template else "") for c in COMMAND_CATALOG.values()] + [
        "show cable modem", "show cable modem count cable-mac ", "show cable modem counters",
        "show cable modem heavyhitter upstream", "show cable modem ofdm-iuc-flaps",
        "show interface cable-upstream codeword-error-summary", "show ofdm global-settings",
        "show ofdm profile-summary", "show video service-group", "show video sessions",
        "show cable mac-domain cable rcc", "show running-config", "show clock", "ping ",
        "configure", "exit", "end", "clear cable modem ",
    ]))

DISRUPTIVE_RE = re.compile(
    r"^\s*(conf(ig(ure)?)?|clear|reset|reload|reboot|shut(down)?|no|write|copy|delete|erase|format|debug|"
    r"undebug|test|cable\s+modem\s+reset)\b|\breset\b", re.I)


def is_disruptive(cmd: str) -> bool:
    return bool(DISRUPTIVE_RE.search(cmd))


# =============================================================================
# OUI vendor map
# =============================================================================
OUI_VENDOR_MAP: Dict[str, str] = {
    "00015C": "CADANT / ARRIS", "0015CE": "ARRIS", "001DDF": "ARRIS", "203D66": "ARRIS / COMMSCOPE",
    "40B89A": "ARRIS / COMMSCOPE", "8461A0": "ARRIS / COMMSCOPE", "F8A097": "ARRIS / COMMSCOPE",
    "001018": "BROADCOM", "001A6B": "UBEE", "64ED57": "UBEE", "001E46": "CISCO / SA", "0023BE": "CISCO / SA",
    "105F06": "TECHNICOLOR", "589043": "TECHNICOLOR", "E45D51": "TECHNICOLOR", "3894ED": "NETGEAR",
    "A040A0": "NETGEAR", "0050F1": "HITRON", "68B6FC": "HITRON",
}


def lookup_mac_vendor(mac: str) -> str:
    hx = re.sub(r"[^0-9A-Fa-f]", "", mac).upper()
    if len(hx) < 6:
        return "Invalid MAC"
    oui = hx[:6]
    return f"{oui[:2]}:{oui[2:4]}:{oui[4:6]} ({OUI_VENDOR_MAP.get(oui, 'unlisted vendor')})"


# =============================================================================
# Output parsing & colourisation (pure functions - no Tk, safe in worker threads)
# =============================================================================
TOKEN_RE = re.compile(r"\S+")
GROUP_RE = re.compile(r"\S+(?: \S+)*")
NUM_RE = re.compile(r"(?<![\w./])(-?\d{1,3}(?:,\d{3})+(?:\.\d+)?|-?\d+(?:\.\d+)?)([A-Za-z%°]*)")
UNITS_OK = {"", "db", "dbmv", "dbm", "dbc", "%", "c", "°c", "dbuv"}
MAC_RE = re.compile(r"\b(?:[0-9a-fA-F]{4}\.[0-9a-fA-F]{4}\.[0-9a-fA-F]{4}|(?:[0-9a-fA-F]{2}[:-]){5}[0-9a-fA-F]{2})\b")
IP_RE = re.compile(r"\b(?:25[0-5]|2[0-4]\d|1?\d?\d)(?:\.(?:25[0-5]|2[0-4]\d|1?\d?\d)){3}\b")
IFACE_RE = re.compile(r"\b\d{1,2}/\d{1,2}(?:/\d{1,3}){1,2}(?:[-:]\d{1,2}/\d{1,2}(?:/\d{1,3}){1,2})?\b")
STATE_WORD_RE = re.compile(
    r"\b(operational|w-online|p-online|online(?:\([a-z]+\))?|offline|init\([a-z0-9]+\)|ranging|"
    r"reject\([a-z]+\)|degraded|failed|failure|critical|alarm)\b", re.I)
ERROR_RE = re.compile(r"^\s*%|invalid input|unknown command|incomplete command|ambiguous command|syntax error|^\s*error\b", re.I)
CMD_PROMPT_LINE_RE = re.compile(
    r"^([\w.\-]+(?:\([\w\-]+\))?[#>])\s*((?:show|ping|conf|configure|clear|exit|end|no|interface|cable|page-length|terminal)\b.*)$",
    re.I)
KV_KEY_RE = re.compile(r"([A-Za-z](?:[A-Za-z0-9/()%._\-]| (?! ))*?)\s*[:=](?=\s|$|-?\d)")
SEV_RANK = {"crit": 3, "warn": 2, "ok": 1, None: 0}


def worst(*sevs: Optional[str]) -> Optional[str]:
    best = None
    for s in sevs:
        if SEV_RANK.get(s, 0) > SEV_RANK.get(best, 0):
            best = s
    return best


def find_numbers(s: str, base: int = 0) -> List[Tuple[int, int, float]]:
    out = []
    for m in NUM_RE.finditer(s):
        unit = m.group(2).lower()
        if unit not in UNITS_OK:
            continue
        end = m.end()
        if end < len(s) and s[end] in "./":
            continue
        try:
            out.append((base + m.start(1), base + end, float(m.group(1).replace(",", ""))))
        except ValueError:
            pass
    return out


def classify_value(metric: str, v: float, th: Dict[str, dict]) -> Optional[str]:
    r = th.get(metric)
    if not r:
        return None
    lo, hi = r.get("range", [-1e18, 1e18])
    if not (lo <= v <= hi):
        return None
    kind, good, warn_ = r["kind"], r["good"], r["warn"]
    if kind == "min":
        return "ok" if v >= good else "warn" if v >= warn_ else "crit"
    if kind == "max":
        return "ok" if v <= good else "warn" if v <= warn_ else "crit"
    return "ok" if good[0] <= v <= good[1] else "warn" if warn_[0] <= v <= warn_[1] else "crit"


def classify_state(s: str) -> Optional[str]:
    x = s.lower().strip()
    if not x:
        return None
    if re.search(r"not\s*oper|offline|\bdown\b|reject|fail|denied|critical|alarm|\berror|\bfault|unreach|absent|missing", x):
        return "crit"
    if re.search(r"init|ranging|p-online|registered|degraded|partial|\bdhcp|\btftp|\btod\b|\bwarn|minor|resetting|booting|\bdsx|expired", x):
        return "warn"
    if re.search(r"operational|online|\boper\b|\bup\b|active|normal|\bok\b|in\s*service|enabled|present|running|ready|good", x):
        return "ok"
    return None


def classify_header(name: str) -> Tuple[Optional[str], Optional[str]]:
    """Map a column header / key to (metric, role)."""
    s = name.lower().strip()
    if not s or len(s) > 70:
        return None, None
    c = re.sub(r"[\s_\-]", "", s)
    is_ds = c.startswith("ds") or "downstream" in s or bool(re.search(r"\bds", s))
    is_us = c.startswith("us") or "upstream" in s or bool(re.search(r"\bus\b", s))
    pct = "%" in s or "pct" in s or "percent" in s or "ratio" in s
    if "snr" in c or "rxmer" in c or re.search(r"\bmer\b", s) or c.startswith("mer"):
        return ("ds_snr" if is_ds and not is_us else "us_snr"), None
    if "uncorr" in c:
        return ("uncorr_pct" if pct else "uncorr"), None
    if "corr" in c:
        return ("corr_pct" if pct else None), None
    if "util" in c:
        return "util", None
    if "temp" in c:
        return "temp", None
    if "flap" in c:
        return "flap", None
    if "dbm" in s and "dbmv" not in s and ("rx" in c or "receiv" in c):
        return "opt_rx", None
    if "pwr" in c or "power" in c or "level" in c:
        if "tx" in c or "transmit" in c:
            return "us_tx", None
        if is_us and ("rx" in c or "receiv" in c):
            return "cmts_rx", None
        if is_ds or "receiv" in c:
            return "ds_rx", None
        if is_us:
            return "us_tx", None
        if "rx" in c:
            return "cmts_rx", None
        return None, None
    if re.search(r"\b(state|status)\b", s):
        return None, "state"
    if re.search(r"mac\s*addr|^mac$|^cm\s*mac|modem\s*mac|^cmmac", s):
        return None, "mac"
    if re.search(r"^ip\b|\bip\s*addr|^ipv[46]|cpe\s*ip", s):
        return None, "ip"
    if "offline" in c:
        return None, "offline"
    if re.search(r"\b(oper|operational|online|registered|reg|active|w-online)\b", s):
        return None, "online"
    if "total" in c:
        return None, "total"
    if re.search(r"interface|\bi/f\b|^if$|^port\b|cable-?mac|mac\s*domain|^chan|channel|^upstream$|^downstream$|^ifc", s):
        return None, "iface"
    return None, None


@dataclass
class Column:
    name: str
    lo: int
    hi: int
    metric: Optional[str] = None
    role: Optional[str] = None


@dataclass
class Table:
    columns: List[Column]
    rows: List[List[str]] = field(default_factory=list)
    cell_sev: List[List[Optional[str]]] = field(default_factory=list)
    row_sev: List[Optional[str]] = field(default_factory=list)
    row_issues: List[List[str]] = field(default_factory=list)

    def col(self, role: str = None, metric: str = None) -> Optional[int]:
        for i, c in enumerate(self.columns):
            if (role and c.role == role) or (metric and c.metric == metric):
                return i
        return None

    @property
    def has_modems(self) -> bool:
        return self.col(role="mac") is not None


@dataclass
class Parsed:
    lines: List[str]
    spans: List[List[Tuple[int, int, str]]]
    tables: List[Table]
    metrics: Dict[str, List[float]]
    scope_metrics: Dict[str, Dict[str, List[float]]]
    points: List[Tuple[str, float, str, Optional[str]]]   # metric, value, label, sev
    summary: Optional[Dict[str, int]]
    errors: List[str]

    @property
    def issue_count(self) -> int:
        return sum(1 for t in self.tables for s in t.row_sev if s in ("warn", "crit"))


def _is_dash_line(line: str) -> bool:
    s = line.strip()
    return len(s) >= 3 and ("--" in s or "==" in s) and not re.search(r"[^\-=+| ]", s)


def _nearest(los: List[int], his: List[int], s: int, e: int) -> int:
    mid = (s + e - 1) / 2.0
    i = bisect_right(los, mid) - 1
    best, bd = 0, float("inf")
    for j in (i - 1, i, i + 1, i + 2):
        if 0 <= j < len(los):
            ov = min(e, his[j]) - max(s, los[j])
            d = -ov if ov > 0 else min(abs(mid - los[j]), abs(mid - (his[j] - 1)))
            if d < bd:
                best, bd = j, d
    return best


def _header_groups(line: str) -> List[Tuple[int, int, str]]:
    groups = [(m.start(), m.end(), m.group()) for m in GROUP_RE.finditer(line)]
    if len(groups) < 3:
        groups = [(m.start(), m.end(), m.group()) for m in TOKEN_RE.finditer(line)]
    return groups


def _is_header_candidate(line: str) -> bool:
    if not line.strip() or ":" in line or re.search(r"\d+\.\d+", line) or MAC_RE.search(line):
        return False
    groups = _header_groups(line)
    if len(groups) < 3:
        return False
    score = sum(1 for g in groups if any(classify_header(g[2])))
    return score >= 2


def _is_prompt_line(line: str) -> bool:
    return bool(CMD_PROMPT_LINE_RE.match(line))


def _resolve(spans: List[Tuple[int, int, str, int]]) -> List[Tuple[int, int, str]]:
    if not spans:
        return []
    taken: List[Tuple[int, int, str]] = []
    for s, e, tag, _ in sorted(spans, key=lambda x: (-x[3], x[0])):
        if s >= e or any(s < te and ts < e for ts, te, _ in taken):
            continue
        taken.append((s, e, tag))
    taken.sort()
    return taken


def parse_output(text: str, th: Dict[str, dict]) -> Parsed:
    """Single structural pass + colourising pass over CLI output."""
    lines = [ln.expandtabs(8).rstrip() for ln in text.replace("\r", "").split("\n")]
    n = len(lines)
    kind = ["free"] * n
    row_tbl: List[Optional[Table]] = [None] * n
    tables: List[Table] = []

    def collect_rows(start: int, table: Table) -> Tuple[int, int]:
        j = start
        count = 0
        while j < n:
            ln = lines[j]
            if not ln.strip() or _is_prompt_line(ln):
                break
            if _is_dash_line(ln):
                if j + 1 < n and lines[j + 1].strip() and not _is_dash_line(lines[j + 1]):
                    kind[j] = "dash"
                    j += 1
                    continue
                kind[j] = "dash"
                j += 1
                break
            if len(TOKEN_RE.findall(ln)) < 2 or (_is_header_candidate(ln) and j > start):
                break
            kind[j] = "row"
            row_tbl[j] = table
            count += 1
            j += 1
        return j, count

    # ---- pass 1: structure ------------------------------------------------------------------
    i = 0
    while i < n:
        ln = lines[i]
        if _is_dash_line(ln):
            runs = [(m.start(), m.end()) for m in re.finditer(r"[-=]{2,}", ln)]
            kind[i] = "dash"
            if len(runs) >= 2:
                los = [r[0] for r in runs]
                his = [r[1] for r in runs]

                def aligned(text: str) -> bool:
                    toks = [(m.start(), m.end()) for m in TOKEN_RE.finditer(text)]
                    hits = sum(1 for s, e in toks if any(min(e, hi) > max(s, lo) for lo, hi in runs))
                    return bool(toks) and hits / len(toks) >= 0.8

                hdr = []
                j = i - 1
                while j >= 0 and len(hdr) < 3 and lines[j].strip() and kind[j] == "free" \
                        and not _is_dash_line(lines[j]) and not _is_prompt_line(lines[j]) \
                        and (not hdr or aligned(lines[j])):
                    hdr.insert(0, j)
                    j -= 1
                names: List[List[str]] = [[] for _ in runs]
                for h in hdr:
                    for m in TOKEN_RE.finditer(lines[h]):
                        names[_nearest(los, his, m.start(), m.end())].append(m.group())
                    kind[h] = "hdr"
                cols = []
                for nm, (lo, hi) in zip(names, runs):
                    col = Column(" ".join(nm), lo, hi)
                    col.metric, col.role = classify_header(col.name)
                    cols.append(col)
                t = Table(cols)
                i, cnt = collect_rows(i + 1, t)
                if cnt:
                    tables.append(t)
                continue
            i += 1
            continue
        if kind[i] == "free" and _is_header_candidate(ln):
            k = i + 1
            if k < n and _is_dash_line(lines[k]):
                if len(re.findall(r"[-=]{2,}", lines[k])) >= 2:
                    i = k
                    continue
                k += 1
            if k < n and lines[k].strip() and re.search(r"\d", lines[k]):
                groups = _header_groups(ln)
                cols = []
                for s, e, name in groups:
                    col = Column(name, s, e)
                    col.metric, col.role = classify_header(name)
                    cols.append(col)
                t = Table(cols)
                kind[i] = "hdr"
                if k == i + 2:
                    kind[i + 1] = "dash"
                j, cnt = collect_rows(k, t)
                if cnt:
                    tables.append(t)
                i = max(j, i + 1)
                continue
        i += 1

    # ---- pass 2: colourise + extract --------------------------------------------------------------
    metrics: Dict[str, List[float]] = {}
    scope: Dict[str, Dict[str, List[float]]] = {"node": {}, "modem": {}}
    points: List[Tuple[str, float, str, Optional[str]]] = []
    errors: List[str] = []
    kv_roles: Dict[str, int] = {}
    all_spans: List[List[Tuple[int, int, str]]] = []

    def add_metric(metric: str, v: float, sev: Optional[str], label: str, sc: str):
        metrics.setdefault(metric, []).append(v)
        scope[sc].setdefault(metric, []).append(v)
        points.append((metric, v, label, sev))

    for idx, ln in enumerate(lines):
        sp: List[Tuple[int, int, str, int]] = []
        k = kind[idx]
        if not ln.strip():
            all_spans.append([])
            continue
        if k == "hdr":
            all_spans.append([(0, len(ln), "th")])
            continue
        if k == "dash":
            all_spans.append([(0, len(ln), "dim")])
            continue
        pm = CMD_PROMPT_LINE_RE.match(ln)
        if pm:
            all_spans.append([(0, pm.end(1), "prompt"), (pm.start(2), len(ln), "cmd")])
            continue
        if ERROR_RE.search(ln):
            errors.append(ln.strip())
            all_spans.append([(0, len(ln), "crit")])
            continue

        for m in MAC_RE.finditer(ln):
            sp.append((m.start(), m.end(), "mac", 1))
        for m in IP_RE.finditer(ln):
            sp.append((m.start(), m.end(), "ip", 1))
        for m in IFACE_RE.finditer(ln):
            sp.append((m.start(), m.end(), "iface", 1))

        if k == "row":
            t = row_tbl[idx]
            cols = t.columns
            los = [c.lo for c in cols]
            his = [c.hi for c in cols]
            cell_toks: List[List[Tuple[int, int, str]]] = [[] for _ in cols]
            for m in TOKEN_RE.finditer(ln):
                cell_toks[_nearest(los, his, m.start(), m.end())].append((m.start(), m.end(), m.group()))
            cells = [" ".join(tk_[2] for tk_ in ct) for ct in cell_toks]
            sevs: List[Optional[str]] = [None] * len(cols)
            issues: List[str] = []
            sc = "modem" if t.has_modems else "node"
            li = t.col(role="iface")
            lm = t.col(role="mac")
            label = (cells[lm] if lm is not None and cells[lm] else
                     cells[li] if li is not None and cells[li] else
                     next((c for c in cells if c), ""))
            role_vals: Dict[str, Tuple[int, int]] = {}
            for ci, col in enumerate(cols):
                ct = cell_toks[ci]
                if not ct:
                    continue
                cs, ce = ct[0][0], ct[-1][1]
                if col.metric:
                    for (ts, te, tt) in ct:
                        for (ns, ne, v) in find_numbers(tt, ts):
                            sev = classify_value(col.metric, v, th)
                            if sev:
                                sp.append((ns, ne, sev, 3))
                                sevs[ci] = worst(sevs[ci], sev)
                                if sev != "ok":
                                    issues.append(f"{METRIC_SHORT.get(col.metric, col.metric)} {v:g}")
                                add_metric(col.metric, v, sev, label, sc)
                elif col.role == "state":
                    sev = classify_state(cells[ci])
                    if sev:
                        sp.append((cs, ce, sev, 3))
                        sevs[ci] = sev
                        if sev != "ok":
                            issues.append(cells[ci])
                elif col.role in ("total", "online", "offline"):
                    m = re.search(r"-?\d[\d,]*", cells[ci])
                    if m:
                        role_vals[col.role] = (ci, int(m.group().replace(",", "")))
                elif col.role in ("mac", "ip", "iface"):
                    sp.append((cs, ce, col.role, 2))
            # port health from counts
            tot = role_vals.get("total", (None, None))[1]
            onl = role_vals.get("online", (None, None))[1]
            off = role_vals.get("offline", (None, None))[1]
            if tot is not None and onl is None and off is not None:
                onl = tot - off
            if tot is not None and off is None and onl is not None:
                off = tot - onl
            pct_sev = None
            if tot and onl is not None and 0 <= onl <= tot:
                pct = 100.0 * onl / tot
                pct_sev = classify_value("online_pct", pct, th)
                if pct_sev:
                    for role in ("online", "offline"):
                        if role in role_vals:
                            ci = role_vals[role][0]
                            ct = cell_toks[ci]
                            sp.append((ct[0][0], ct[-1][1], pct_sev, 3))
                            sevs[ci] = pct_sev
                    if li is not None and cell_toks[li]:
                        ct = cell_toks[li]
                        sp.append((ct[0][0], ct[-1][1], pct_sev, 4))
                    if pct_sev != "ok":
                        issues.append(f"Online {pct:.1f}% ({off} off)")
                    points.append(("online_pct", pct, label, pct_sev))
            t.rows.append(cells)
            t.cell_sev.append(sevs)
            t.row_sev.append(worst(*sevs, pct_sev))
            t.row_issues.append(issues)
        else:
            # key : value  (verbose output)
            keys = list(KV_KEY_RE.finditer(ln))
            for ki, km in enumerate(keys):
                key = km.group(1).strip()
                vstart = km.end()
                vend = keys[ki + 1].start() if ki + 1 < len(keys) else len(ln)
                region = ln[vstart:vend]
                metric, role = classify_header(key)
                if metric:
                    for (ns, ne, v) in find_numbers(region, vstart):
                        sev = classify_value(metric, v, th)
                        if sev:
                            sp.append((ns, ne, sev, 2))
                            add_metric(metric, v, sev, key, "modem")
                elif role == "state":
                    val = region.strip()
                    sev = classify_state(val)
                    if sev and val:
                        off0 = vstart + (len(region) - len(region.lstrip()))
                        sp.append((off0, off0 + len(val), sev, 2))
                elif role in ("total", "online", "offline"):
                    m = re.search(r"\d[\d,]*", region)
                    if m and role not in kv_roles:
                        kv_roles[role] = int(m.group().replace(",", ""))
            for m in STATE_WORD_RE.finditer(ln):
                sev = classify_state(m.group())
                if sev:
                    sp.append((m.start(), m.end(), sev, 1))
        all_spans.append(_resolve(sp))

    # ---- summary ---------------------------------------------------------------------------
    summary = None
    best_total = -1
    for t in tables:
        s = _table_summary(t)
        if s and s["total"] > best_total:
            summary, best_total = s, s["total"]
    if summary is None and kv_roles.get("total"):
        tot = kv_roles["total"]
        onl = kv_roles.get("online")
        off = kv_roles.get("offline")
        if onl is None and off is not None:
            onl = tot - off
        if onl is not None:
            summary = {"total": tot, "online": onl, "offline": tot - onl}

    return Parsed(lines, all_spans, tables, metrics, scope, points, summary, errors)


def _table_summary(t: Table) -> Optional[Dict[str, int]]:
    it, io_, ioff = t.col(role="total"), t.col(role="online"), t.col(role="offline")

    def num(row, ci):
        if ci is None:
            return None
        m = re.search(r"-?\d[\d,]*", row[ci])
        return int(m.group().replace(",", "")) if m else None

    if it is None and io_ is None:
        ist, im = t.col(role="state"), t.col(role="mac")
        if ist is not None and im is not None and t.rows:
            sev = [classify_state(r[ist]) for r in t.rows]
            return {"total": len(t.rows), "online": sev.count("ok"), "offline": sev.count("crit")}
        return None
    total_row = next((r for r in t.rows if r and r[0].strip().lower().startswith("total")), None)
    rows = [total_row] if total_row else [r for r in t.rows if not (r and r[0].strip().lower().startswith("total"))]
    tot = onl = off = 0
    got = False
    for r in rows:
        a, b, c = num(r, it), num(r, io_), num(r, ioff)
        if a is None:
            continue
        if b is None and c is not None:
            b = a - c
        if b is None:
            continue
        got = True
        tot += a
        onl += b
        off += (c if c is not None else a - b)
    return {"total": tot, "online": onl, "offline": off} if got and tot > 0 else None


def merge_parsed(items: List[Parsed]) -> Parsed:
    if len(items) == 1:
        return items[0]
    out = Parsed([], [], [], {}, {"node": {}, "modem": {}}, [], None, [])
    for p in items:
        out.tables.extend(p.tables)
        for k, v in p.metrics.items():
            out.metrics.setdefault(k, []).extend(v)
        for sc in ("node", "modem"):
            for k, v in p.scope_metrics[sc].items():
                out.scope_metrics[sc].setdefault(k, []).extend(v)
        out.points.extend(p.points)
        out.errors.extend(p.errors)
        if p.summary and (out.summary is None or p.summary["total"] > out.summary["total"]):
            out.summary = p.summary
    return out


def render_segments(p: Parsed) -> list:
    """Build Text.insert() args: text, tags, text, tags ..."""
    out: list = []
    plain: List[str] = []

    def flush():
        if plain:
            out.append("".join(plain))
            out.append(())
            plain.clear()

    for line, spans in zip(p.lines, p.spans):
        if not spans:
            plain.append(line + "\n")
            continue
        pos = 0
        for s, e, tag in spans:
            if s > pos:
                plain.append(line[pos:s])
            flush()
            out.append(line[s:e])
            out.append((tag,))
            pos = e
        plain.append(line[pos:] + "\n")
    flush()
    return out


# =============================================================================
# Triage / dispatch classification
# =============================================================================
@dataclass
class Finding:
    sev: str
    title: str
    detail: str


@dataclass
class Triage:
    sev: Optional[str]
    findings: List[Finding]
    dispatch: str
    dispatch_sev: Optional[str]
    dispatch_detail: str
    blast: List[Tuple[str, str]]   # (sev, text)
    impaired: int


def run_triage(p: Parsed, cmd: str, th: Dict[str, dict]) -> Triage:
    f: List[Finding] = []

    def counts(metric: str, vals: Optional[List[float]] = None):
        vals = p.metrics.get(metric, []) if vals is None else vals
        return vals, Counter(classify_value(metric, v, th) for v in vals)

    def thr(metric, key):
        return th[metric][key]

    vals, c = counts("us_snr")
    if c["crit"] or c["warn"]:
        sev = "crit" if c["crit"] else "warn"
        f.append(Finding(sev, "Upstream SNR / MER degraded",
                         f"{c['crit']} critical / {c['warn']} marginal readings, minimum {min(vals):.1f} dB "
                         f"(good >= {thr('us_snr', 'good')} dB). Check return path: ingress, loose fittings, pads, laser."))
    vals, c = counts("ds_snr")
    if c["crit"] or c["warn"]:
        f.append(Finding("crit" if c["crit"] else "warn", "Downstream SNR / MER low",
                         f"{c['crit'] + c['warn']} readings below {thr('ds_snr', 'good')} dB, minimum {min(vals):.1f} dB. "
                         "Look for forward-path noise, tilt or a damaged drop."))
    vals, c = counts("us_tx")
    hot = [v for v in vals if v > thr("us_tx", "good")[1]]
    if hot:
        f.append(Finding("crit" if c["crit"] else "warn", "Modems transmitting hot",
                         f"{len(hot)} upstream transmit readings above {thr('us_tx', 'good')[1]} dBmV (max {max(hot):.1f}). "
                         "Excess return-path loss - splitters, drop length or a bad fitting."))
    vals, c = counts("ds_rx")
    if c["crit"] or c["warn"]:
        f.append(Finding("crit" if c["crit"] else "warn", "Downstream receive level out of range",
                         f"{c['crit'] + c['warn']} readings outside {thr('ds_rx', 'good')} dBmV "
                         f"(range {min(vals):.1f} to {max(vals):.1f})."))
    vals, c = counts("cmts_rx")
    if c["crit"] or c["warn"]:
        f.append(Finding("crit" if c["crit"] else "warn", "CMTS upstream receive level off target",
                         f"{c['crit'] + c['warn']} readings outside {thr('cmts_rx', 'good')} dBmV."))
    vals, c = counts("uncorr_pct")
    if c["crit"] or c["warn"]:
        f.append(Finding("crit" if c["crit"] else "warn", "Uncorrectable codewords",
                         f"Peak {max(vals):.2f}% uncorrectable. Impulse noise / ingress likely."))
    vals, c = counts("uncorr")
    if c["crit"] or c["warn"]:
        f.append(Finding("warn", "Uncorrectable codeword counters elevated",
                         f"{sum(vals):,.0f} uncorrectables across {len(vals)} counters (cumulative)."))
    vals, c = counts("flap")
    if c["crit"] or c["warn"]:
        f.append(Finding("crit" if c["crit"] else "warn", "Flapping modems",
                         f"{c['crit'] + c['warn']} entries above {thr('flap', 'good')} flaps, max {max(vals):.0f}."))
    vals, c = counts("util")
    if c["crit"] or c["warn"]:
        f.append(Finding("crit" if c["crit"] else "warn", "Channel utilisation high",
                         f"{c['crit'] + c['warn']} channels above {thr('util', 'good')}% (max {max(vals):.0f}%)."))
    vals, c = counts("temp")
    if c["crit"] or c["warn"]:
        f.append(Finding("crit" if c["crit"] else "warn", "Temperature high", f"Max {max(vals):.0f} C."))
    if p.summary and p.summary["total"]:
        pct = 100.0 * p.summary["online"] / p.summary["total"]
        sev = classify_value("online_pct", pct, th)
        if sev in ("warn", "crit"):
            f.append(Finding(sev, "Modems offline",
                             f"{p.summary['offline']:,} of {p.summary['total']:,} modems not online ({pct:.1f}% online)."))
    if p.errors:
        f.append(Finding("warn", "CLI returned an error", p.errors[0][:160]))

    # blast radius
    is_flap = "flap" in cmd.lower()
    impaired: List[Tuple[str, str]] = []
    modem_rows = 0
    modem_issue = False
    for t in p.tables:
        im = t.col(role="mac")
        if im is None:
            continue
        ii = t.col(role="iface")
        for ri, row in enumerate(t.rows):
            modem_rows += 1
            if t.row_sev[ri] in ("warn", "crit"):
                modem_issue = True
            if is_flap or t.row_sev[ri] == "crit":
                impaired.append((row[im].lower(), row[ii] if ii is not None else ""))
    blast: List[Tuple[str, str]] = []
    concentrated = False
    if impaired:
        n = len(impaired)
        ch = Counter(i for _, i in impaired if i)
        ouis = Counter(lookup_mac_vendor(m) for m, _ in impaired)
        if ch:
            top, cnt = ch.most_common(1)[0]
            share = 100.0 * cnt / n
            if share >= 60 and n >= 3:
                concentrated = True
                blast.append(("crit", f"Shared impairment: {share:.0f}% ({cnt}/{n}) of impaired modems are on {top}. "
                                      "Treat as outside plant on that port - do not roll individual premise trucks."))
            else:
                blast.append(("", "Interface spread: " + ", ".join(f"{k} ({v})" for k, v in ch.most_common(5))))
        top_oui, oc = ouis.most_common(1)[0]
        if n >= 4 and oc / n >= 0.75:
            blast.append(("warn", f"Hardware concentration: {oc}/{n} impaired modems are {top_oui} - check firmware."))
        else:
            blast.append(("", "Vendors: " + ", ".join(f"{k} ({v})" for k, v in ouis.most_common(3))))
        blast.insert(0, ("", f"{n} impaired modem(s) analysed."))

    # ports (non-modem tables) whose online % or RF is critical
    bad_ports: List[str] = []
    for t in p.tables:
        if t.has_modems:
            continue
        ii = t.col(role="iface")
        for ri, row in enumerate(t.rows):
            label = row[ii] if ii is not None else (row[0] if row else "")
            if t.row_sev[ri] == "crit" and label and not label.lower().startswith("total"):
                bad_ports.append(label)
    if bad_ports:
        f.append(Finding("crit", "Critical ports / channels",
                         f"{len(bad_ports)} port(s) below threshold: " + ", ".join(bad_ports[:12]) +
                         (" ..." if len(bad_ports) > 12 else "")))

    node_snr = p.scope_metrics["node"].get("us_snr", [])
    node_crit = bool(bad_ports) or any(classify_value("us_snr", v, th) == "crit" for v in node_snr)
    if len(impaired) >= 5 and not is_flap:
        node_crit = True
    rf_present = any(p.metrics.get(m) for m in ("us_snr", "ds_snr", "us_tx", "ds_rx", "cmts_rx", "uncorr_pct", "uncorr"))
    if p.scope_metrics["modem"]:
        modem_issue = modem_issue or any(
            classify_value(m, v, th) in ("warn", "crit")
            for m, vs in p.scope_metrics["modem"].items() for v in vs)

    if node_crit or concentrated:
        d, ds = "OUTSIDE PLANT", "crit"
        parts = []
        if node_snr:
            parts.append(f"min node US SNR {min(node_snr):.1f} dB")
        if bad_ports:
            parts.append(f"critical ports {', '.join(bad_ports[:4])}")
        if impaired:
            parts.append(f"{len(impaired)} impaired modems")
        dd = f"Shared impairment ({'; '.join(parts) or 'node-level'}). Escalate to plant maintenance / sweep, hold premise dispatch."
    elif modem_issue and modem_rows <= 2:
        d, ds = "PREMISE / DROP", "warn"
        dd = "Single-modem anomaly while node data looks healthy. Inspect drop, ground block, splitters and in-home wiring."
    elif modem_issue:
        d, ds = "ISOLATED MODEMS", "warn"
        dd = (f"{len(impaired)} modem(s) critical with no shared port pattern. Troubleshoot individually - "
              "premise issues are most likely.")
    elif rf_present or p.summary:
        d, ds = "NOMINAL", "ok"
        dd = "Evaluated RF and registration parameters are within thresholds. No truck roll indicated."
    else:
        d, ds, dd = "NO RF DATA", None, "This output has no RF metrics to classify."

    overall = worst(*[x.sev for x in f]) or ("ok" if rf_present or p.summary else None)
    return Triage(overall, f, d, ds, dd, blast, len(impaired))


# =============================================================================
# Flight recorder (SQLite history)
# =============================================================================
class FlightRecorder:
    KEEP_ROWS = 20000

    def __init__(self, db_path: Path = DB_FILE):
        self.db_path = db_path
        try:
            with self._conn() as c:
                c.execute("""CREATE TABLE IF NOT EXISTS audit_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, timestamp TEXT NOT NULL, cmts_host TEXT NOT NULL,
                    command TEXT NOT NULL, output TEXT NOT NULL)""")
                c.execute("CREATE INDEX IF NOT EXISTS idx_host ON audit_logs(cmts_host)")
                c.execute("DELETE FROM audit_logs WHERE id < (SELECT MAX(id) FROM audit_logs) - ?", (self.KEEP_ROWS,))
        except Exception:
            pass

    def _conn(self):
        c = sqlite3.connect(self.db_path, timeout=10)
        c.execute("PRAGMA journal_mode=WAL")
        return c

    def record(self, host: str, cmd: str, out: str) -> None:
        with self._conn() as c:
            c.execute("INSERT INTO audit_logs (timestamp, cmts_host, command, output) VALUES (?,?,?,?)",
                      (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), host, cmd, out))

    def search(self, query: str, limit: int = 200):
        with self._conn() as c:
            if not query.strip():
                return c.execute("SELECT id, timestamp, cmts_host, command FROM audit_logs ORDER BY id DESC LIMIT ?",
                                 (limit,)).fetchall()
            q = f"%{query.strip()}%"
            return c.execute("""SELECT id, timestamp, cmts_host, command FROM audit_logs
                                WHERE cmts_host LIKE ? OR command LIKE ? OR output LIKE ?
                                ORDER BY id DESC LIMIT ?""", (q, q, q, limit)).fetchall()

    def get(self, rec_id: int):
        with self._conn() as c:
            return c.execute("SELECT timestamp, cmts_host, command, output FROM audit_logs WHERE id=?",
                             (rec_id,)).fetchone()

    def recent(self, limit: int = 500):
        with self._conn() as c:
            return c.execute("SELECT timestamp, cmts_host, command, output FROM audit_logs ORDER BY id DESC LIMIT ?",
                             (limit,)).fetchall()


def chronic_offender_report(rows, th) -> list:
    """Returns Text.insert args."""
    mac_hits: Counter = Counter()
    mac_hosts: Dict[str, set] = {}
    iface_hits: Counter = Counter()
    host_snr: Dict[str, List[float]] = {}
    for ts, host, cmd, out in rows:
        if len(out) > 400_000:
            continue
        p = parse_output(out, th)
        snr = p.scope_metrics["node"].get("us_snr") or p.metrics.get("us_snr")
        if snr:
            host_snr.setdefault(host, []).append(sum(snr) / len(snr))
        flap = "flap" in cmd.lower()
        for t in p.tables:
            im, ii = t.col(role="mac"), t.col(role="iface")
            for ri, row in enumerate(t.rows):
                if flap or t.row_sev[ri] == "crit":
                    if im is not None and row[im]:
                        m = row[im].lower()
                        mac_hits[m] += 1
                        mac_hosts.setdefault(m, set()).add(host)
                    if ii is not None and row[ii]:
                        iface_hits[f"{host}  {row[ii]}"] += 1
    a: list = [f"Chronic offender report - {len(rows)} stored outputs analysed\n", ("hdr",),
               f"Generated {datetime.now():%Y-%m-%d %H:%M:%S}\n\n", ("dim",),
               "Recurring impaired modems\n", ("th",)]
    if mac_hits:
        for mac, cnt in mac_hits.most_common(15):
            a += ["  " + mac, ("mac",), f"   hits {cnt:<4} ", ("crit" if cnt >= 3 else "warn",),
                  f"{', '.join(sorted(mac_hosts[mac]))}   {lookup_mac_vendor(mac)}\n", ()]
    else:
        a += ["  none found\n", ("ok",)]
    a += ["\nRecurring impaired interfaces\n", ("th",)]
    if iface_hits:
        for key, cnt in iface_hits.most_common(12):
            a += [f"  {key:<34}", ("iface",), f" appearances {cnt}\n", ("crit" if cnt >= 3 else "warn",)]
    else:
        a += ["  none found\n", ("ok",)]
    a += ["\nNode upstream SNR baselines\n", ("th",)]
    if host_snr:
        for h, s in sorted(host_snr.items()):
            avg = sum(s) / len(s)
            a += [f"  {h:<22} samples {len(s):<4} avg ", (), f"{avg:.1f} dB", (classify_value("us_snr", avg, th) or "",),
                  f"   range {min(s):.1f}-{max(s):.1f} dB (drift {max(s) - min(s):.1f})\n", ()]
    else:
        a += ["  run 'Cable Noise & SNR' to build baselines\n", ("dim",)]
    return a


# =============================================================================
# SSH engine
# =============================================================================
class SessionError(Exception):
    pass


ANSI_RE = re.compile(r"\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])")
CTRL_RE = re.compile(r"[\x00-\x07\x0b\x0c\x0e-\x1f\x7f]")
PAGER_RE = re.compile(r"(?:-+\s*\(?more\)?[^\n]*?-+|<-+\s*more\s*-+>|press any key[^\n]*|\(q\)uit[^\n]*)\s*$", re.I)
CONFIRM_RE = re.compile(r"(\[y(?:es)?/n(?:o)?\]|\(y(?:es)?/n(?:o)?\)|\[confirm\]|continue\?)\s*:?\s*$", re.I)
DEVICE_PROMPT_RE = re.compile(r"^\S{1,80}[#>]\s?$")
SHELL_PROMPT_RE = re.compile(r"^.{0,160}[#>$%]\s?$")
YESNO_RE = re.compile(r"\(yes/no[^)]*\)\??\s*$", re.I)
USER_RE = re.compile(r"(user ?name|login)( as)?\s*:\s*$", re.I)
PASS_RE = re.compile(r"(password|passcode|pass ?phrase)[^\n:]*:\s*$", re.I)
LOGIN_FAIL_RE = re.compile(
    r"permission denied|authentication failed|access denied|login incorrect|login failed|connection refused|"
    r"could not resolve|name or service not known|no route to host|connection timed out|"
    r"host key verification failed|unknown host|bad password|connection closed by", re.I)


def split_host_port(host: str, default: int = 22) -> Tuple[str, int]:
    h = host.strip()
    if h.count(":") == 1:
        a, b = h.split(":")
        if b.isdigit():
            return a, int(b)
    return h, default


def _apply_cr(s: str) -> str:
    out = []
    for line in s.split("\n"):
        line = line.rstrip("\r")
        if "\r" in line:
            line = line.rsplit("\r", 1)[1]
        out.append(line)
    return "\n".join(out)


def _apply_bs(s: str) -> str:
    while "\x08" in s:
        s2 = re.sub(r"[^\x08\n]\x08", "", s)
        if s2 == s:
            break
        s = s2
    return s.replace("\x08", "")


def open_transport(host: str, port: int, user: str, password: str, otp: str, otp_mode: str,
                   timeout: float) -> paramiko.Transport:
    try:
        sock = socket.create_connection((host, port), timeout=timeout)
    except OSError as e:
        raise SessionError(f"Cannot reach {host}:{port} - {e}") from e
    t = paramiko.Transport(sock)
    t.banner_timeout = timeout
    try:
        t.start_client(timeout=timeout)
        if otp and otp_mode == "separate":
            used = {"pw": False}

            def handler(_title, _instr, prompts):
                ans = []
                for p, _echo in prompts:
                    if "pass" in p.lower() and not used["pw"]:
                        ans.append(password)
                        used["pw"] = True
                    else:
                        ans.append(otp)
                return ans

            try:
                t.auth_interactive(user, handler)
            except paramiko.BadAuthenticationType:
                t.auth_password(user, password + otp)
        else:
            t.auth_password(user, password + otp)   # paramiko falls back to keyboard-interactive itself
    except paramiko.AuthenticationException as e:
        t.close()
        raise SessionError(f"Authentication to {host} failed ({e}). Check user, password and token.") from e
    except Exception as e:
        t.close()
        raise SessionError(f"SSH negotiation with {host} failed: {e}") from e
    if not t.is_authenticated():
        t.close()
        raise SessionError(f"Authentication to {host} failed.")
    t.set_keepalive(30)
    return t


class CliChannel:
    """One interactive shell channel with prompt-aware command execution."""

    def __init__(self, chan, host: str, own_transport=None):
        self.chan = chan
        self.host = host
        self.own_transport = own_transport
        self.prompt_re: Optional[re.Pattern] = None
        self.prompt_text = ""
        self.bastion_line = ""
        self.dead = False
        self.lock = threading.Lock()
        self.last_used = time.time()
        self._dec = codecs.getincrementaldecoder("utf-8")("replace")

    @property
    def alive(self) -> bool:
        try:
            tr = self.chan.get_transport()
            return not self.dead and not self.chan.closed and tr is not None and tr.is_active()
        except Exception:
            return False

    def send(self, s: str) -> None:
        self.chan.send(s)

    def close(self) -> None:
        self.dead = True
        try:
            self.chan.close()
        except Exception:
            pass
        if self.own_transport is not None:
            try:
                self.own_transport.close()
            except Exception:
                pass

    def _drain(self, wait: float = 0.0) -> str:
        end = time.monotonic() + wait
        got = ""
        self.chan.settimeout(0.02)
        while True:
            try:
                data = self.chan.recv(65536)
                if not data:
                    break
                got += self._dec.decode(data)
                continue
            except socket.timeout:
                pass
            if time.monotonic() >= end:
                break
        return got

    def read_until(self, done: Callable[[str], Optional[str]], timeout: float, cancel=None,
                   on_progress=None, pager: bool = True) -> Tuple[str, str]:
        buf = ""
        deadline = time.monotonic() + timeout
        self.chan.settimeout(0.05)
        while True:
            if cancel is not None and cancel.is_set():
                try:
                    self.chan.send("\x03")
                except Exception:
                    pass
                return buf, "cancelled"
            try:
                data = self.chan.recv(65536)
                if not data:
                    return buf, "closed"
                chunk = ANSI_RE.sub("", self._dec.decode(data))
                cut = buf.rfind("\n") + 1
                pending = CTRL_RE.sub("", _apply_bs(_apply_cr(buf[cut:] + chunk)))
                buf = buf[:cut] + pending
                if pager:
                    tail_start = max(0, len(buf) - 200)
                    m = PAGER_RE.search(buf, tail_start)
                    if m:
                        buf = buf[:m.start()]
                        self.chan.send(" ")
                        continue
                if on_progress is not None:
                    on_progress(buf)
                r = done(buf)
                if r:
                    return buf, r
            except socket.timeout:
                pass
            if self.chan.closed:
                return buf, "closed"
            if time.monotonic() > deadline:
                return buf, "timeout"

    # ---- login -------------------------------------------------------------------------------
    def login(self, first_cmd: Optional[str], user: str, password: str, timeout: float,
              is_target: Callable[[str], bool]) -> str:
        if first_cmd:
            self._drain()
            self.send(first_cmd + "\n")
        deadline = time.monotonic() + timeout
        pw_count = user_count = 0
        nudged = False
        transcript = ""

        def state(buf: str) -> Optional[str]:
            last = buf.rsplit("\n", 1)[-1]
            if LOGIN_FAIL_RE.search(buf):
                return "fail"
            if YESNO_RE.search(last):
                return "yesno"
            if PASS_RE.search(last):
                return "password"
            if USER_RE.search(last):
                return "user"
            if is_target(last):
                return "prompt"
            return None

        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise SessionError(f"Login to {self.host} timed out. Last output: {transcript[-200:].strip()!r}")
            buf, why = self.read_until(state, min(remaining, 5.0))
            transcript += buf
            if why == "fail":
                buf += _apply_cr(self._drain(0.4))
                line = next((ln for ln in buf.splitlines() if LOGIN_FAIL_RE.search(ln)), buf[-160:])
                raise SessionError(f"{self.host}: {line.strip()}")
            if why == "yesno":
                self.send("yes\n")
            elif why == "user":
                user_count += 1
                if user_count > 3:
                    raise SessionError(f"{self.host}: username prompt repeated - login rejected.")
                self.send(user + "\n")
            elif why == "password":
                pw_count += 1
                if pw_count > 2:
                    raise SessionError(f"{self.host}: password rejected.")
                self.send(password + "\n")
            elif why == "prompt":
                return transcript
            elif why == "closed":
                raise SessionError(f"{self.host}: connection closed during login. {transcript[-160:].strip()}")
            elif why == "timeout" and not nudged:
                nudged = True
                self.send("\n")

    def learn_prompt(self) -> str:
        self._drain(0.2)
        self.send("\n")
        buf, _ = self.read_until(
            lambda b: "p" if "\n" in b and DEVICE_PROMPT_RE.match(b.rsplit("\n", 1)[-1]) else None, 10)
        last = buf.rsplit("\n", 1)[-1].strip()
        if not last:
            raise SessionError(f"{self.host}: could not detect the CLI prompt.")
        base = re.sub(r"(\([^)]*\))?[#>]$", "", last)
        self.prompt_text = last
        self.prompt_re = re.compile("^" + re.escape(base) + r"(?:\([^)\n]*\))?[#>]\s?$")
        return last

    # ---- command execution -----------------------------------------------------------------------
    def run(self, cmd: str, timeout: float = 120, cancel=None, on_progress=None) -> Tuple[str, str]:
        with self.lock:
            self.last_used = time.time()
            self._drain()

            def done(buf: str) -> Optional[str]:
                if "\n" not in buf:
                    return None
                last = buf.rsplit("\n", 1)[-1]
                if (self.prompt_re.match(last) if self.prompt_re else DEVICE_PROMPT_RE.match(last)):
                    return "prompt"
                if self.bastion_line and last.strip() == self.bastion_line:
                    return "exited"
                if CONFIRM_RE.search(last):
                    return "confirm"
                return None

            self.send(cmd + "\n")
            buf, why = self.read_until(done, timeout, cancel, on_progress)
            if why == "cancelled":
                more, why2 = self.read_until(done, 6)
                buf += more
                why = "cancelled" if why2 == "prompt" else why2
            if why in ("exited", "closed"):
                self.dead = True
            lines = buf.split("\n")
            if why in ("prompt", "exited", "cancelled") and lines:
                tail = lines[-1].strip()
                if self.prompt_re and self.prompt_re.match(tail):
                    self.prompt_text = tail
                    lines = lines[:-1]
                elif why == "exited":
                    lines = lines[:-1]
            if lines and cmd.strip() and lines[0].strip().endswith(cmd.strip()[-60:]):
                lines = lines[1:]
            return "\n".join(lines).strip("\n"), why


class SessionManager:
    MAX_CACHED = 4

    def __init__(self, log: Callable[[str], None]):
        self.log = log
        self.transport: Optional[paramiko.Transport] = None
        self.bastion_key: Optional[Tuple[str, str]] = None
        self.channels: Dict[str, CliChannel] = {}
        self.lock = threading.RLock()

    def bastion_ready(self, bastion: str, user: str) -> bool:
        return (self.transport is not None and self.transport.is_active()
                and self.bastion_key == (bastion.lower(), user))

    def get(self, host: str) -> Optional[CliChannel]:
        with self.lock:
            cli = self.channels.get(host.lower())
            if cli and cli.alive:
                return cli
            if cli:
                self.channels.pop(host.lower(), None)
                cli.close()
            return None

    def needs_otp(self, bastion: str, user: str, host: str) -> bool:
        return bool(bastion) and not self.bastion_ready(bastion, user) and self.get(host) is None

    def connect_bastion(self, bastion: str, user: str, password: str, otp: str, st: dict) -> None:
        self.close_all()
        h, port = split_host_port(bastion)
        self.log(f"Authenticating to bastion {h} as {user} ...")
        t = open_transport(h, port, user, password, otp, str(st["otp_mode"]), float(st["login_timeout"]))
        with self.lock:
            self.transport = t
            self.bastion_key = (bastion.lower(), user)
        self.log(f"Bastion {h} authenticated. The session is kept open - no new token needed until it drops.")

    @staticmethod
    def _shell(transport) -> paramiko.Channel:
        ch = transport.open_session(timeout=20)
        ch.get_pty(term="vt100", width=511, height=1000)
        ch.invoke_shell()
        return ch

    def open_cmts(self, bastion: str, user: str, password: str, otp: str, host: str, st: dict,
                  cache: bool = True) -> CliChannel:
        timeout = float(st["login_timeout"])
        if bastion:
            if not self.bastion_ready(bastion, user):
                raise SessionError("Bastion session is not open. Run the command again and enter a new VIP token.")
            try:
                ch = self._shell(self.transport)
            except Exception as e:
                raise SessionError(f"Bastion refused a new channel ({e}). It may limit concurrent sessions.") from e
            cli = CliChannel(ch, host)
            try:
                cli.login(None, user, password + otp, 20, lambda last: bool(SHELL_PROMPT_RE.match(last)))
                cli._drain(0.2)
                cli.send("\n")
                buf, _ = cli.read_until(lambda b: "p" if "\n" in b and SHELL_PROMPT_RE.match(b.rsplit("\n", 1)[-1]) else None, 6)
                cli.bastion_line = buf.rsplit("\n", 1)[-1].strip()
                hop = str(st["hop_cmd"]).format(host=host, user=user)
                self.log(f"Hopping to {host} ...")
                cli.login(hop, user, password, timeout,
                          lambda last: bool(DEVICE_PROMPT_RE.match(last)) and last.strip() != cli.bastion_line)
            except Exception:
                cli.close()
                raise
        else:
            h, port = split_host_port(host)
            self.log(f"Connecting directly to {h} ...")
            t = open_transport(h, port, user, password, "", "append", timeout)
            cli = CliChannel(self._shell(t), host, own_transport=t)
            try:
                cli.login(None, user, password, timeout, lambda last: bool(DEVICE_PROMPT_RE.match(last)))
            except Exception:
                cli.close()
                raise
        cli.learn_prompt()
        for pc in st.get("paging_cmds") or []:
            if pc.strip():
                cli.run(pc.strip(), timeout=8)
        if cache:
            with self.lock:
                old = self.channels.pop(host.lower(), None)
                if old:
                    old.close()
                self.channels[host.lower()] = cli
                while len(self.channels) > self.MAX_CACHED:
                    lru = min(self.channels, key=lambda k: self.channels[k].last_used)
                    self.channels.pop(lru).close()
        return cli

    def drop(self, host: str) -> None:
        with self.lock:
            cli = self.channels.pop(host.lower(), None)
        if cli:
            cli.close()

    def close_all(self) -> None:
        with self.lock:
            chans = list(self.channels.values())
            self.channels.clear()
            t, self.transport, self.bastion_key = self.transport, None, None
        for c in chans:
            c.close()
        if t is not None:
            try:
                t.close()
            except Exception:
                pass


# =============================================================================
# Data holders
# =============================================================================
@dataclass
class Creds:
    bastion: str
    user: str
    password: str
    host: str


@dataclass
class Result:
    host: str
    cmd: str
    output: str
    parsed: Parsed
    duration: float
    status: str
    ts: datetime


# =============================================================================
# Autocomplete popup
# =============================================================================
class AutocompletePopup(tk.Toplevel):
    def __init__(self, master, P: dict, on_pick: Callable[[str], None]):
        super().__init__(master)
        self.wm_overrideredirect(True)
        self.attributes("-topmost", True)
        self.on_pick = on_pick
        self.items: List[str] = []
        self.configure(bg=P["border"])
        self.lb = tk.Listbox(self, font=(MONO_FONT, 11), bg=P["card"], fg=P["text"], selectbackground=P["accent"],
                             selectforeground=P["on_accent"], bd=0, highlightthickness=0, activestyle="none")
        self.lb.pack(fill="both", expand=True, padx=1, pady=1)
        self.lb.bind("<ButtonRelease-1>", lambda e: self._pick())
        self.withdraw()

    @property
    def visible(self) -> bool:
        return bool(self.winfo_ismapped())

    def show(self, items: List[str], x: int, y_bottom: int, width: int):
        if not items:
            self.withdraw()
            return
        self.items = items[:8]
        self.lb.delete(0, "end")
        for it in self.items:
            self.lb.insert("end", "  " + it)
        self.lb.selection_set(0)
        h = len(self.items) * 24 + 6
        self.geometry(f"{max(width, 420)}x{h}+{x}+{y_bottom - h - 4}")
        self.deiconify()
        self.lift()

    def move(self, d: int):
        cur = self.lb.curselection()
        i = ((cur[0] if cur else 0) + d) % max(1, len(self.items))
        self.lb.selection_clear(0, "end")
        self.lb.selection_set(i)
        self.lb.see(i)

    def current(self) -> Optional[str]:
        cur = self.lb.curselection()
        return self.items[cur[0]] if cur else None

    def _pick(self):
        v = self.current()
        if v:
            self.on_pick(v)
        self.withdraw()


# =============================================================================
# Main application
# =============================================================================
TABS = ["Terminal", "Table", "Charts", "Triage", "Fleet", "Diff", "History", "Log"]
LOGIN_HINT = "Leave Bastion empty to SSH straight to the CMTS."


class CommandCenter(ctk.CTk):
    def __init__(self):
        _migrate_legacy_files()
        super().__init__()
        self.cfg = load_config()
        self.st: dict = {**DEFAULT_SETTINGS, **{k: v for k, v in self.cfg.get("settings", {}).items()
                                                if k in DEFAULT_SETTINGS}}
        if "theme" in self.cfg and self.cfg["theme"] in THEMES:
            self.st["theme"] = self.cfg["theme"]
        self.th: Dict[str, dict] = json.loads(json.dumps(DEFAULT_THRESHOLDS))
        for k, v in (self.cfg.get("thresholds") or {}).items():
            if k in self.th and isinstance(v, dict):
                self.th[k].update({kk: vv for kk, vv in v.items() if kk in ("good", "warn")})
        self.P = THEMES.get(str(self.st["theme"]), THEMES["Midnight"])
        ctk.set_appearance_mode(self.P["mode"])
        self.title(f"{APP_NAME}  {VERSION}")
        try:
            ico = resource_path("assets/app.ico")
            if ico.is_file():
                self.iconbitmap(default=str(ico))
        except Exception:
            pass
        self.geometry("1520x930")
        self.minsize(1180, 720)

        # state
        self.sessions = SessionManager(self.log)
        self.flight = FlightRecorder(DB_FILE)
        self.jobs: "queue.Queue[Optional[Callable]]" = queue.Queue()
        self._uiq: "queue.SimpleQueue" = queue.SimpleQueue()
        self.cancel_event = threading.Event()
        self.busy = 0
        self.results: List[Result] = []
        self.last_parsed: Optional[Parsed] = None
        self.last_triage: Optional[Triage] = None
        self.term_blocks: deque = deque(maxlen=60)
        self.log_lines: deque = deque(maxlen=3000)
        self.kpi_state: Dict[str, Tuple[str, str, Optional[str]]] = {}
        self.trend: Dict[str, deque] = {}
        self.fleet_results: Dict[str, Result] = {}
        self.fleet_rows: List[dict] = []
        self.diff_args: list = []
        self.snapshot: Optional[dict] = None
        self.triage_args: list = []
        self.cmd_history: List[str] = list(self.cfg.get("history", []))[-200:]
        self.hist_idx = len(self.cmd_history)
        self.watching = False
        self.watch_cmds: List[str] = []
        self.watch_cycle = 0
        self.watch_interval = 15
        self.connected_host = ""
        self.fig = None
        self.canvas = None
        self.charts_dirty = True
        self.tbl_rows: list = []
        self.tbl_table: Optional[Table] = None
        self.tbl_sort: Tuple[int, bool] = (-1, False)
        self.find_hits: List[Tuple[str, str]] = []
        self.find_idx = -1
        self.current_tab = "Terminal"
        self.skip_confirm = False

        self._build_ui()
        threading.Thread(target=self._job_loop, daemon=True).start()
        self.after(25, self._drain_ui)
        self.after(4000, self._monitor_connection)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.log(f"{APP_NAME} {VERSION} ready. Config: {APP_DIR}")

    # ------------------------------------------------------------------ threading helpers
    def ui(self, fn: Callable, *args) -> None:
        self._uiq.put((fn, args))

    def _drain_ui(self) -> None:
        t0 = time.perf_counter()
        try:
            while time.perf_counter() - t0 < 0.06:
                fn, args = self._uiq.get_nowait()
                try:
                    fn(*args)
                except Exception as e:  # keep the pump alive
                    self._log_sync(f"UI error: {e!r}")
        except queue.Empty:
            pass
        self.after(15, self._drain_ui)

    def _job_loop(self) -> None:
        while True:
            fn = self.jobs.get()
            if fn is None:
                return
            try:
                fn()
            except SessionError as e:
                self.log(f"Connection error: {e}")
                self.ui(self._job_error, str(e))
            except Exception as e:
                self.log(f"Error: {e!r}")
                self.ui(self._job_error, f"{type(e).__name__}: {e}")
            finally:
                self.ui(self._job_finished)

    def _enqueue(self, fn: Callable) -> None:
        self.busy += 1
        self.cancel_event.clear()
        self._set_busy_ui()
        self.jobs.put(fn)

    def _job_finished(self) -> None:
        self.busy = max(0, self.busy - 1)
        self._set_busy_ui()

    def _job_error(self, msg: str) -> None:
        if self.watching:
            self._stop_watch()
        self._set_activity(msg, "crit")
        messagebox.showerror("Command Center", msg, parent=self)

    def log(self, msg: str) -> None:
        self.ui(self._log_sync, msg)

    def _log_sync(self, msg: str) -> None:
        line = f"[{datetime.now():%H:%M:%S}] {msg}"
        self.log_lines.append(line)
        if hasattr(self, "log_text"):
            self.log_text.insert("end", line + "\n")
            self.log_text.see("end")

    # ------------------------------------------------------------------ widget helpers
    def F(self, size: int = 12, weight: str = "normal", mono: bool = False):
        return ctk.CTkFont(family=MONO_FONT if mono else UI_FONT, size=size, weight=weight)

    def _btn(self, parent, text, cmd, kind="secondary", width=0, height=30, **kw):
        P = self.P
        styles = {
            "primary": dict(fg_color=P["accent"], hover_color=P["accent_hover"], text_color=P["on_accent"]),
            "secondary": dict(fg_color=P["hover"], hover_color=P["border"], text_color=P["text"]),
            "ghost": dict(fg_color="transparent", hover_color=P["hover"], text_color=P["muted"]),
            "danger": dict(fg_color=P["hover"], hover_color=P["crit"], text_color=P["crit"]),
        }
        opts = dict(text=text, command=cmd, height=height, corner_radius=7, font=self.F(12, "bold" if kind == "primary" else "normal"),
                    text_color_disabled=P["dim"], **styles[kind])
        if width:
            opts["width"] = width
        opts.update(kw)
        return ctk.CTkButton(parent, **opts)

    def _entry(self, parent, placeholder="", show=None, mono=False, height=30, **kw):
        P = self.P
        return ctk.CTkEntry(parent, placeholder_text=placeholder, show=show, height=height, corner_radius=7,
                            fg_color=P["input"], border_color=P["border"], border_width=1, text_color=P["text"],
                            placeholder_text_color=P["dim"], font=self.F(12, mono=mono), **kw)

    def _option(self, parent, values, cmd=None, width=0, **kw):
        P = self.P
        opts = dict(values=values, command=cmd, height=30, corner_radius=7, fg_color=P["hover"],
                    button_color=P["hover"], button_hover_color=P["border"], text_color=P["text"],
                    dropdown_fg_color=P["card"], dropdown_hover_color=P["hover"], dropdown_text_color=P["text"],
                    font=self.F(12), dropdown_font=self.F(12), dynamic_resizing=False)
        if width:
            opts["width"] = width
        opts.update(kw)
        return ctk.CTkOptionMenu(parent, **opts)

    def _label(self, parent, text, size=12, weight="normal", color=None, **kw):
        return ctk.CTkLabel(parent, text=text, font=self.F(size, weight), text_color=color or self.P["text"], **kw)

    def _section(self, parent, title):
        P = self.P
        card = ctk.CTkFrame(parent, fg_color=P["card"], corner_radius=10, border_width=1, border_color=P["border"])
        card.pack(fill="x", padx=10, pady=(0, 10))
        self._label(card, title.upper(), 10, "bold", P["muted"]).pack(anchor="w", padx=12, pady=(9, 4))
        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="x", padx=12, pady=(0, 11))
        return body

    def _field(self, parent, label, **kw):
        self._label(parent, label, 11, color=self.P["muted"]).pack(anchor="w", pady=(4, 1))
        e = self._entry(parent, **kw)
        e.pack(fill="x")
        return e

    def _make_text(self, parent, wrap="none", mono=True, size=11, readonly=True):
        P = self.P
        frame = ctk.CTkFrame(parent, fg_color=P["term_bg"], corner_radius=8, border_width=1, border_color=P["border"])
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        t = tk.Text(frame, wrap=wrap, bg=P["term_bg"], fg=P["term_fg"], insertbackground=P["text"],
                    selectbackground=P["sel"], selectforeground=P["text"], font=(MONO_FONT if mono else UI_FONT, size),
                    bd=0, highlightthickness=0, padx=10, pady=8, undo=False)
        ys = ctk.CTkScrollbar(frame, command=t.yview, button_color=P["border"], button_hover_color=P["muted"])
        xs = ctk.CTkScrollbar(frame, command=t.xview, orientation="horizontal", button_color=P["border"],
                              button_hover_color=P["muted"])
        t.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        t.grid(row=0, column=0, sticky="nsew", padx=(2, 0), pady=(2, 0))
        ys.grid(row=0, column=1, sticky="ns", pady=4)
        if wrap == "none":
            xs.grid(row=1, column=0, sticky="ew", padx=4)
        self._config_tags(t, size)
        if readonly:
            nav = {"Left", "Right", "Up", "Down", "Prior", "Next", "Home", "End"}
            t.bind("<Key>", lambda e: None if (e.state & 0x4 or e.keysym in nav) else "break")
            t.bind("<<Paste>>", lambda e: "break")
            t.bind("<<Cut>>", lambda e: "break")
        return frame, t

    def _config_tags(self, t: tk.Text, size: int = 11) -> None:
        P = self.P
        bold = (MONO_FONT, size, "bold")
        t.tag_configure("ok", foreground=P["ok"])
        t.tag_configure("warn", foreground=P["warn"], font=bold)
        t.tag_configure("crit", foreground=P["crit"], font=bold)
        t.tag_configure("mac", foreground=P["mac"])
        t.tag_configure("ip", foreground=P["ip"])
        t.tag_configure("iface", foreground=P["iface"])
        t.tag_configure("th", foreground=P["th"], font=bold)
        t.tag_configure("dim", foreground=P["dim"])
        t.tag_configure("prompt", foreground=P["muted"], font=bold)
        t.tag_configure("cmd", foreground=P["text"], font=bold)
        t.tag_configure("hdr", foreground=P["accent"], font=bold)
        t.tag_configure("add", foreground=P["ok"])
        t.tag_configure("del", foreground=P["crit"])
        t.tag_configure("hit", background=P["hit"])
        t.tag_configure("hit_cur", background=P["hit_cur"], foreground="#000000")
        t.tag_raise("hit")
        t.tag_raise("hit_cur")
        t.tag_raise("sel")

    # ------------------------------------------------------------------ UI construction
    def _build_ui(self) -> None:
        P = self.P
        self.configure(fg_color=P["bg"])
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self._build_topbar()
        self._build_sidebar()
        self._build_main()
        self._build_statusbar()
        self.popup = AutocompletePopup(self, P, self._apply_autocomplete)
        self._setup_tree_style()
        self.bind_all("<Control-f>", lambda e: self._open_find())
        self.bind_all("<Control-l>", lambda e: self.clear_terminal())
        self._show_tab(self.current_tab)

    def _build_topbar(self) -> None:
        P = self.P
        bar = ctk.CTkFrame(self, fg_color=P["panel"], corner_radius=0, height=58, border_width=0)
        bar.grid(row=0, column=0, columnspan=2, sticky="ew")
        bar.grid_columnconfigure(2, weight=1)
        left = ctk.CTkFrame(bar, fg_color="transparent")
        left.grid(row=0, column=0, padx=(18, 10), pady=10, sticky="w")
        self._label(left, "E6000", 17, "bold", P["accent"]).pack(side="left")
        self._label(left, " Command Center", 17, "bold").pack(side="left")
        self._label(left, f"  v{VERSION}", 11, color=P["dim"]).pack(side="left", pady=(4, 0))

        tgt = ctk.CTkFrame(bar, fg_color="transparent")
        tgt.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        self._label(tgt, "CMTS", 11, "bold", P["muted"]).pack(side="left", padx=(0, 8))
        self.host_entry = self._entry(tgt, "hostname or IP", mono=True, width=240)
        self.host_entry.pack(side="left")
        self.host_entry.bind("<Return>", lambda e: self.connect())
        if self.cfg.get("last_host"):
            self.host_entry.insert(0, self.cfg["last_host"])
        self._btn(tgt, "Connect", self.connect, "primary", width=92).pack(side="left", padx=(8, 4))
        self._btn(tgt, "Disconnect", self.disconnect, "secondary", width=96).pack(side="left")

        right = ctk.CTkFrame(bar, fg_color="transparent")
        right.grid(row=0, column=3, padx=(10, 16), pady=10, sticky="e")
        pill = ctk.CTkFrame(right, fg_color=P["card"], corner_radius=14, border_width=1, border_color=P["border"])
        pill.pack(side="left", padx=(0, 10))
        self.conn_dot = self._label(pill, "●", 13, color=P["dim"])
        self.conn_dot.pack(side="left", padx=(10, 4), pady=3)
        self.conn_lbl = self._label(pill, "Not connected", 12, color=P["muted"])
        self.conn_lbl.pack(side="left", padx=(0, 12), pady=3)
        self._btn(right, "Settings", self.open_settings, "secondary", width=84).pack(side="left")

    def _build_sidebar(self) -> None:
        P = self.P
        sb = ctk.CTkScrollableFrame(self, width=300, corner_radius=0, fg_color=P["panel"],
                                    scrollbar_button_color=P["border"], scrollbar_button_hover_color=P["muted"])
        sb.grid(row=1, column=0, sticky="nsw")
        ctk.CTkFrame(sb, fg_color="transparent", height=6).pack()

        # Access
        body = self._section(sb, "Access")
        self.bastion_entry = self._field(body, "Bastion host", placeholder="jump.example.net (optional)")
        self.user_entry = self._field(body, "Username (P-number)", placeholder="P1234567")
        self.pw_entry = self._field(body, "Password", show="•")
        self._label(body, LOGIN_HINT, 10, color=P["dim"], wraplength=260, justify="left").pack(anchor="w", pady=(6, 0))
        if self.cfg.get("bastion_host"):
            self.bastion_entry.insert(0, self.cfg["bastion_host"])
        if self.cfg.get("p_num"):
            self.user_entry.insert(0, self.cfg["p_num"])

        # Parameters
        body = self._section(sb, "Parameters")
        row = ctk.CTkFrame(body, fg_color="transparent")
        row.pack(fill="x")
        row.grid_columnconfigure((0, 1), weight=1)
        c1 = ctk.CTkFrame(row, fg_color="transparent")
        c1.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        c2 = ctk.CTkFrame(row, fg_color="transparent")
        c2.grid(row=0, column=1, sticky="ew", padx=(4, 0))
        self.cmac_entry = self._field(c1, "Cable-Mac", placeholder="1", mono=True)
        self.iface_entry = self._field(c2, "Interface", placeholder="12/0/48", mono=True)
        self.mac_entry = self._field(body, "Modem MAC", placeholder="0000.0000.0000", mono=True)

        # Quick commands
        body = self._section(sb, "Quick commands")
        self.quick_menu = self._option(body, list(COMMAND_CATALOG), self._on_quick_select)
        self.quick_menu.pack(fill="x")
        self.quick_desc = self._label(body, "", 11, color=P["muted"], wraplength=262, justify="left")
        self.quick_desc.pack(anchor="w", pady=(6, 8))
        r = ctk.CTkFrame(body, fg_color="transparent")
        r.pack(fill="x")
        r.grid_columnconfigure(0, weight=1)
        self.run_quick_btn = self._btn(r, "Run", self.run_quick, "primary")
        self.run_quick_btn.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self._btn(r, "Edit in CLI", self.quick_to_cli, "secondary", width=96).grid(row=0, column=1)
        self._on_quick_select(self.quick_menu.get())

        # Playbooks
        body = self._section(sb, "Playbooks")
        self.pb_menu = self._option(body, list(PLAYBOOK_CATALOG), self._on_pb_select)
        self.pb_menu.pack(fill="x")
        self.pb_desc = self._label(body, "", 11, color=P["muted"], wraplength=262, justify="left")
        self.pb_desc.pack(anchor="w", pady=(6, 8))
        self.run_pb_btn = self._btn(body, "Run playbook", self.run_playbook, "secondary")
        self.run_pb_btn.pack(fill="x")
        self._on_pb_select(self.pb_menu.get())

        # Live watch
        body = self._section(sb, "Live watch")
        self._label(body, "Re-runs the last command on an interval.", 11, color=P["muted"]).pack(anchor="w", pady=(0, 6))
        r = ctk.CTkFrame(body, fg_color="transparent")
        r.pack(fill="x")
        r.grid_columnconfigure(0, weight=1)
        self.watch_btn = self._btn(r, "Start watch", self.toggle_watch, "secondary")
        self.watch_btn.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.watch_menu = self._option(r, ["5 s", "15 s", "30 s", "60 s", "5 min"], self._on_watch_interval, width=90)
        self.watch_menu.set({5: "5 s", 15: "15 s", 30: "30 s", 60: "60 s", 300: "5 min"}.get(self.watch_interval, "15 s"))
        self.watch_menu.grid(row=0, column=1)
        self.watch_lbl = self._label(body, "Idle", 11, color=P["dim"])
        self.watch_lbl.pack(anchor="w", pady=(6, 0))

        # Maintenance
        body = self._section(sb, "Maintenance window")
        r = ctk.CTkFrame(body, fg_color="transparent")
        r.pack(fill="x")
        r.grid_columnconfigure((0, 1), weight=1)
        self._btn(r, "Snapshot (pre)", self.take_snapshot, "secondary").grid(row=0, column=0, sticky="ew", padx=(0, 4))
        self._btn(r, "Compare (post)", self.run_diff, "secondary").grid(row=0, column=1, sticky="ew", padx=(4, 0))
        self.snap_lbl = self._label(body, "No baseline captured", 11, color=P["dim"])
        self.snap_lbl.pack(anchor="w", pady=(6, 0))
        if self.snapshot:
            self.snap_lbl.configure(text=f"Baseline: {self.snapshot['host']} @ {self.snapshot['ts']}", text_color=P["ok"])

        # Fleet
        body = self._section(sb, "Fleet audit")
        self._label(body, "CMTS hosts (comma / newline separated)", 11, color=P["muted"]).pack(anchor="w")
        self.fleet_box = ctk.CTkTextbox(body, height=70, fg_color=P["input"], border_color=P["border"], border_width=1,
                                        text_color=P["text"], font=self.F(11, mono=True), corner_radius=7)
        self.fleet_box.pack(fill="x", pady=(2, 6))
        if self.cfg.get("fleet_hosts"):
            self.fleet_box.insert("1.0", self.cfg["fleet_hosts"])
        self.fleet_cmd = self._field(body, "Command", mono=True)
        self.fleet_cmd.insert(0, self.cfg.get("fleet_cmd", "show cable modem summary"))
        self.fleet_btn = self._btn(body, "Run on fleet", self.run_fleet, "secondary")
        self.fleet_btn.pack(fill="x", pady=(8, 0))

    def _build_main(self) -> None:
        P = self.P
        main = ctk.CTkFrame(self, fg_color=P["bg"], corner_radius=0)
        main.grid(row=1, column=1, sticky="nsew", padx=(12, 14), pady=(12, 8))
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(2, weight=1)

        # KPI strip
        kpi = ctk.CTkFrame(main, fg_color="transparent")
        kpi.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self.kpi: Dict[str, Tuple] = {}
        defs = [("online", "Online"), ("offline", "Offline"), ("us_snr", "Min US SNR"), ("ds_snr", "Min DS SNR"),
                ("cw", "Codewords"), ("dispatch", "Dispatch")]
        for i, (key, title) in enumerate(defs):
            kpi.grid_columnconfigure(i, weight=1, uniform="kpi")
            card = ctk.CTkFrame(kpi, fg_color=P["card"], corner_radius=10, border_width=1, border_color=P["border"])
            card.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 5, 0 if i == len(defs) - 1 else 5))
            self._label(card, title.upper(), 10, "bold", P["muted"]).pack(anchor="w", padx=12, pady=(9, 0))
            v = self._label(card, "—", 20, "bold", P["dim"])
            v.pack(anchor="w", padx=12)
            s = self._label(card, " ", 10, color=P["dim"])
            s.pack(anchor="w", padx=12, pady=(0, 8))
            self.kpi[key] = (v, s)
        for key, (val, sub, sev) in self.kpi_state.items():
            self._set_kpi(key, val, sub, sev)

        # Tab bar + actions
        tb = ctk.CTkFrame(main, fg_color="transparent")
        tb.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        tb.grid_columnconfigure(1, weight=1)
        self.tabbar = ctk.CTkSegmentedButton(
            tb, values=TABS, command=self._show_tab, height=32, corner_radius=8, font=self.F(12, "bold"),
            fg_color=P["card"], selected_color=P["accent"], selected_hover_color=P["accent_hover"],
            unselected_color=P["card"], unselected_hover_color=P["hover"], text_color=P["text"])
        self.tabbar.grid(row=0, column=0, sticky="w")
        acts = ctk.CTkFrame(tb, fg_color="transparent")
        acts.grid(row=0, column=2, sticky="e")
        for text, cmd in (("Find", self._open_find), ("Copy", self.copy_output), ("Ticket notes", self.copy_ticket),
                          ("HTML report", self.export_html), ("Excel report", self.export_excel),
                          ("Output → Excel", self.export_output_excel), ("Clear", self.clear_terminal)):
            self._btn(acts, text, cmd, "ghost", height=30).pack(side="left", padx=1)

        # View stack
        stack = ctk.CTkFrame(main, fg_color="transparent")
        stack.grid(row=2, column=0, sticky="nsew")
        stack.grid_rowconfigure(0, weight=1)
        stack.grid_columnconfigure(0, weight=1)
        self.views: Dict[str, ctk.CTkFrame] = {}
        for name in TABS:
            f = ctk.CTkFrame(stack, fg_color="transparent")
            f.grid(row=0, column=0, sticky="nsew")
            f.grid_rowconfigure(1, weight=1)
            f.grid_columnconfigure(0, weight=1)
            self.views[name] = f
        self._build_terminal_view(self.views["Terminal"])
        self._build_table_view(self.views["Table"])
        self._build_charts_view(self.views["Charts"])
        self._build_simple_text_view("Triage", "triage_text", "Run a command or playbook to see automated triage.")
        self._build_fleet_view(self.views["Fleet"])
        self._build_simple_text_view("Diff", "diff_text", "Capture a pre-maintenance snapshot, then Compare (post).")
        self._build_history_view(self.views["History"])
        self._build_simple_text_view("Log", "log_text", "")
        for line in self.log_lines:
            self.log_text.insert("end", line + "\n")
        if self.triage_args:
            self._write_args(self.triage_text, self.triage_args)
        if self.diff_args:
            self._write_args(self.diff_text, self.diff_args)

        # CLI dock
        dock = ctk.CTkFrame(main, fg_color=P["card"], corner_radius=10, border_width=1, border_color=P["border"])
        dock.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        dock.grid_columnconfigure(1, weight=1)
        self.prompt_lbl = self._label(dock, "not connected >", 13, "bold", P["dim"])
        self.prompt_lbl.configure(font=self.F(13, "bold", mono=True))
        self.prompt_lbl.grid(row=0, column=0, padx=(14, 8), pady=9)
        self.cli = self._entry(dock, "Type any CLI command  -  Tab completes, Up/Down history, paste several lines to send in order",
                               mono=True, height=36)
        self.cli.grid(row=0, column=1, sticky="ew", pady=9)
        self.cli.bind("<Return>", lambda e: self.send_cli())
        self.cli.bind("<KeyRelease>", self._on_cli_key)
        self.cli.bind("<Up>", lambda e: self._cli_nav(-1))
        self.cli.bind("<Down>", lambda e: self._cli_nav(1))
        self.cli.bind("<Tab>", self._on_cli_tab)
        self.cli.bind("<Escape>", lambda e: self.popup.withdraw())
        self.cli.bind("<FocusOut>", lambda e: self.after(150, self.popup.withdraw))
        self.send_btn = self._btn(dock, "Send", self.send_cli, "primary", width=80, height=36)
        self.send_btn.grid(row=0, column=2, padx=(8, 4), pady=9)
        self.stop_btn = self._btn(dock, "Stop", self.stop, "danger", width=70, height=36, state="disabled")
        self.stop_btn.grid(row=0, column=3, padx=(4, 12), pady=9)
        if self.connected_host:
            self._on_connected(self.connected_host, self._prompt_cache)

    def _build_statusbar(self) -> None:
        P = self.P
        bar = ctk.CTkFrame(self, fg_color=P["panel"], corner_radius=0, height=28)
        bar.grid(row=2, column=0, columnspan=2, sticky="ew")
        bar.grid_columnconfigure(0, weight=1)
        self.activity = self._label(bar, "Ready", 11, color=P["muted"], anchor="w")
        self.activity.grid(row=0, column=0, sticky="ew", padx=14, pady=3)
        leg = ctk.CTkFrame(bar, fg_color="transparent")
        leg.grid(row=0, column=1, padx=14)
        for text, key in (("OK", "ok"), ("Marginal", "warn"), ("Critical", "crit"), ("MAC", "mac"), ("IP", "ip"),
                          ("Port", "iface")):
            self._label(leg, "●", 11, color=P[key]).pack(side="left", padx=(8, 2))
            self._label(leg, text, 11, color=P["muted"]).pack(side="left")

    def _build_terminal_view(self, f) -> None:
        P = self.P
        self.find_bar = ctk.CTkFrame(f, fg_color=P["card"], corner_radius=8, border_width=1, border_color=P["border"])
        self.find_bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        self.find_entry = self._entry(self.find_bar, "Find in output", width=260, height=28)
        self.find_entry.pack(side="left", padx=8, pady=6)
        self.find_entry.bind("<KeyRelease>", self._on_find_key)
        self.find_entry.bind("<Return>", lambda e: self._find_step(1))
        self.find_entry.bind("<Shift-Return>", lambda e: self._find_step(-1))
        self.find_entry.bind("<Escape>", lambda e: self._close_find())
        self._btn(self.find_bar, "Prev", lambda: self._find_step(-1), "secondary", width=56, height=28).pack(side="left", padx=2)
        self._btn(self.find_bar, "Next", lambda: self._find_step(1), "secondary", width=56, height=28).pack(side="left", padx=2)
        self.find_count = self._label(self.find_bar, "", 11, color=P["muted"])
        self.find_count.pack(side="left", padx=10)
        self._btn(self.find_bar, "Close", self._close_find, "ghost", width=56, height=28).pack(side="right", padx=6)
        self.find_bar.grid_remove()
        frame, self.term = self._make_text(f)
        frame.grid(row=1, column=0, sticky="nsew")
        self.term.bind("<Double-Button-1>", self._on_term_double)
        self.term.bind("<Button-3>", self._on_term_right)
        self.term.bind("<Button-2>", self._on_term_right)
        if self.term_blocks:
            for args in self.term_blocks:
                self.term.insert("end", *args)
            self.term.see("end")
        else:
            self.term.insert("end", "Enter credentials on the left, a CMTS hostname above, then Connect or run any command.\n"
                                    "Right-click a MAC, IP, port or cable-mac for drill-downs. Double-click a MAC to stage it.\n",
                             ("dim",))

    def _build_table_view(self, f) -> None:
        P = self.P
        top = ctk.CTkFrame(f, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        self.tbl_menu = self._option(top, ["(no tables yet)"], self._on_table_pick, width=420)
        self.tbl_menu.pack(side="left")
        self.tbl_filter = self._entry(top, "Filter rows", width=220)
        self.tbl_filter.pack(side="left", padx=8)
        self.tbl_filter.bind("<KeyRelease>", lambda e: self._populate_tree())
        self.tbl_issues = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(top, text="Issues only", variable=self.tbl_issues, command=self._populate_tree,
                        font=self.F(12), text_color=P["text"], fg_color=P["accent"], hover_color=P["accent_hover"],
                        border_color=P["border"], checkbox_width=18, checkbox_height=18).pack(side="left", padx=4)
        self.tbl_count = self._label(top, "", 11, color=P["muted"])
        self.tbl_count.pack(side="left", padx=10)
        wrap = ctk.CTkFrame(f, fg_color=P["term_bg"], corner_radius=8, border_width=1, border_color=P["border"])
        wrap.grid(row=1, column=0, sticky="nsew")
        wrap.grid_rowconfigure(0, weight=1)
        wrap.grid_columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(wrap, style="CC.Treeview", show="headings", selectmode="browse")
        ys = ctk.CTkScrollbar(wrap, command=self.tree.yview, button_color=P["border"])
        xs = ctk.CTkScrollbar(wrap, command=self.tree.xview, orientation="horizontal", button_color=P["border"])
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(4, 0), pady=(4, 0))
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Double-1>", self._on_tree_double)
        self._refresh_table_menu()

    def _build_charts_view(self, f) -> None:
        self.chart_holder = ctk.CTkFrame(f, fg_color=self.P["card"], corner_radius=8, border_width=1,
                                         border_color=self.P["border"])
        self.chart_holder.grid(row=1, column=0, sticky="nsew")
        self.chart_msg = self._label(self.chart_holder, "Charts appear after a noise / PHY / modem command.", 12,
                                     color=self.P["muted"])
        self.chart_msg.pack(expand=True)
        self.fig = None
        self.canvas = None
        self.charts_dirty = True

    def _build_simple_text_view(self, tab: str, attr: str, placeholder: str) -> None:
        frame, t = self._make_text(self.views[tab], wrap="none" if tab != "Log" else "word")
        frame.grid(row=1, column=0, sticky="nsew")
        if placeholder:
            t.insert("end", placeholder + "\n", ("dim",))
        setattr(self, attr, t)

    def _build_fleet_view(self, f) -> None:
        P = self.P
        top = ctk.CTkFrame(f, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        self.fleet_lbl = self._label(top, "Fleet leaderboard  -  worst nodes first. Double-click a node to open its output.",
                                     12, color=P["muted"])
        self.fleet_lbl.pack(side="left")
        wrap = ctk.CTkFrame(f, fg_color=P["term_bg"], corner_radius=8, border_width=1, border_color=P["border"])
        wrap.grid(row=1, column=0, sticky="nsew")
        wrap.grid_rowconfigure(0, weight=1)
        wrap.grid_columnconfigure(0, weight=1)
        cols = ("rank", "host", "online", "total", "offline", "us_snr", "uncorr", "issues", "verdict", "secs")
        heads = ("#", "CMTS", "Online %", "Modems", "Offline", "Min US SNR", "Max uncorr", "Issue rows", "Verdict", "Time")
        widths = (40, 180, 90, 90, 80, 100, 100, 90, 260, 70)
        self.fleet_tree = ttk.Treeview(wrap, columns=cols, show="headings", style="CC.Treeview")
        for c, h, w in zip(cols, heads, widths):
            self.fleet_tree.heading(c, text=h)
            self.fleet_tree.column(c, width=w, anchor="w" if c in ("host", "verdict") else "e", stretch=c == "verdict")
        ys = ctk.CTkScrollbar(wrap, command=self.fleet_tree.yview, button_color=P["border"])
        self.fleet_tree.configure(yscrollcommand=ys.set)
        self.fleet_tree.grid(row=0, column=0, sticky="nsew", padx=(4, 0), pady=4)
        ys.grid(row=0, column=1, sticky="ns")
        self.fleet_tree.bind("<Double-1>", self._on_fleet_double)
        self._fill_fleet_tree()

    def _build_history_view(self, f) -> None:
        P = self.P
        top = ctk.CTkFrame(f, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        self.hist_entry = self._entry(top, "Search host, command or output text", width=340)
        self.hist_entry.pack(side="left")
        self.hist_entry.bind("<Return>", lambda e: self.refresh_history())
        self._btn(top, "Search", self.refresh_history, "secondary", width=80).pack(side="left", padx=6)
        self._btn(top, "Replay in terminal", self.replay_history, "secondary").pack(side="left", padx=(0, 6))
        self._btn(top, "Chronic offenders", self.run_chronic, "secondary").pack(side="left")
        body = ctk.CTkFrame(f, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew")
        body.grid_rowconfigure(0, weight=1)
        body.grid_columnconfigure(0, weight=2)
        body.grid_columnconfigure(1, weight=3)
        wrap = ctk.CTkFrame(body, fg_color=P["term_bg"], corner_radius=8, border_width=1, border_color=P["border"])
        wrap.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        wrap.grid_rowconfigure(0, weight=1)
        wrap.grid_columnconfigure(0, weight=1)
        self.hist_tree = ttk.Treeview(wrap, columns=("ts", "host", "cmd"), show="headings", style="CC.Treeview")
        for c, h, w in (("ts", "Time", 140), ("host", "CMTS", 120), ("cmd", "Command", 260)):
            self.hist_tree.heading(c, text=h)
            self.hist_tree.column(c, width=w, anchor="w", stretch=c == "cmd")
        ys = ctk.CTkScrollbar(wrap, command=self.hist_tree.yview, button_color=P["border"])
        self.hist_tree.configure(yscrollcommand=ys.set)
        self.hist_tree.grid(row=0, column=0, sticky="nsew", padx=(4, 0), pady=4)
        ys.grid(row=0, column=1, sticky="ns")
        self.hist_tree.bind("<<TreeviewSelect>>", self._on_hist_select)
        frame, self.hist_text = self._make_text(body)
        frame.grid(row=0, column=1, sticky="nsew")

    def _setup_tree_style(self) -> None:
        P = self.P
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("CC.Treeview", background=P["term_bg"], fieldbackground=P["term_bg"], foreground=P["term_fg"],
                     rowheight=24, borderwidth=0, font=(MONO_FONT, 10))
        st.configure("CC.Treeview.Heading", background=P["card"], foreground=P["muted"], relief="flat",
                     borderwidth=0, font=(UI_FONT, 10, "bold"), padding=(6, 5))
        st.map("CC.Treeview", background=[("selected", P["sel"])], foreground=[("selected", P["text"])])
        st.map("CC.Treeview.Heading", background=[("active", P["hover"])])
        for tree in (self.tree, self.fleet_tree, self.hist_tree):
            tree.tag_configure("crit", foreground=P["crit"])
            tree.tag_configure("warn", foreground=P["warn"])
            tree.tag_configure("ok", foreground=P["term_fg"])
            tree.tag_configure("odd", background=P["stripe"])

    # ------------------------------------------------------------------ tabs
    def _show_tab(self, name: str) -> None:
        self.current_tab = name
        self.tabbar.set(name)
        self.views[name].tkraise()
        if name == "Charts" and self.charts_dirty:
            self.after(10, self._draw_charts)
        if name == "History":
            self.refresh_history()

    # ------------------------------------------------------------------ form helpers
    def _creds(self, host: Optional[str] = None) -> Optional[Creds]:
        c = Creds(self.bastion_entry.get().strip(), self.user_entry.get().strip(), self.pw_entry.get(),
                  (host or self.host_entry.get()).strip())
        missing = [n for n, v in (("CMTS host", c.host), ("username", c.user), ("password", c.password)) if not v]
        if missing:
            messagebox.showwarning("Missing details", "Please fill in: " + ", ".join(missing), parent=self)
            return None
        return c

    def _params(self) -> dict:
        return dict(cmac=self.cmac_entry.get(), interface=self.iface_entry.get(), mac_address=self.mac_entry.get())

    def _ask_otp(self) -> Optional[str]:
        d = ctk.CTkInputDialog(text="Enter the 6-digit VIP token for the bastion:", title="Bastion authentication",
                               fg_color=self.P["card"], button_fg_color=self.P["accent"],
                               button_hover_color=self.P["accent_hover"], entry_fg_color=self.P["input"],
                               entry_border_color=self.P["border"], entry_text_color=self.P["text"])
        v = d.get_input()
        if not v or not v.strip():
            self._set_activity("Cancelled - no VIP token entered.", "warn")
            return None
        return v.strip()

    def _set_activity(self, text: str, sev: Optional[str] = None) -> None:
        self.activity.configure(text=text, text_color=self.P.get(sev, self.P["muted"]) if sev else self.P["muted"])

    def _set_busy_ui(self) -> None:
        busy = self.busy > 0
        self.stop_btn.configure(state="normal" if busy or self.watching else "disabled")
        for b in (self.run_quick_btn, self.run_pb_btn, self.fleet_btn):
            b.configure(state="disabled" if busy else "normal")
        if not busy and not self.watching:
            self._set_activity("Ready")

    # ------------------------------------------------------------------ connection
    def connect(self) -> None:
        creds = self._creds()
        if not creds:
            return
        otp = ""
        if self.sessions.needs_otp(creds.bastion, creds.user, creds.host):
            otp = self._ask_otp()
            if otp is None:
                return
        self._set_activity(f"Connecting to {creds.host} ...")
        self._enqueue(lambda: self._job_connect(creds, otp))

    def _job_connect(self, creds: Creds, otp: str) -> None:
        cli = self._get_cli(creds, otp)
        self.ui(self._on_connected, creds.host, cli.prompt_text)
        self.log(f"Connected to {creds.host} ({cli.prompt_text})")

    def _get_cli(self, creds: Creds, otp: str) -> CliChannel:
        cli = self.sessions.get(creds.host)
        if cli:
            return cli
        if creds.bastion and not self.sessions.bastion_ready(creds.bastion, creds.user):
            if not otp:
                raise SessionError("Bastion session expired. Run the command again to enter a new VIP token.")
            self.sessions.connect_bastion(creds.bastion, creds.user, creds.password, otp, self.st)
        self.ui(self._set_activity, f"Logging in to {creds.host} ...")
        cli = self.sessions.open_cmts(creds.bastion, creds.user, creds.password, otp, creds.host, self.st)
        self.ui(self._on_connected, creds.host, cli.prompt_text)
        return cli

    _prompt_cache = ""

    def _on_connected(self, host: str, prompt: str) -> None:
        self.connected_host = host
        self._prompt_cache = prompt
        self.conn_dot.configure(text_color=self.P["ok"])
        self.conn_lbl.configure(text=f"Connected  ·  {host}", text_color=self.P["text"])
        self.prompt_lbl.configure(text=prompt or f"{host}#", text_color=self.P["accent"])

    def _on_disconnected(self, why: str = "Not connected") -> None:
        self.connected_host = ""
        self.conn_dot.configure(text_color=self.P["dim"])
        self.conn_lbl.configure(text=why, text_color=self.P["muted"])
        self.prompt_lbl.configure(text="not connected >", text_color=self.P["dim"])

    def disconnect(self) -> None:
        self.cancel_event.set()
        self._stop_watch()
        threading.Thread(target=self.sessions.close_all, daemon=True).start()
        self._on_disconnected()
        self.log("Disconnected.")

    def _monitor_connection(self) -> None:
        try:
            if self.connected_host and self.busy == 0:
                if self.sessions.get(self.connected_host) is None:
                    bastion_ok = self.sessions.transport is not None and self.sessions.transport.is_active()
                    self._on_disconnected("Bastion open · CMTS closed" if bastion_ok else "Session dropped")
                    if bastion_ok:
                        self.conn_dot.configure(text_color=self.P["warn"])
        finally:
            self.after(4000, self._monitor_connection)

    # ------------------------------------------------------------------ command dispatch
    def run_commands(self, cmds: List[str], label: Optional[str] = None, watch: bool = False,
                     host: Optional[str] = None) -> bool:
        cmds = [c for c in cmds if c.strip()]
        if not cmds:
            return False
        creds = self._creds(host)
        if not creds:
            return False
        if self.st["confirm_disruptive"] and not watch:
            risky = [c for c in cmds if is_disruptive(c)]
            if risky and not messagebox.askyesno(
                    "Confirm command",
                    "These commands can change the CMTS or affect subscribers:\n\n" + "\n".join(risky) +
                    f"\n\nSend to {creds.host}?", parent=self, icon="warning"):
                return False
        otp = ""
        if self.sessions.needs_otp(creds.bastion, creds.user, creds.host):
            otp = self._ask_otp()
            if otp is None:
                return False
        self._enqueue(lambda: self._job_run(creds, otp, cmds, label, watch))
        return True

    def _job_run(self, creds: Creds, otp: str, cmds: List[str], label: Optional[str], watch: bool) -> None:
        cli = self._get_cli(creds, otp)
        batch: List[Result] = []
        timeout = float(self.st["cmd_timeout"])
        for i, cmd in enumerate(cmds):
            if self.cancel_event.is_set():
                break
            self.ui(self._set_activity, f"Running {cmd}" + (f"  ({i + 1}/{len(cmds)})" if len(cmds) > 1 else ""))
            last = [0.0]

            def progress(buf, cmd=cmd):
                now = time.perf_counter()
                if now - last[0] > 0.3:
                    last[0] = now
                    self.ui(self._set_activity, f"Receiving {cmd[:50]}  ·  {buf.count(chr(10)):,} lines")

            t0 = time.perf_counter()
            out, why = cli.run(cmd, timeout, self.cancel_event, progress)
            dt = time.perf_counter() - t0
            parsed = parse_output(out, self.th)
            res = Result(creds.host, cmd, out, parsed, dt, why, datetime.now())
            batch.append(res)
            try:
                self.flight.record(creds.host, cmd, out)
            except Exception as e:
                self.log(f"History write failed: {e}")
            self.ui(self._show_result, res, render_segments(parsed), watch and i == 0)
            if why == "timeout":
                self.log(f"'{cmd}' timed out after {timeout:.0f}s - output may be incomplete.")
            if why in ("closed", "exited"):
                self.sessions.drop(creds.host)
                self.ui(self._on_disconnected, "CMTS session closed")
                break
        if batch:
            combined = merge_parsed([r.parsed for r in batch])
            tri = run_triage(combined, label or batch[-1].cmd, self.th)
            self.ui(self._on_batch_done, batch, combined, tri, label, watch)

    def _show_result(self, res: Result, segs: list, replace_all: bool) -> None:
        nlines = res.output.count("\n") + 1 if res.output else 0
        note = {"timeout": "  ·  TIMED OUT", "cancelled": "  ·  stopped", "confirm": "  ·  waiting for confirmation",
                "closed": "  ·  session closed", "exited": "  ·  left CMTS"}.get(res.status, "")
        hdr = ["\n" if self.term.index("end-1c") != "1.0" and not replace_all else "", (),
               f"{res.ts:%H:%M:%S}  ", ("dim",), f"{res.host}  ", ("hdr",), res.cmd, ("cmd",),
               f"   {res.duration:.2f}s · {nlines:,} lines{note}\n", ("warn",) if note else ("dim",)]
        if self.watching:
            hdr[2] = f"LIVE #{self.watch_cycle}  {res.ts:%H:%M:%S}  "
        if not res.output.strip():
            segs = ["(no output)\n", ("dim",)]
        if replace_all:
            self.term.delete("1.0", "end")
            self.term_blocks.clear()
        block = hdr + segs
        self.term_blocks.append(block)
        self.term.insert("end", *block)
        last = int(self.term.index("end-1c").split(".")[0])
        over = last - int(self.st["term_max_lines"])
        if over > 0:
            self.term.delete("1.0", f"{over + 1}.0")
        self.term.see("end")
        if self.find_bar.winfo_ismapped() and self.find_entry.get():
            self._on_find_key()
        if res.status == "confirm":
            self._set_activity("CMTS is asking for confirmation - type the answer in the CLI and press Send.", "warn")
        self.log(f"{res.host}: {res.cmd} -> {res.status}, {nlines} lines in {res.duration:.2f}s")

    def _on_batch_done(self, batch: List[Result], combined: Parsed, tri: Triage, label: Optional[str],
                       watch: bool) -> None:
        self.results = batch
        self.last_parsed = combined
        self.last_triage = tri
        if not watch:
            self.watch_cmds = [r.cmd for r in batch]
        self._update_kpis(combined, tri, label or batch[-1].cmd)
        self._render_triage(tri, batch, label)
        self._refresh_table_menu()
        self._append_trend(batch[0].host, combined)
        self.charts_dirty = True
        has_rf = bool(combined.metrics.get("us_snr") or combined.metrics.get("ds_snr") or combined.metrics.get("us_tx"))
        if self.current_tab == "Charts":
            self._draw_charts()
        elif has_rf and self.st["auto_charts"] and not watch:
            self._show_tab("Charts")
        issues = combined.issue_count
        sev = tri.sev
        msg = f"Done · {sum(r.duration for r in batch):.2f}s"
        if issues:
            msg += f" · {issues} row(s) flagged"
        if tri.findings:
            msg += f" · {tri.findings[0].title}"
        self._set_activity(msg, sev if sev in ("warn", "crit") else None)
        if self.watching and watch:
            self.watch_lbl.configure(text=f"Cycle #{self.watch_cycle} at {datetime.now():%H:%M:%S} · next in {self.watch_interval}s",
                                     text_color=self.P["ok"])
            self.after(self.watch_interval * 1000, self._watch_tick)

    def send_cli(self) -> None:
        self.popup.withdraw()
        raw = self.cli.get()
        cmds = [c.strip() for c in re.split(r"[\r\n]+", raw) if c.strip()]
        if not cmds:
            return
        for c in cmds:
            if not self.cmd_history or self.cmd_history[-1] != c:
                self.cmd_history.append(c)
        self.cmd_history = self.cmd_history[-200:]
        self.hist_idx = len(self.cmd_history)
        if self.run_commands(cmds):
            self.cli.delete(0, "end")
            if self.current_tab not in ("Terminal", "Table", "Charts", "Triage"):
                self._show_tab("Terminal")

    def stop(self) -> None:
        self.cancel_event.set()
        self._stop_watch()
        self._set_activity("Stopping ...", "warn")

    def _quick_cmd(self) -> Optional[str]:
        spec = COMMAND_CATALOG[self.quick_menu.get()]
        p = self._params()
        missing = [k for k in spec.params if not p.get(k, "").strip()]
        if missing:
            names = {"cmac": "Cable-Mac", "interface": "Interface", "mac_address": "Modem MAC"}
            messagebox.showwarning("Parameter needed", f"'{spec.title}' needs: " + ", ".join(names[m] for m in missing),
                                   parent=self)
            return None
        return spec.format(**p)

    def run_quick(self) -> None:
        cmd = self._quick_cmd()
        if cmd:
            self.run_commands([cmd])

    def quick_to_cli(self) -> None:
        spec = COMMAND_CATALOG[self.quick_menu.get()]
        self.cli.delete(0, "end")
        self.cli.insert(0, spec.format(**self._params()))
        self.cli.focus_set()

    def _on_quick_select(self, choice: str) -> None:
        self.quick_desc.configure(text=COMMAND_CATALOG[choice].desc)

    def _on_pb_select(self, choice: str) -> None:
        spec = PLAYBOOK_CATALOG[choice]
        self.pb_desc.configure(text=f"{spec.desc}  ({len(spec.templates)} commands)")

    def run_playbook(self) -> None:
        name = self.pb_menu.get()
        spec = PLAYBOOK_CATALOG[name]
        p = self._params()
        missing = [k for k in spec.params if not p.get(k, "").strip()]
        if missing:
            messagebox.showwarning("Parameter needed", f"'{name}' needs: " + ", ".join(missing), parent=self)
            return
        self.run_commands(spec.format(**p), label=f"Playbook: {name}")

    # ------------------------------------------------------------------ watch
    def _on_watch_interval(self, choice: str) -> None:
        self.watch_interval = {"5 s": 5, "15 s": 15, "30 s": 30, "60 s": 60, "5 min": 300}[choice]

    def toggle_watch(self) -> None:
        if self.watching:
            self._stop_watch()
            return
        cmds = self.watch_cmds
        if not cmds:
            q = self._quick_cmd()
            cmds = [q] if q else []
        if not cmds:
            return
        self.watch_cmds = cmds
        self.watching = True
        self.watch_cycle = 0
        self.watch_btn.configure(text="Stop watch", fg_color=self.P["crit"], text_color="#FFFFFF")
        self.watch_lbl.configure(text=f"Watching: {cmds[0][:40]}", text_color=self.P["ok"])
        self._watch_tick()

    def _watch_tick(self) -> None:
        if not self.watching:
            return
        if self.busy:
            self.after(1000, self._watch_tick)
            return
        self.watch_cycle += 1
        if not self.run_commands(self.watch_cmds, watch=True):
            self._stop_watch()

    def _stop_watch(self) -> None:
        if not self.watching:
            return
        self.watching = False
        self.watch_btn.configure(text="Start watch", fg_color=self.P["hover"], text_color=self.P["text"])
        self.watch_lbl.configure(text="Idle", text_color=self.P["dim"])
        self._set_busy_ui()

    # ------------------------------------------------------------------ KPIs & triage view
    def _set_kpi(self, key: str, value: str, sub: str = "", sev: Optional[str] = None) -> None:
        self.kpi_state[key] = (value, sub, sev)
        v, s = self.kpi[key]
        v.configure(text=value, text_color=self.P.get(sev, self.P["text"]) if sev else self.P["text"])
        s.configure(text=sub or " ")

    def _update_kpis(self, p: Parsed, tri: Triage, source: str) -> None:
        stamp = f"{datetime.now():%H:%M:%S}"
        if p.summary and p.summary["total"]:
            pct = 100.0 * p.summary["online"] / p.summary["total"]
            sev = classify_value("online_pct", pct, self.th)
            self._set_kpi("online", f"{pct:.1f}%", f"{p.summary['online']:,} of {p.summary['total']:,} · {stamp}", sev)
            off = p.summary["offline"]
            self._set_kpi("offline", f"{off:,}", f"modems not online · {stamp}", sev if off else "ok")
        for key in ("us_snr", "ds_snr"):
            vals = p.metrics.get(key)
            if vals:
                mn = min(vals)
                avg = sum(vals) / len(vals)
                self._set_kpi(key, f"{mn:.1f} dB", f"avg {avg:.1f} · {len(vals)} readings", classify_value(key, mn, self.th))
        if p.metrics.get("uncorr_pct"):
            mx = max(p.metrics["uncorr_pct"])
            self._set_kpi("cw", f"{mx:.2f}%", "peak uncorrectable", classify_value("uncorr_pct", mx, self.th))
        elif p.metrics.get("uncorr"):
            tot = sum(p.metrics["uncorr"])
            self._set_kpi("cw", f"{tot:,.0f}", "uncorrectables (cumulative)", "warn" if tot else "ok")
        if tri.dispatch_sev is not None:
            self._set_kpi("dispatch", tri.dispatch, source[:34], tri.dispatch_sev)

    def _render_triage(self, tri: Triage, batch: List[Result], label: Optional[str]) -> None:
        a: list = [f"Automated triage  ·  {batch[0].host}  ·  {datetime.now():%Y-%m-%d %H:%M:%S}\n", ("hdr",),
                   f"{label or batch[-1].cmd}\n\n", ("dim",),
                   "Dispatch verdict   ", ("th",), f"{tri.dispatch}\n", (tri.dispatch_sev or "dim",),
                   f"  {tri.dispatch_detail}\n\n", ()]
        a += ["Findings\n", ("th",)]
        if tri.findings:
            for fd in sorted(tri.findings, key=lambda x: -SEV_RANK[x.sev]):
                a += [f"  {'CRITICAL' if fd.sev == 'crit' else 'WARNING '}  ", (fd.sev,), f"{fd.title}\n", ("cmd",),
                      f"            {fd.detail}\n", ()]
        else:
            a += ["  No thresholds exceeded.\n", ("ok",)]
        if tri.blast:
            a += ["\nBlast radius\n", ("th",)]
            for sev, text in tri.blast:
                a += [f"  {text}\n", (sev,) if sev else ()]
        flagged = [(t, i) for t in self.last_parsed.tables for i, s in enumerate(t.row_sev) if s == "crit"][:25]
        if flagged:
            a += ["\nCritical rows (first 25)\n", ("th",)]
            for t, i in flagged:
                a += ["  " + "  ".join(c for c in t.rows[i] if c)[:150] + "\n", ("crit",),
                      "      " + "; ".join(t.row_issues[i]) + "\n", ("dim",)]
        self.triage_args = a
        self._write_args(self.triage_text, a)

    def _write_args(self, t: tk.Text, args: list) -> None:
        t.delete("1.0", "end")
        if args:
            t.insert("end", *args)

    # ------------------------------------------------------------------ table view
    def _all_tables(self) -> List[Tuple[Result, Table]]:
        return [(r, t) for r in self.results for t in r.parsed.tables if t.rows]

    def _refresh_table_menu(self) -> None:
        tabs = self._all_tables()
        if not tabs:
            self.tbl_menu.configure(values=["(no tables yet)"])
            self.tbl_menu.set("(no tables yet)")
            self.tbl_table = None
            self._populate_tree()
            return
        labels = [f"{i + 1}. {r.cmd[:46]}  ·  {len(t.rows):,} rows" for i, (r, t) in enumerate(tabs)]
        self._tbl_choices = dict(zip(labels, [t for _, t in tabs]))
        self.tbl_menu.configure(values=labels)
        best = max(range(len(tabs)), key=lambda i: (tabs[i][1].has_modems, len(tabs[i][1].rows)))
        self.tbl_menu.set(labels[best])
        self.tbl_table = tabs[best][1]
        self.tbl_sort = (-1, False)
        self._populate_tree()

    def _on_table_pick(self, label: str) -> None:
        t = getattr(self, "_tbl_choices", {}).get(label)
        if t is not None:
            self.tbl_table = t
            self.tbl_sort = (-1, False)
            self._populate_tree()

    def _populate_tree(self) -> None:
        tree = self.tree
        tree.delete(*tree.get_children())
        t = self.tbl_table
        if t is None:
            tree["columns"] = ("msg",)
            tree.heading("msg", text="")
            tree.column("msg", width=600)
            tree.insert("", "end", values=("Run a command that returns a table (modem list, noise, summary ...)",))
            self.tbl_count.configure(text="")
            return
        names = ["Health"] + [c.name or f"col {i + 1}" for i, c in enumerate(t.columns)] + ["Issues"]
        ids = [f"c{i}" for i in range(len(names))]
        tree["columns"] = ids
        for i, (cid, name) in enumerate(zip(ids, names)):
            arrow = ""
            if self.tbl_sort[0] == i:
                arrow = "  ▼" if self.tbl_sort[1] else "  ▲"
            tree.heading(cid, text=name + arrow, command=lambda i=i: self._sort_tree(i))
            col = t.columns[i - 1] if 0 < i <= len(t.columns) else None
            width = 70 if i == 0 else 260 if i == len(names) - 1 else max(70, min(240, 9 * max(len(name), 6)))
            anchor = "e" if col is not None and (col.metric or col.role in ("total", "online", "offline")) else "w"
            tree.column(cid, width=width, anchor=anchor, stretch=False)
        rows = []
        for cells, sev, issues in zip(t.rows, t.row_sev, t.row_issues):
            label = {"crit": "CRIT", "warn": "WARN", "ok": "OK"}.get(sev, "")
            rows.append(([label] + cells + ["; ".join(issues)], sev))
        q = self.tbl_filter.get().strip().lower()
        if q:
            rows = [r for r in rows if any(q in str(v).lower() for v in r[0])]
        if self.tbl_issues.get():
            rows = [r for r in rows if r[1] in ("warn", "crit")]
        ci, rev = self.tbl_sort
        if ci >= 0:
            rows.sort(key=lambda r: _sort_key(r[0][ci], ci == 0), reverse=rev)
        shown = rows[:20000]
        for n, (vals, sev) in enumerate(shown):
            tree.insert("", "end", values=vals, tags=(sev or "ok",) + (("odd",) if n % 2 else ()))
        bad = sum(1 for s in t.row_sev if s in ("warn", "crit"))
        self.tbl_count.configure(text=f"{len(shown):,} of {len(t.rows):,} rows · {bad:,} flagged")

    def _sort_tree(self, ci: int) -> None:
        cur, rev = self.tbl_sort
        self.tbl_sort = (ci, (not rev) if cur == ci else False)
        self._populate_tree()

    def _on_tree_double(self, _e) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        mac = next((v for v in vals if MAC_RE.fullmatch(str(v))), None)
        if mac:
            self._stage_mac(mac)

    # ------------------------------------------------------------------ charts
    def _draw_charts(self) -> None:
        if self.current_tab != "Charts" and self.fig is None:
            return
        try:
            import matplotlib
            matplotlib.use("TkAgg")
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            from matplotlib.figure import Figure
        except Exception:
            self.chart_msg.configure(text="matplotlib is not installed - charts are unavailable.")
            return
        P = self.P
        p = self.last_parsed
        if self.fig is None:
            self.chart_msg.pack_forget()
            self.fig = Figure(figsize=(9, 6), dpi=100, facecolor=P["card"])
            self.canvas = FigureCanvasTkAgg(self.fig, master=self.chart_holder)
            self.canvas.get_tk_widget().configure(bg=P["card"], highlightthickness=0)
            self.canvas.get_tk_widget().pack(fill="both", expand=True, padx=6, pady=6)
        fig = self.fig
        fig.clear()
        gs = fig.add_gridspec(2, 2, height_ratios=[1.15, 1])
        ax1 = fig.add_subplot(gs[0, :])
        ax2 = fig.add_subplot(gs[1, 0])
        ax3 = fig.add_subplot(gs[1, 1])

        def style(ax, title):
            ax.set_facecolor(P["term_bg"])
            ax.set_title(title, color=P["text"], fontsize=10, fontweight="bold", loc="left")
            ax.tick_params(colors=P["muted"], labelsize=8)
            for s in ax.spines.values():
                s.set_color(P["border"])
            ax.grid(True, color=P["border"], linewidth=0.6, alpha=0.7)
            ax.set_axisbelow(True)

        sevc = {"ok": P["ok"], "warn": P["warn"], "crit": P["crit"], None: P["dim"]}
        # 1 - SNR per channel / modem
        metric = "us_snr" if p and any(m == "us_snr" for m, *_ in p.points) else "ds_snr"
        pts = [(lbl, v, s) for m, v, lbl, s in (p.points if p else []) if m == metric][:96]
        if pts:
            xs = range(len(pts))
            ax1.bar(xs, [v for _, v, _ in pts], color=[sevc[s] for *_, s in pts], width=0.75)
            th = self.th[metric]
            ax1.axhline(th["good"], color=P["ok"], linestyle="--", linewidth=1, alpha=0.8)
            ax1.axhline(th["warn"], color=P["crit"], linestyle="--", linewidth=1, alpha=0.8)
            if len(pts) <= 32:
                ax1.set_xticks(list(xs))
                ax1.set_xticklabels([str(lbl)[-14:] for lbl, *_ in pts], rotation=45, ha="right", fontsize=7)
            else:
                ax1.set_xticks([])
            ax1.set_ylabel("dB", color=P["muted"], fontsize=8)
            lo = min(v for _, v, _ in pts)
            ax1.set_ylim(max(0, min(lo, th["warn"]) - 5), max(v for _, v, _ in pts) + 3)
        style(ax1, f"{'Upstream' if metric == 'us_snr' else 'Downstream'} SNR / MER  ·  {len(pts)} readings"
              if pts else "SNR / MER  ·  no readings in last output")

        # 2 - power distributions
        drew = False
        for m, color, name in (("us_tx", P["accent"], "Modem US Tx"), ("ds_rx", P["iface"], "Modem DS Rx"),
                               ("cmts_rx", P["mac"], "CMTS US Rx")):
            vals = p.metrics.get(m) if p else None
            if vals:
                ax2.hist(vals, bins=min(30, max(6, len(vals) // 3)), color=color, alpha=0.75, label=f"{name} ({len(vals)})")
                g = self.th[m]["good"]
                ax2.axvspan(g[0], g[1], color=color, alpha=0.08)
                drew = True
        if drew:
            leg = ax2.legend(fontsize=7, facecolor=P["card"], edgecolor=P["border"])
            for txt in leg.get_texts():
                txt.set_color(P["text"])
            ax2.set_xlabel("dBmV (shaded = good window)", color=P["muted"], fontsize=8)
        style(ax2, "Power levels" if drew else "Power levels  ·  none in last output")

        # 3 - trend
        tr = self.trend.get(self.connected_host or (self.results[0].host if self.results else ""), deque())
        if len(tr) >= 2:
            times = [x[0] for x in tr]
            if any(x[1] is not None for x in tr):
                ax3.plot(times, [x[1] for x in tr], color=P["crit"], marker="o", markersize=3, linewidth=1.4, label="Min US SNR")
                ax3.plot(times, [x[2] for x in tr], color=P["accent"], marker="o", markersize=3, linewidth=1.4, label="Avg US SNR")
                ax3.set_ylabel("dB", color=P["muted"], fontsize=8)
            if any(x[3] is not None for x in tr):
                ax4 = ax3.twinx()
                ax4.plot(times, [x[3] for x in tr], color=P["ok"], linewidth=1.2, linestyle=":", label="Online %")
                ax4.tick_params(colors=P["muted"], labelsize=8)
                ax4.set_ylabel("online %", color=P["muted"], fontsize=8)
            ax3.tick_params(axis="x", rotation=30)
            leg = ax3.legend(fontsize=7, facecolor=P["card"], edgecolor=P["border"], loc="lower left")
            for txt in leg.get_texts():
                txt.set_color(P["text"])
        style(ax3, "Trend (watch / repeated runs)" if len(tr) >= 2 else "Trend  ·  run Live watch to build history")
        fig.tight_layout(pad=1.6)
        self.canvas.draw_idle()
        self.charts_dirty = False

    def _append_trend(self, host: str, p: Parsed) -> None:
        us = p.metrics.get("us_snr")
        pct = 100.0 * p.summary["online"] / p.summary["total"] if p.summary and p.summary["total"] else None
        if not us and pct is None:
            return
        self.trend.setdefault(host, deque(maxlen=60)).append(
            (datetime.now().strftime("%H:%M:%S"), min(us) if us else None, sum(us) / len(us) if us else None, pct))

    # ------------------------------------------------------------------ terminal interactions
    def clear_terminal(self) -> None:
        self.term.delete("1.0", "end")
        self.term_blocks.clear()
        self.find_hits = []
        self.find_count.configure(text="")

    def copy_output(self) -> None:
        t = {"Triage": self.triage_text, "Diff": self.diff_text, "Log": self.log_text,
             "History": self.hist_text}.get(self.current_tab, self.term)
        try:
            text = t.get("sel.first", "sel.last")
        except tk.TclError:
            text = t.get("1.0", "end").strip()
        if text:
            self.clipboard_clear()
            self.clipboard_append(text)
            self._set_activity(f"Copied {len(text.splitlines()):,} lines to clipboard.", "ok")

    def _open_find(self) -> None:
        self._show_tab("Terminal")
        self.find_bar.grid()
        self.find_entry.focus_set()
        self.find_entry.select_range(0, "end")

    def _close_find(self) -> None:
        self.term.tag_remove("hit", "1.0", "end")
        self.term.tag_remove("hit_cur", "1.0", "end")
        self.find_bar.grid_remove()

    def _on_find_key(self, e=None) -> None:
        if e is not None and e.keysym in ("Return", "Escape", "Shift_L", "Shift_R"):
            return
        q = self.find_entry.get()
        t = self.term
        t.tag_remove("hit", "1.0", "end")
        t.tag_remove("hit_cur", "1.0", "end")
        self.find_hits = []
        if not q:
            self.find_count.configure(text="")
            return
        pos = "1.0"
        cnt = tk.IntVar()
        while len(self.find_hits) < 5000:
            pos = t.search(q, pos, stopindex="end", nocase=True, count=cnt)
            if not pos:
                break
            end = f"{pos}+{cnt.get()}c"
            self.find_hits.append((pos, end))
            t.tag_add("hit", pos, end)
            pos = end
        self.find_idx = len(self.find_hits) - 1
        self._find_step(0)

    def _find_step(self, d: int) -> None:
        if not self.find_hits:
            self.find_count.configure(text="no matches" if self.find_entry.get() else "")
            return
        self.find_idx = (self.find_idx + d) % len(self.find_hits)
        s, e = self.find_hits[self.find_idx]
        self.term.tag_remove("hit_cur", "1.0", "end")
        self.term.tag_add("hit_cur", s, e)
        self.term.see(s)
        self.find_count.configure(text=f"{self.find_idx + 1} / {len(self.find_hits)}")

    def _stage_mac(self, mac: str) -> None:
        self.mac_entry.delete(0, "end")
        self.mac_entry.insert(0, mac)
        self.cli.delete(0, "end")
        self.cli.insert(0, f"show cable modem {mac} verbose")
        self.cli.focus_set()
        self._set_activity(f"Staged {mac} ({lookup_mac_vendor(mac)}) - press Enter to run.")

    def _on_term_double(self, e) -> None:
        idx = self.term.index(f"@{e.x},{e.y}")
        line = self.term.get(f"{idx} linestart", f"{idx} lineend")
        m = MAC_RE.search(line)
        if m:
            self._stage_mac(m.group())

    def _on_term_right(self, e) -> None:
        P = self.P
        idx = self.term.index(f"@{e.x},{e.y}")
        line = self.term.get(f"{idx} linestart", f"{idx} lineend")
        menu = tk.Menu(self, tearoff=0, bg=P["card"], fg=P["text"], activebackground=P["accent"],
                       activeforeground=P["on_accent"], bd=0, font=(UI_FONT, 10))
        mac = MAC_RE.search(line)
        ip = IP_RE.search(line)
        ifc = IFACE_RE.search(line)
        cm = re.search(r"cable-mac\s*(\d+)|\bcmac\s*(\d+)", line, re.I)
        if mac:
            m = mac.group()
            menu.add_command(label=f"{m}   {lookup_mac_vendor(m)}", state="disabled")
            menu.add_command(label="Modem 360 + truck-roll playbook", command=lambda: self._run_mac_playbook(m))
            for label, sfx in (("Verbose", "verbose"), ("PHY levels", "phy"), ("CPE devices", "cpe"),
                               ("Service flows", "service-flow")):
                menu.add_command(label=f"{label}", command=lambda s=sfx: self.run_commands([f"show cable modem {m} {s}"]))
            menu.add_command(label="Stage in CLI", command=lambda: self._stage_mac(m))
            menu.add_separator()
        if ip:
            menu.add_command(label=f"Ping {ip.group()}", command=lambda: self.run_commands([f"ping {ip.group()}"]))
            menu.add_separator()
        if ifc:
            i = ifc.group()
            menu.add_command(label=f"Port {i}: set as interface parameter",
                             command=lambda: (self.iface_entry.delete(0, "end"), self.iface_entry.insert(0, i)))
            menu.add_separator()
        if cm:
            c = cm.group(1) or cm.group(2)
            menu.add_command(label=f"Cable-mac {c}: blast-radius sweep", command=lambda: self._run_cmac_playbook(c))
            menu.add_command(label=f"Cable-mac {c}: noise", command=lambda: self.run_commands([f"show cable noise cable-mac {c}"]))
            menu.add_separator()
        menu.add_command(label="Copy line", command=lambda: (self.clipboard_clear(), self.clipboard_append(line)))
        menu.add_command(label="Copy ticket notes", command=self.copy_ticket)
        menu.add_command(label="Export output to Excel...", command=self.export_output_excel)
        try:
            menu.tk_popup(e.x_root, e.y_root)
        finally:
            menu.grab_release()

    def _run_mac_playbook(self, mac: str) -> None:
        self.mac_entry.delete(0, "end")
        self.mac_entry.insert(0, mac)
        self.pb_menu.set("Modem 360 + Truck-Roll Classifier")
        self._on_pb_select(self.pb_menu.get())
        self.run_playbook()

    def _run_cmac_playbook(self, c: str) -> None:
        self.cmac_entry.delete(0, "end")
        self.cmac_entry.insert(0, c)
        self.pb_menu.set("Cable-Mac Blast-Radius Sweep")
        self._on_pb_select(self.pb_menu.get())
        self.run_playbook()

    # ------------------------------------------------------------------ CLI autocomplete / history
    def _on_cli_key(self, e) -> None:
        if e.keysym in ("Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R"):
            return
        q = self.cli.get().strip().lower()
        if len(q) < 2:
            self.popup.withdraw()
            return
        words = q.split()
        pool = list(dict.fromkeys(list(reversed(self.cmd_history)) + AUTOCOMPLETE_BASE))
        starts = [c for c in pool if c.lower().startswith(q) and c.lower() != q]
        contains = [c for c in pool if c not in starts and all(w in c.lower() for w in words) and c.lower() != q]
        items = (starts + contains)[:8]
        if items:
            self.popup.show(items, self.cli.winfo_rootx(), self.cli.winfo_rooty(), self.cli.winfo_width())
        else:
            self.popup.withdraw()

    def _cli_nav(self, d: int):
        if self.popup.visible:
            self.popup.move(d)
            return "break"
        if not self.cmd_history:
            return "break"
        self.hist_idx = max(0, min(len(self.cmd_history), self.hist_idx + d))
        self.cli.delete(0, "end")
        if self.hist_idx < len(self.cmd_history):
            self.cli.insert(0, self.cmd_history[self.hist_idx])
        return "break"

    def _on_cli_tab(self, _e):
        if self.popup.visible and self.popup.current():
            self._apply_autocomplete(self.popup.current())
        return "break"

    def _apply_autocomplete(self, cmd: str) -> None:
        self.cli.delete(0, "end")
        self.cli.insert(0, cmd)
        self.cli.icursor("end")
        self.popup.withdraw()
        self.cli.focus_set()

    # ------------------------------------------------------------------ maintenance snapshot / diff
    def take_snapshot(self) -> None:
        creds = self._creds()
        if not creds:
            return
        otp = ""
        if self.sessions.needs_otp(creds.bastion, creds.user, creds.host):
            otp = self._ask_otp()
            if otp is None:
                return
        cmds = list(self.st["snapshot_cmds"])
        self._enqueue(lambda: self._job_snapshot(creds, otp, cmds, post=False))

    def run_diff(self) -> None:
        if not self.snapshot:
            messagebox.showinfo("No baseline", "Take a pre-maintenance snapshot first.", parent=self)
            return
        creds = self._creds()
        if not creds:
            return
        if creds.host.lower() != self.snapshot["host"].lower() and not messagebox.askyesno(
                "Different CMTS", f"Baseline was taken on {self.snapshot['host']}. Compare against {creds.host} anyway?",
                parent=self):
            return
        otp = ""
        if self.sessions.needs_otp(creds.bastion, creds.user, creds.host):
            otp = self._ask_otp()
            if otp is None:
                return
        self._enqueue(lambda: self._job_snapshot(creds, otp, list(self.snapshot["outputs"]), post=True))

    def _job_snapshot(self, creds: Creds, otp: str, cmds: List[str], post: bool) -> None:
        cli = self._get_cli(creds, otp)
        outs: Dict[str, str] = {}
        for cmd in cmds:
            self.ui(self._set_activity, f"{'Post' if post else 'Pre'}-maintenance: {cmd}")
            out, _ = cli.run(cmd, float(self.st["cmd_timeout"]), self.cancel_event)
            outs[cmd] = out
            parsed = parse_output(out, self.th)
            res = Result(creds.host, cmd, out, parsed, 0.0, "prompt", datetime.now())
            self.ui(self._show_result, res, render_segments(parsed), False)
        if not post:
            snap = {"host": creds.host, "ts": datetime.now().strftime("%H:%M:%S"), "outputs": outs}
            self.ui(self._on_snapshot, snap)
        else:
            self.ui(self._on_diff, build_diff_args(self.snapshot, creds.host, outs, self.th))

    def _on_snapshot(self, snap: dict) -> None:
        self.snapshot = snap
        self.snap_lbl.configure(text=f"Baseline: {snap['host']} @ {snap['ts']}", text_color=self.P["ok"])
        self._set_activity(f"Baseline captured for {snap['host']} ({len(snap['outputs'])} commands).", "ok")

    def _on_diff(self, args: list) -> None:
        self.diff_args = args
        self._write_args(self.diff_text, args)
        self._show_tab("Diff")

    # ------------------------------------------------------------------ fleet
    def run_fleet(self) -> None:
        hosts = [h for h in re.split(r"[,\s]+", self.fleet_box.get("1.0", "end")) if h.strip()]
        hosts = list(dict.fromkeys(hosts))
        cmd = self.fleet_cmd.get().strip()
        if not hosts or not cmd:
            messagebox.showwarning("Fleet audit", "Enter at least one host and a command.", parent=self)
            return
        creds = self._creds(hosts[0])
        if not creds:
            return
        if self.st["confirm_disruptive"] and is_disruptive(cmd) and not messagebox.askyesno(
                "Confirm fleet command", f"Send this to {len(hosts)} CMTS nodes?\n\n{cmd}", parent=self, icon="warning"):
            return
        otp = ""
        if creds.bastion and not self.sessions.bastion_ready(creds.bastion, creds.user):
            otp = self._ask_otp()
            if otp is None:
                return
        self.fleet_lbl.configure(text=f"Auditing {len(hosts)} nodes ...", text_color=self.P["accent"])
        self._show_tab("Fleet")
        self._enqueue(lambda: self._job_fleet(creds, otp, hosts, cmd))

    def _job_fleet(self, creds: Creds, otp: str, hosts: List[str], cmd: str) -> None:
        if creds.bastion and not self.sessions.bastion_ready(creds.bastion, creds.user):
            self.sessions.connect_bastion(creds.bastion, creds.user, creds.password, otp, self.st)
        workers = max(1, min(int(self.st["fleet_workers"]), len(hosts)))
        self.log(f"Fleet audit: {len(hosts)} nodes, {workers} parallel channels, '{cmd}'")
        start = time.perf_counter()
        results: Dict[str, Result] = {}
        rows: List[dict] = []
        done = [0]

        def one(host: str) -> Result:
            t0 = time.perf_counter()
            cli = self.sessions.open_cmts(creds.bastion, creds.user, creds.password, otp, host, self.st, cache=False)
            try:
                out, why = cli.run(cmd, float(self.st["cmd_timeout"]), self.cancel_event)
            finally:
                cli.close()
            return Result(host, cmd, out, parse_output(out, self.th), time.perf_counter() - t0, why, datetime.now())

        with ThreadPoolExecutor(max_workers=workers) as pool:
            futs = {pool.submit(one, h): h for h in hosts}
            for fut in as_completed(futs):
                h = futs[fut]
                done[0] += 1
                try:
                    r = fut.result()
                    results[h] = r
                    rows.append(fleet_row(r, self.th))
                    try:
                        self.flight.record(h, cmd, r.output)
                    except Exception:
                        pass
                    self.log(f"[{done[0]}/{len(hosts)}] {h} done in {r.duration:.1f}s")
                except Exception as e:
                    rows.append({"host": h, "online": None, "total": None, "offline": None, "us_snr": None,
                                 "uncorr": None, "issues": None, "verdict": f"UNREACHABLE: {str(e)[:80]}",
                                 "sev": "crit", "secs": None})
                    self.log(f"[{done[0]}/{len(hosts)}] {h} failed: {e}")
                self.ui(self._set_activity, f"Fleet audit {done[0]}/{len(hosts)} ...")
        rows.sort(key=lambda r: (-SEV_RANK.get(r["sev"], 0), r["online"] if r["online"] is not None else -1,
                                 r["us_snr"] if r["us_snr"] is not None else 99))
        self.ui(self._on_fleet_done, results, rows, time.perf_counter() - start, cmd)

    def _on_fleet_done(self, results, rows, elapsed, cmd) -> None:
        self.fleet_results = results
        self.fleet_rows = rows
        self._fill_fleet_tree()
        bad = sum(1 for r in rows if r["sev"] == "crit")
        self.fleet_lbl.configure(
            text=f"{len(rows)} nodes · '{cmd}' · {elapsed:.1f}s · {bad} critical  -  double-click a node to open its output",
            text_color=self.P["crit"] if bad else self.P["muted"])
        self._set_activity(f"Fleet audit finished in {elapsed:.1f}s.", "crit" if bad else "ok")

    def _fill_fleet_tree(self) -> None:
        tr = self.fleet_tree
        tr.delete(*tr.get_children())

        def f(v, fmt):
            return "—" if v is None else fmt.format(v)

        for i, r in enumerate(self.fleet_rows, 1):
            tr.insert("", "end", iid=r["host"], tags=(r["sev"] or "ok",) + (("odd",) if i % 2 == 0 else ()), values=(
                i, r["host"], f(r["online"], "{:.1f}"), f(r["total"], "{:,}"), f(r["offline"], "{:,}"),
                f(r["us_snr"], "{:.1f} dB"), f(r["uncorr"], "{}"), f(r["issues"], "{:,}"), r["verdict"],
                f(r["secs"], "{:.1f}s")))

    def _on_fleet_double(self, _e) -> None:
        sel = self.fleet_tree.selection()
        if not sel or sel[0] not in self.fleet_results:
            return
        r = self.fleet_results[sel[0]]
        self.host_entry.delete(0, "end")
        self.host_entry.insert(0, r.host)
        self._show_result(r, render_segments(r.parsed), False)
        tri = run_triage(r.parsed, r.cmd, self.th)
        self._on_batch_done([r], r.parsed, tri, None, False)
        self._show_tab("Terminal")

    # ------------------------------------------------------------------ history
    def refresh_history(self) -> None:
        q = self.hist_entry.get() if hasattr(self, "hist_entry") else ""

        def job():
            try:
                recs = self.flight.search(q)
            except Exception as e:
                recs = []
                self.log(f"History search failed: {e}")
            self.ui(self._fill_history, recs)

        threading.Thread(target=job, daemon=True).start()

    def _fill_history(self, recs) -> None:
        tr = self.hist_tree
        tr.delete(*tr.get_children())
        for i, (rid, ts, host, cmd) in enumerate(recs):
            tr.insert("", "end", iid=str(rid), values=(ts, host, cmd), tags=(("odd",) if i % 2 else ()))

    def _on_hist_select(self, _e=None) -> None:
        sel = self.hist_tree.selection()
        if not sel:
            return
        rec = self.flight.get(int(sel[0]))
        if not rec:
            return
        ts, host, cmd, out = rec
        p = parse_output(out, self.th)
        self._write_args(self.hist_text, [f"{ts}  ", ("dim",), f"{host}  ", ("hdr",), f"{cmd}\n\n", ("cmd",)]
                         + render_segments(p))

    def replay_history(self) -> None:
        sel = self.hist_tree.selection()
        if not sel:
            messagebox.showinfo("History", "Select a stored output first.", parent=self)
            return
        ts, host, cmd, out = self.flight.get(int(sel[0]))
        p = parse_output(out, self.th)
        r = Result(host, cmd, out, p, 0.0, "replay", datetime.strptime(ts, "%Y-%m-%d %H:%M:%S"))
        self._show_result(r, render_segments(p), False)
        self._on_batch_done([r], p, run_triage(p, cmd, self.th), f"Replay: {cmd}", False)
        self._show_tab("Terminal")

    def run_chronic(self) -> None:
        self._set_activity("Analysing stored history ...")

        def job():
            try:
                args = chronic_offender_report(self.flight.recent(500), self.th)
            except Exception as e:
                args = [f"Analysis failed: {e}\n", ("crit",)]
            self.ui(self._write_args, self.hist_text, args)
            self.ui(self._set_activity, "Chronic offender analysis complete.", "ok")

        threading.Thread(target=job, daemon=True).start()

    # ------------------------------------------------------------------ exports
    def copy_ticket(self) -> None:
        if not self.results:
            self._set_activity("Nothing to copy yet - run a command first.", "warn")
            return
        tri = self.last_triage
        k = self.kpi_state
        kpis = " | ".join(f"{name}: **{k[key][0]}**" for key, name in
                          (("online", "Online"), ("offline", "Offline"), ("us_snr", "Min US SNR"),
                           ("ds_snr", "Min DS SNR"), ("cw", "Codewords")) if key in k)
        r = self.results[-1]
        snippet = "\n".join(r.output.splitlines()[:40])
        lines = [f"### E6000 diagnostic summary - {r.host}",
                 f"- **Time:** {datetime.now():%Y-%m-%d %H:%M:%S}",
                 f"- **Commands:** " + ", ".join(f"`{x.cmd}`" for x in self.results),
                 f"- **KPIs:** {kpis or 'n/a'}",
                 f"- **Dispatch verdict:** **{tri.dispatch}** - {tri.dispatch_detail}", "", "#### Findings"]
        lines += [f"- [{f_.sev.upper()}] **{f_.title}** - {f_.detail}" for f_ in tri.findings] or ["- No thresholds exceeded."]
        lines += [f"- {t}" for _, t in tri.blast]
        lines += ["", "#### CLI evidence", "```text", snippet, "```"]
        self.clipboard_clear()
        self.clipboard_append("\n".join(lines))
        self._set_activity("Ticket notes copied to clipboard (Markdown).", "ok")

    def _export_initial_dir(self) -> str:
        """Last export folder, else Documents, else home. Never the install folder:
        when installed under Program Files the working directory isn't writable."""
        d = str(self.st.get("export_dir") or "")
        if d and Path(d).is_dir():
            return d
        docs = Path.home() / "Documents"
        return str(docs if docs.is_dir() else Path.home())

    def export_html(self) -> None:
        if not self.results and not self.fleet_rows:
            messagebox.showinfo("Export", "Nothing to export yet.", parent=self)
            return
        host = self.results[-1].host if self.results else "fleet"
        path = filedialog.asksaveasfilename(parent=self, title="Save HTML report", defaultextension=".html",
                                            initialdir=self._export_initial_dir(),
                                            initialfile=f"E6000_{host}_{datetime.now():%Y%m%d_%H%M%S}.html",
                                            filetypes=[("HTML", "*.html")])
        if not path:
            return
        self.st["export_dir"] = str(Path(path).parent)
        img = ""
        if self.fig is not None or self.last_parsed:
            try:
                prev = self.current_tab
                self.current_tab = "Charts"
                self._draw_charts()
                self.current_tab = prev
                if self.fig is not None:
                    buf = io.BytesIO()
                    self.fig.savefig(buf, format="png", dpi=120, facecolor=self.P["card"], bbox_inches="tight")
                    img = f"<img alt='charts' src='data:image/png;base64,{base64.b64encode(buf.getvalue()).decode()}'/>"
            except Exception:
                img = ""
        try:
            Path(path).write_text(build_html_report(self, img), encoding="utf-8")
            webbrowser.open(Path(path).resolve().as_uri())
            self.log(f"HTML report saved: {path}")
        except Exception as e:
            messagebox.showerror("Export failed", str(e), parent=self)

    def export_excel(self) -> None:
        if openpyxl is None:
            messagebox.showerror("Excel", "openpyxl is not installed.", parent=self)
            return
        if not self.results and not self.fleet_results:
            messagebox.showinfo("Export", "Nothing to export yet.", parent=self)
            return
        host = self.results[-1].host if self.results else "fleet"
        path = filedialog.asksaveasfilename(parent=self, title="Save workbook", defaultextension=".xlsx",
                                            initialdir=self._export_initial_dir(),
                                            initialfile=f"E6000_{host}_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        self.st["export_dir"] = str(Path(path).parent)
        try:
            build_workbook(self).save(path)
            self._set_activity(f"Workbook saved: {path}", "ok")
            self.log(f"Workbook saved: {path}")
        except Exception as e:
            messagebox.showerror("Export failed", str(e), parent=self)

    def export_output_excel(self) -> None:
        """Copy the on-screen output (Terminal, or the active Triage/Diff/Log/History view) to an .xlsx
        file at a location the user picks. One worksheet row per line, colours preserved. If text is
        selected, only the selection is exported."""
        if openpyxl is None:
            messagebox.showerror("Excel", "openpyxl is not installed (pip install openpyxl).", parent=self)
            return
        view = self.current_tab if self.current_tab in ("Triage", "Diff", "Log", "History") else "Terminal"
        t = {"Triage": self.triage_text, "Diff": self.diff_text, "Log": self.log_text,
             "History": self.hist_text}.get(view, self.term)
        try:
            start, end, scope = t.index("sel.first"), t.index("sel.last"), "selection"
        except tk.TclError:
            start, end, scope = "1.0", "end-1c", "all"
        lines = text_widget_runs(t, start, end)
        while lines and not "".join(s for s, _ in lines[-1]).strip():
            lines.pop()
        if not any("".join(s for s, _ in ln).strip() for ln in lines):
            messagebox.showinfo("Export", f"The {view} view is empty - nothing to export.", parent=self)
            return
        host = self.connected_host or (self.results[-1].host if self.results else "") or "E6000"
        path = filedialog.asksaveasfilename(
            parent=self, title=f"Save {view.lower()} output as Excel", defaultextension=".xlsx",
            initialdir=self._export_initial_dir(), initialfile=f"E6000_{host}_{view.lower()}_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")])
        if not path:
            return
        self.st["export_dir"] = str(Path(path).parent)
        meta = {"view": view, "host": host, "scope": scope}
        self._set_activity(f"Exporting {len(lines):,} lines to Excel...", "warn")

        def job():
            try:
                build_output_workbook(lines, meta).save(path)
            except PermissionError:
                self.ui(messagebox.showerror, "Export failed",
                        f"Could not write to:\n{path}\n\nIs the file open in Excel? Close it and try again.")
                self.ui(self._set_activity, "Excel export failed - file in use or no write access.", "crit")
                return
            except Exception as e:
                self.ui(messagebox.showerror, "Export failed", str(e))
                self.ui(self._set_activity, f"Excel export failed: {e}", "crit")
                return
            what = "selected lines" if scope == "selection" else "lines"
            self.ui(self._set_activity, f"Exported {len(lines):,} {what} to {path}", "ok")
            self.log(f"{view} output saved to Excel: {path}")

        threading.Thread(target=job, daemon=True).start()

    # ------------------------------------------------------------------ settings
    def open_settings(self) -> None:
        SettingsDialog(self)

    def apply_settings(self, st: dict, th: Dict[str, dict]) -> None:
        theme_changed = st["theme"] != self.st["theme"]
        th_changed = json.dumps(th, sort_keys=True) != json.dumps(self.th, sort_keys=True)
        self.st = st
        self.th = th
        self._save_config()
        if theme_changed:
            self._rebuild_ui()
        if th_changed and self.results:
            results = self.results

            def job():
                rs = [replace(r, parsed=parse_output(r.output, self.th)) for r in results]
                comb = merge_parsed([r.parsed for r in rs])
                tri = run_triage(comb, rs[-1].cmd, self.th)
                self.ui(self._reapply_results, rs, comb, tri)

            threading.Thread(target=job, daemon=True).start()

    def _reapply_results(self, rs, comb, tri) -> None:
        self.clear_terminal()
        for r in rs:
            self._show_result(r, render_segments(r.parsed), False)
        self._on_batch_done(rs, comb, tri, None, False)

    def _rebuild_ui(self) -> None:
        state = {
            "bastion": self.bastion_entry.get(), "user": self.user_entry.get(), "pw": self.pw_entry.get(),
            "host": self.host_entry.get(), "cmac": self.cmac_entry.get(), "iface": self.iface_entry.get(),
            "mac": self.mac_entry.get(), "fleet": self.fleet_box.get("1.0", "end").strip(),
            "fleet_cmd": self.fleet_cmd.get(), "cli": self.cli.get(), "quick": self.quick_menu.get(),
            "pb": self.pb_menu.get(),
        }
        self.P = THEMES.get(str(self.st["theme"]), THEMES["Midnight"])
        ctk.set_appearance_mode(self.P["mode"])
        for w in list(self.winfo_children()):
            w.destroy()
        self.fig = None
        self.canvas = None
        self._build_ui()
        for ent, key in ((self.bastion_entry, "bastion"), (self.user_entry, "user"), (self.pw_entry, "pw"),
                         (self.host_entry, "host"), (self.cmac_entry, "cmac"), (self.iface_entry, "iface"),
                         (self.mac_entry, "mac"), (self.fleet_cmd, "fleet_cmd"), (self.cli, "cli")):
            ent.delete(0, "end")
            if state[key]:
                ent.insert(0, state[key])
        self.fleet_box.delete("1.0", "end")
        self.fleet_box.insert("1.0", state["fleet"])
        self.quick_menu.set(state["quick"])
        self._on_quick_select(state["quick"])
        self.pb_menu.set(state["pb"])
        self._on_pb_select(state["pb"])
        self._refresh_table_menu()
        if self.watching:
            self.watch_btn.configure(text="Stop watch", fg_color=self.P["crit"], text_color="#FFFFFF")

    def _save_config(self) -> None:
        save_config({
            "bastion_host": self.bastion_entry.get().strip(),
            "p_num": self.user_entry.get().strip(),
            "last_host": self.host_entry.get().strip(),
            "fleet_hosts": self.fleet_box.get("1.0", "end").strip(),
            "fleet_cmd": self.fleet_cmd.get().strip(),
            "theme": self.st["theme"],
            "settings": self.st,
            "thresholds": {k: {"good": v["good"], "warn": v["warn"]} for k, v in self.th.items()},
            "history": self.cmd_history[-200:],
        })

    def _on_close(self) -> None:
        self.watching = False
        self.cancel_event.set()
        try:
            self._save_config()
        except Exception:
            pass
        threading.Thread(target=self.sessions.close_all, daemon=True).start()
        self.jobs.put(None)
        self.after(150, self.destroy)


def _sort_key(v, is_health: bool = False):
    if is_health:
        return (0, {"CRIT": 0, "WARN": 1, "OK": 2}.get(str(v), 3), "")
    s = str(v).replace(",", "")
    m = re.match(r"\s*(-?\d+(?:\.\d+)?)", s)
    return (0, float(m.group(1)), "") if m else (1, 0.0, s.lower())


def fleet_row(r: Result, th: Dict[str, dict]) -> dict:
    p = r.parsed
    tri = run_triage(p, r.cmd, th)
    s = p.summary
    online = 100.0 * s["online"] / s["total"] if s and s["total"] else None
    us = p.metrics.get("us_snr")
    unc = p.metrics.get("uncorr_pct")
    sev = tri.sev
    if r.status == "timeout":
        sev = worst(sev, "warn")
    verdict = tri.findings[0].title if tri.findings else ("Healthy" if sev == "ok" else "No RF data" if sev is None else "")
    if r.status not in ("prompt",):
        verdict += f" ({r.status})"
    return {"host": r.host, "online": online, "total": s["total"] if s else None, "offline": s["offline"] if s else None,
            "us_snr": min(us) if us else None, "uncorr": f"{max(unc):.2f}%" if unc else None,
            "issues": p.issue_count, "verdict": verdict, "sev": sev or "ok", "secs": r.duration}


def build_diff_args(snapshot: dict, host: str, post: Dict[str, str], th) -> list:
    a: list = [f"Maintenance comparison  ·  baseline {snapshot['host']} @ {snapshot['ts']}  ->  {host} @ "
               f"{datetime.now():%H:%M:%S}\n\n", ("hdr",)]
    # port-level online comparison
    def port_map(text):
        p = parse_output(text, th)
        m = {}
        for t in p.tables:
            ii = t.col(role="iface")
            if ii is None:
                ii = 0
            for row in t.rows:
                s = _table_summary(Table(t.columns, [row]))
                if s and row[ii]:
                    m[row[ii]] = s
        return m

    changes = []
    for cmd, before in snapshot["outputs"].items():
        after = post.get(cmd, "")
        pre_m, post_m = port_map(before), port_map(after)
        for port in sorted(set(pre_m) | set(post_m)):
            b, c = pre_m.get(port), post_m.get(port)
            if b and c and b["online"] != c["online"]:
                changes.append((port, b["online"], c["online"], c["total"]))
            elif b and not c:
                changes.append((port, b["online"], None, b["total"]))
            elif c and not b:
                changes.append((port, None, c["online"], c["total"]))
    a += ["Port online counts\n", ("th",)]
    if changes:
        for port, b, c, tot in changes:
            if b is not None and c is not None:
                d = c - b
                a += [f"  {port:<24}", ("iface",), f" {b:>6,} -> {c:>6,}  ", (), f"({d:+,})\n", ("add" if d > 0 else "del",)]
            else:
                a += [f"  {port:<24}", ("iface",), f" {'missing after' if c is None else 'new after'}\n",
                      ("del" if c is None else "add",)]
    else:
        a += ["  No per-port online changes detected.\n", ("ok",)]
    for cmd, before in snapshot["outputs"].items():
        after = post.get(cmd, "")
        a += [f"\n{cmd}\n", ("th",)]
        diff = list(difflib.unified_diff(before.splitlines(), after.splitlines(), "pre", "post", lineterm="", n=1))
        if not diff:
            a += ["  identical\n", ("ok",)]
            continue
        for ln in diff[2:]:
            tag = "add" if ln.startswith("+") else "del" if ln.startswith("-") else "dim" if ln.startswith("@@") else ""
            a += [ln + "\n", (tag,) if tag else ()]
    return a


def build_html_report(app: CommandCenter, img_html: str) -> str:
    e = html.escape
    P = THEMES["Midnight"]
    tri = app.last_triage
    sevcol = {"ok": P["ok"], "warn": P["warn"], "crit": P["crit"]}
    kpis = "".join(
        f"<div class='kpi'><div class='t'>{e(name)}</div><div class='v' style='color:{sevcol.get(app.kpi_state[k][2], P['text'])}'>"
        f"{e(app.kpi_state[k][0])}</div><div class='s'>{e(app.kpi_state[k][1])}</div></div>"
        for k, name in (("online", "Online"), ("offline", "Offline"), ("us_snr", "Min US SNR"), ("ds_snr", "Min DS SNR"),
                        ("cw", "Codewords"), ("dispatch", "Dispatch")) if k in app.kpi_state)
    findings = ""
    if tri:
        items = "".join(f"<li><b style='color:{sevcol[f.sev]}'>{e(f.title)}</b> - {e(f.detail)}</li>" for f in tri.findings)
        items += "".join(f"<li>{e(t)}</li>" for _, t in tri.blast)
        findings = (f"<div class='card'><h2>Dispatch: <span style='color:{sevcol.get(tri.dispatch_sev, P['text'])}'>"
                    f"{e(tri.dispatch)}</span></h2><p>{e(tri.dispatch_detail)}</p><ul>{items or '<li>No thresholds exceeded.</li>'}</ul></div>")
    fleet = ""
    if app.fleet_rows:
        parts = []
        for i, r in enumerate(app.fleet_rows, 1):
            online = "" if r["online"] is None else f"{r['online']:.1f}%"
            snr = "" if r["us_snr"] is None else f"{r['us_snr']:.1f}"
            total = "" if r["total"] is None else f"{r['total']:,}"
            offline = "" if r["offline"] is None else f"{r['offline']:,}"
            parts.append(
                f"<tr style='color:{sevcol.get(r['sev'], P['text'])}'><td>{i}</td><td>{e(r['host'])}</td>"
                f"<td>{online}</td><td>{total}</td><td>{offline}</td><td>{snr}</td>"
                f"<td>{e(str(r['uncorr'] or ''))}</td><td>{e(r['verdict'])}</td></tr>")
        rows = "".join(parts)
        fleet = ("<div class='card'><h2>Fleet leaderboard</h2><table><tr><th>#</th><th>CMTS</th><th>Online</th><th>Modems</th>"
                 f"<th>Offline</th><th>Min US SNR</th><th>Uncorr</th><th>Verdict</th></tr>{rows}</table></div>")
    tables = ""
    for r in app.results:
        for t in r.parsed.tables[:4]:
            head = "".join(f"<th>{e(c.name)}</th>" for c in t.columns)
            body = ""
            for cells, sevs in list(zip(t.rows, t.cell_sev))[:400]:
                body += "<tr>" + "".join(
                    f"<td style='color:{sevcol.get(s, P['term_fg'])}'>{e(c)}</td>" for c, s in zip(cells, sevs)) + "</tr>"
            tables += f"<div class='card'><h2>{e(r.cmd)}</h2><table><tr>{head}</tr>{body}</table></div>"
    raw = "".join(f"<h3>{e(r.host)} - {e(r.cmd)}</h3><pre>{e(r.output[:20000])}</pre>" for r in app.results)
    host = app.results[-1].host if app.results else "fleet"
    return f"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>E6000 report - {e(host)}</title>
<style>body{{background:{P['bg']};color:{P['text']};font-family:'Segoe UI',system-ui,sans-serif;margin:28px;}}
h1{{margin:0 0 4px}} h2{{font-size:16px;margin:0 0 10px}} .sub{{color:{P['muted']};margin-bottom:18px}}
.kpis{{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:16px}}
.kpi,.card{{background:{P['card']};border:1px solid {P['border']};border-radius:10px;padding:14px}} .card{{margin-bottom:16px;overflow-x:auto}}
.t{{color:{P['muted']};font-size:11px;font-weight:700;text-transform:uppercase}} .v{{font-size:22px;font-weight:700;margin-top:4px}}
.s{{color:{P['muted']};font-size:11px}} table{{border-collapse:collapse;font-family:Consolas,monospace;font-size:12px}}
th,td{{border-bottom:1px solid {P['border']};padding:4px 10px;text-align:left;white-space:nowrap}} th{{color:{P['muted']}}}
pre{{background:{P['term_bg']};padding:12px;border-radius:8px;border:1px solid {P['border']};font-size:12px;overflow-x:auto}}
img{{max-width:100%;border-radius:8px}}</style></head><body>
<h1>E6000 diagnostic report</h1><div class="sub">{e(host)} · generated {datetime.now():%Y-%m-%d %H:%M:%S} · Command Center {VERSION}</div>
<div class="kpis">{kpis}</div>{findings}{fleet}{('<div class="card">' + img_html + '</div>') if img_html else ''}
{tables}<div class="card"><h2>Raw output</h2>{raw}</div></body></html>"""


def build_workbook(app: CommandCenter):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    fills = {"ok": PatternFill("solid", start_color="E6F4EA"), "warn": PatternFill("solid", start_color="FFF4D6"),
             "crit": PatternFill("solid", start_color="FDE2E1")}
    fonts = {"ok": Font(color="1A7F37"), "warn": Font(color="9A6700", bold=True), "crit": Font(color="CF222E", bold=True)}
    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", start_color="1F3A68")
    thin = Side(style="thin", color="D9DFE9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    used: set = set()

    def sheet(name: str):
        base = re.sub(r"[\[\]:*?/\\]", "_", name)[:28] or "Sheet"
        n, i = base, 2
        while n.lower() in used:
            n = f"{base[:25]}_{i}"
            i += 1
        used.add(n.lower())
        return wb.create_sheet(n)

    def write_table(ws, title: str, t: Table):
        ws.append([title])
        ws["A1"].font = Font(bold=True, size=13)
        ws.append([])
        headers = ["Health"] + [c.name or f"col {i + 1}" for i, c in enumerate(t.columns)] + ["Issues"]
        ws.append(headers)
        for c in ws[3]:
            c.font, c.fill, c.alignment = head_font, head_fill, Alignment(horizontal="center")
        for cells, sevs, rsev, iss in zip(t.rows, t.cell_sev, t.row_sev, t.row_issues):
            vals = [{"crit": "CRIT", "warn": "WARN", "ok": "OK"}.get(rsev, "")]
            for c in cells:
                s = c.replace(",", "")
                try:
                    vals.append(int(s) if re.fullmatch(r"-?\d+", s) else float(s) if re.fullmatch(r"-?\d+\.\d+", s) else c)
                except ValueError:
                    vals.append(c)
            ws.append(vals + ["; ".join(iss)])
            row = ws.max_row
            if rsev in fills:
                ws.cell(row, 1).fill, ws.cell(row, 1).font = fills[rsev], fonts[rsev]
            for ci, s in enumerate(sevs, start=2):
                if s in ("warn", "crit"):
                    ws.cell(row, ci).fill, ws.cell(row, ci).font = fills[s], fonts[s]
            for ci in range(1, len(vals) + 2):
                ws.cell(row, ci).border = border
        ws.freeze_panes = "B4"
        ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}{ws.max_row}"
        for col in ws.columns:
            w = max(len(str(c.value or "")) for c in col[2:]) if len(col) > 2 else 10
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(w + 2, 8), 60)

    def write_raw(ws, title: str, text: str):
        ws.append([title])
        ws["A1"].font = Font(bold=True, size=13)
        for ln in text.splitlines():
            ws.append([ln])
        ws.column_dimensions["A"].width = 140
        for row in ws.iter_rows(min_row=2):
            row[0].font = Font(name="Consolas", size=10)

    ws = sheet("Summary")
    ws.append(["E6000 Command Center export", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    ws["A1"].font = Font(bold=True, size=14)
    for k, (v, s, sev) in app.kpi_state.items():
        ws.append([k, v, s])
        if sev in fills:
            ws.cell(ws.max_row, 2).fill, ws.cell(ws.max_row, 2).font = fills[sev], fonts[sev]
    if app.last_triage:
        ws.append([])
        ws.append(["Dispatch", app.last_triage.dispatch, app.last_triage.dispatch_detail])
        for f in app.last_triage.findings:
            ws.append([f.sev.upper(), f.title, f.detail])
            ws.cell(ws.max_row, 1).fill, ws.cell(ws.max_row, 1).font = fills[f.sev], fonts[f.sev]
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 110

    if app.fleet_rows:
        ws = sheet("Fleet")
        ws.append(["Rank", "CMTS", "Online %", "Modems", "Offline", "Min US SNR", "Max uncorr", "Issue rows", "Verdict", "Seconds"])
        for c in ws[1]:
            c.font, c.fill = head_font, head_fill
        for i, r in enumerate(app.fleet_rows, 1):
            ws.append([i, r["host"], round(r["online"], 2) if r["online"] is not None else None, r["total"], r["offline"],
                       r["us_snr"], r["uncorr"], r["issues"], r["verdict"], round(r["secs"], 1) if r["secs"] else None])
            if r["sev"] in fills:
                for c in ws[ws.max_row]:
                    c.font = fonts[r["sev"]]
        for col, w in zip("ABCDEFGHIJ", (6, 24, 10, 10, 10, 12, 12, 11, 60, 9)):
            ws.column_dimensions[col].width = w

    sources = list(app.results) + [r for h, r in app.fleet_results.items() if r not in app.results]
    multi_host = len({r.host for r in sources}) > 1
    for n, r in enumerate(sources, 1):
        short = re.sub(r"^show\s+", "", r.cmd.split("|")[0]).strip()
        base = f"{n} {r.host} {short}" if multi_host else f"{n} {short}"
        tabs = [t for t in r.parsed.tables if t.rows]
        if tabs:
            for i, t in enumerate(tabs, 1):
                write_table(sheet(base if len(tabs) == 1 else f"{base} t{i}"), f"{r.host} - {r.cmd}", t)
        else:
            write_raw(sheet(base), f"{r.host} - {r.cmd}", r.output)
    return wb


# Colour tags that carry meaning in exported output (search highlights / selection are ignored).
_EXPORT_TAGS = ("crit", "warn", "ok", "add", "del", "mac", "ip", "iface", "th", "hdr", "cmd", "prompt", "dim")


def text_widget_runs(t: tk.Text, start: str, end: str) -> List[List[Tuple[str, Optional[str]]]]:
    """Read a Text widget range as lines of (text, colour_tag) runs, exactly as shown on screen."""
    active: List[str] = []
    lines: List[List[Tuple[str, Optional[str]]]] = [[]]
    for kind, value, _ in t.dump(start, end, text=True, tag=True):
        if kind == "tagon":
            if value in _EXPORT_TAGS:
                active.append(value)
        elif kind == "tagoff":
            if value in active:
                active.remove(value)
        elif kind == "text":
            tag = next((x for x in _EXPORT_TAGS if x in active), None)  # highest-priority tag wins
            parts = value.split("\n")
            for i, part in enumerate(parts):
                if i:
                    lines.append([])
                if part:
                    cur = lines[-1]
                    if cur and cur[-1][1] == tag:  # merge runs split by highlight/selection tags
                        cur[-1] = (cur[-1][0] + part, tag)
                    else:
                        cur.append((part, tag))
    if lines and not lines[-1]:
        lines.pop()
    return lines


def build_output_workbook(lines: List[List[Tuple[str, Optional[str]]]], meta: dict):
    """Terminal-style workbook: one row per output line in a monospace font, keeping the app's colour
    coding (Daylight palette so it reads on Excel's white background)."""
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont

    P = THEMES["Daylight"]
    colour = {"crit": P["crit"], "warn": P["warn"], "ok": P["ok"], "add": P["ok"], "del": P["crit"],
              "mac": P["mac"], "ip": P["ip"], "iface": P["iface"], "th": P["th"], "hdr": P["accent"],
              "cmd": P["text"], "prompt": P["muted"], "dim": P["dim"]}
    bold = {"crit", "warn", "th", "hdr", "cmd", "prompt"}
    mono, size = "Consolas", 10
    base_font = Font(name=mono, size=size, color="FF" + P["term_fg"].lstrip("#"))
    inline: Dict[Optional[str], InlineFont] = {None: InlineFont(rFont=mono, sz=size, color="FF" + P["term_fg"].lstrip("#"))}
    cell_font: Dict[Optional[str], Font] = {None: base_font}
    for tag, hexc in colour.items():
        argb = "FF" + hexc.lstrip("#")
        inline[tag] = InlineFont(rFont=mono, sz=size, color=argb, b=tag in bold)
        cell_font[tag] = Font(name=mono, size=size, color=argb, bold=tag in bold)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = re.sub(r"[\[\]:*?/\\]", "_", f"{meta['view']} output")[:31]
    ws.append([f"E6000 {meta['view'].lower()} output - {meta['host']}"])
    ws["A1"].font = Font(bold=True, size=13)
    scope = "selected text" if meta.get("scope") == "selection" else "full view"
    ws.append([f"Exported {datetime.now():%Y-%m-%d %H:%M:%S}  ·  {len(lines):,} lines ({scope})  ·  "
               f"Command Center {VERSION}"])
    ws["A2"].font = Font(italic=True, size=9, color="FF" + P["muted"].lstrip("#"))
    first = 4  # row 3 left blank as a spacer
    widest = 20
    for n, runs in enumerate(lines, start=first):
        runs = [(ILLEGAL_CHARACTERS_RE.sub("", s).replace("\t", "    "), tag) for s, tag in runs]
        runs = [(s, tag) for s, tag in runs if s]
        text = "".join(s for s, _ in runs)
        widest = max(widest, len(text))
        c = ws.cell(row=n, column=1)
        if not runs:
            continue
        tags = {tag for _, tag in runs}
        if len(tags) == 1:  # single colour -> plain cell, smaller file
            c.value = text[:32767]
            c.data_type = "s"  # never let a line starting with '=' become a formula
            c.font = cell_font[runs[0][1]]
        else:
            c.value = CellRichText(*[TextBlock(inline[tag], s) for s, tag in runs])
            c.font = base_font
    ws.column_dimensions["A"].width = min(widest + 2, 255)
    ws.freeze_panes = f"A{first}"
    ws.sheet_view.zoomScale = 100
    return wb


# =============================================================================
# Settings dialog
# =============================================================================
class SettingsDialog(ctk.CTkToplevel):
    def __init__(self, app: CommandCenter):
        super().__init__(app)
        self.app = app
        P = app.P
        self.title("Settings")
        self.geometry("760x720")
        self.configure(fg_color=P["bg"])
        self.transient(app)
        self.after(50, self.grab_set)
        body = ctk.CTkScrollableFrame(self, fg_color=P["bg"])
        body.pack(fill="both", expand=True, padx=12, pady=(12, 0))
        st = app.st

        def section(title):
            return app._section(body, title)

        s = section("Appearance")
        self.theme = app._option(s, list(THEMES), width=220)
        self.theme.set(st["theme"])
        self.theme.pack(anchor="w")

        s = section("Connection")
        self.hop = app._field(s, "Hop command from bastion ({host} and {user} are substituted)", mono=True)
        self.hop.insert(0, st["hop_cmd"])
        app._label(s, "VIP token entry", 11, color=P["muted"]).pack(anchor="w", pady=(6, 1))
        self.otp = app._option(s, ["append - password+token in one prompt", "separate - password then token prompt"], width=360)
        self.otp.set("separate - password then token prompt" if st["otp_mode"] == "separate" else "append - password+token in one prompt")
        self.otp.pack(anchor="w")
        grid = ctk.CTkFrame(s, fg_color="transparent")
        grid.pack(fill="x", pady=(6, 0))
        self.login_to = self._num(grid, "Login timeout (s)", st["login_timeout"], 0)
        self.cmd_to = self._num(grid, "Command timeout (s)", st["cmd_timeout"], 1)
        self.workers = self._num(grid, "Fleet parallel channels", st["fleet_workers"], 2)
        app._label(s, "Paging-disable commands sent after login (one per line, errors are ignored; --More-- is also "
                      "answered automatically)", 11, color=P["muted"], wraplength=640, justify="left").pack(anchor="w", pady=(8, 1))
        self.paging = ctk.CTkTextbox(s, height=58, fg_color=P["input"], border_color=P["border"], border_width=1,
                                     text_color=P["text"], font=app.F(11, mono=True))
        self.paging.insert("1.0", "\n".join(st["paging_cmds"]))
        self.paging.pack(fill="x")

        s = section("Behaviour")
        self.confirm = tk.BooleanVar(value=bool(st["confirm_disruptive"]))
        self.autoch = tk.BooleanVar(value=bool(st["auto_charts"]))
        for text, var in (("Confirm before sending disruptive commands (configure, clear, reset, reload, no, ...)", self.confirm),
                          ("Switch to Charts automatically when RF data arrives", self.autoch)):
            ctk.CTkSwitch(s, text=text, variable=var, font=app.F(12), text_color=P["text"], progress_color=P["accent"],
                          button_color=P["text"], fg_color=P["border"]).pack(anchor="w", pady=3)
        app._label(s, "Maintenance snapshot commands (one per line)", 11, color=P["muted"]).pack(anchor="w", pady=(8, 1))
        self.snap = ctk.CTkTextbox(s, height=58, fg_color=P["input"], border_color=P["border"], border_width=1,
                                   text_color=P["text"], font=app.F(11, mono=True))
        self.snap.insert("1.0", "\n".join(st["snapshot_cmds"]))
        self.snap.pack(fill="x")

        s = section("Colour thresholds")
        app._label(s, "Good = green, between good and marginal = amber, beyond marginal = red. "
                      "Window metrics take 'low, high'.", 11, color=P["muted"], wraplength=640, justify="left").pack(anchor="w")
        g = ctk.CTkFrame(s, fg_color="transparent")
        g.pack(fill="x", pady=(6, 0))
        for ci, h in enumerate(("Metric", "Good", "Marginal")):
            app._label(g, h, 11, "bold", P["muted"]).grid(row=0, column=ci, sticky="w", padx=4)
        self.th_entries: Dict[str, Tuple] = {}
        for ri, (key, r) in enumerate(app.th.items(), start=1):
            app._label(g, r["label"], 12).grid(row=ri, column=0, sticky="w", padx=4, pady=2)
            eg = app._entry(g, width=120, mono=True)
            ew = app._entry(g, width=120, mono=True)
            fmt = (lambda v: f"{v[0]:g}, {v[1]:g}") if r["kind"] == "window" else (lambda v: f"{v:g}")
            eg.insert(0, fmt(r["good"]))
            ew.insert(0, fmt(r["warn"]))
            eg.grid(row=ri, column=1, padx=4, pady=2)
            ew.grid(row=ri, column=2, padx=4, pady=2)
            self.th_entries[key] = (eg, ew)

        foot = ctk.CTkFrame(self, fg_color="transparent")
        foot.pack(fill="x", padx=16, pady=12)
        app._btn(foot, "Reset thresholds", self._reset_th, "ghost").pack(side="left")
        app._btn(foot, "Save", self._save, "primary", width=100).pack(side="right")
        app._btn(foot, "Cancel", self.destroy, "secondary", width=90).pack(side="right", padx=8)

    def _num(self, parent, label, value, col):
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.grid(row=0, column=col, sticky="w", padx=(0, 16))
        e = self.app._field(f, label, mono=True, width=150)
        e.insert(0, str(value))
        return e

    def _reset_th(self):
        for key, (eg, ew) in self.th_entries.items():
            r = DEFAULT_THRESHOLDS[key]
            fmt = (lambda v: f"{v[0]:g}, {v[1]:g}") if r["kind"] == "window" else (lambda v: f"{v:g}")
            eg.delete(0, "end")
            eg.insert(0, fmt(r["good"]))
            ew.delete(0, "end")
            ew.insert(0, fmt(r["warn"]))

    def _save(self):
        app = self.app
        try:
            th = json.loads(json.dumps(app.th))
            for key, (eg, ew) in self.th_entries.items():
                for name, ent in (("good", eg), ("warn", ew)):
                    raw = ent.get().strip()
                    if th[key]["kind"] == "window":
                        lo, hi = [float(x) for x in raw.split(",")]
                        th[key][name] = [min(lo, hi), max(lo, hi)]
                    else:
                        th[key][name] = float(raw)
            st = dict(app.st)
            st.update({
                "theme": self.theme.get(),
                "hop_cmd": self.hop.get().strip() or "ssh {host}",
                "otp_mode": "separate" if self.otp.get().startswith("separate") else "append",
                "login_timeout": max(5, int(float(self.login_to.get()))),
                "cmd_timeout": max(5, int(float(self.cmd_to.get()))),
                "fleet_workers": max(1, min(16, int(float(self.workers.get())))),
                "paging_cmds": [x.strip() for x in self.paging.get("1.0", "end").splitlines() if x.strip()],
                "snapshot_cmds": [x.strip() for x in self.snap.get("1.0", "end").splitlines() if x.strip()],
                "confirm_disruptive": bool(self.confirm.get()),
                "auto_charts": bool(self.autoch.get()),
            })
        except Exception as e:
            messagebox.showerror("Invalid value", f"Please check the numbers you entered.\n\n{e}", parent=self)
            return
        self.destroy()
        app.apply_settings(st, th)


if __name__ == "__main__":
    app = CommandCenter()
    app.mainloop()
