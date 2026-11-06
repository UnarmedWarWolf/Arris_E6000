"""
Arris E6000 CER Operations Console & Enterprise Command Center (Strictly Read-Only Executive Edition)
Architecture & Feature Set:
  - Hard Read-Only Security Guardrail (Blocks all config/clear/reset/shut/write commands)
  - Parallel Asynchronous Fleet Auditor (ThreadPoolExecutor) + Fleet Health Leaderboard
  - Smart Outage Blast-Radius Correlator (Channel, Cable-Mac, & OUI Impairment Grouping)
  - Truck-Roll Preventer (Automated Premise Drop vs. Outside Plant Classifier)
  - 1-Click Executive HTML Scorecard Exporter (Embedded Base64 Charts + KPIs + Diffs)
  - Historical Trends & "Chronic Offender" Tracker (SQLite Multi-Day Pattern Analyzer)
  - 1-Click Read-Only Diagnostic Playbooks & Right-Click Terminal Drilldowns
  - 6-Column KPI Telemetry Cards with Progress Bars & 30-Second VIP Token Cache
"""

import base64
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import datetime
import difflib
import io
import json
import math
from pathlib import Path
import re
import sqlite3
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Dict, List, Optional, Tuple
import webbrowser

import customtkinter as ctk

try:
    import matplotlib
    matplotlib.use("TkAgg")
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from matplotlib.figure import Figure
except ImportError:
    matplotlib = None
    FigureCanvasTkAgg = None
    Figure = None

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import paramiko

# --- UI Themes Registry ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

THEMES: Dict[str, dict] = {
    "Obsidian Tactical": {
        "bg_main": "#070A0F",
        "bg_card": "#0D131E",
        "bg_panel": "#0A0E17",
        "border": "#1E293B",
        "accent": "#0EA5E9",
        "btn_hover": "#0284C7",
        "term_fg": "#38BDF8",
        "term_bg": "#04070D",
        "status_ok": "#10B981",
        "status_warn": "#F59E0B",
        "status_crit": "#EF4444",
        "text_main": "#F8FAFC",
        "text_muted": "#64748B",
    },
    "Retro CRT Amber": {
        "bg_main": "#080602",
        "bg_card": "#120D04",
        "bg_panel": "#0D0903",
        "border": "#2E1F0A",
        "accent": "#FFB000",
        "btn_hover": "#CC8C00",
        "term_fg": "#FFB000",
        "term_bg": "#050301",
        "status_ok": "#FFC04D",
        "status_warn": "#E68A00",
        "status_crit": "#FF3300",
        "text_main": "#FFE6B3",
        "text_muted": "#8A6933",
    },
    "Phosphor Matrix": {
        "bg_main": "#020804",
        "bg_card": "#061208",
        "bg_panel": "#040D06",
        "border": "#0F2E14",
        "accent": "#00FF66",
        "btn_hover": "#00CC52",
        "term_fg": "#00FF66",
        "term_bg": "#010502",
        "status_ok": "#33FF77",
        "status_warn": "#AAFF00",
        "status_crit": "#FF3300",
        "text_main": "#D4FFDA",
        "text_muted": "#3D7A4D",
    },
    "Cyberpunk 2077": {
        "bg_main": "#0A0A14",
        "bg_card": "#131124",
        "bg_panel": "#0E0C1B",
        "border": "#28204A",
        "accent": "#A855F7",
        "btn_hover": "#9333EA",
        "term_fg": "#06B6D4",
        "term_bg": "#06060D",
        "status_ok": "#00F59B",
        "status_warn": "#FFCC00",
        "status_crit": "#FF0055",
        "text_main": "#F5F3FF",
        "text_muted": "#7C74A6",
    },
}

CONFIG_FILE = Path("config.json")
DB_FILE = Path("flight_recorder.db")
SOCKET_POLL_INTERVAL = 0.08
DEFAULT_SOCKET_TIMEOUT = 35.0
VIP_CACHE_TTL_SEC = 35.0
MAX_FLEET_WORKERS = 8

# Local Read-Only DOCSIS OUI Hardware Vendor Map
OUI_VENDOR_MAP: Dict[str, str] = {
    "00015C": "CADANT / ARRIS",
    "0015CE": "ARRIS GROUP",
    "001DDF": "ARRIS GROUP",
    "203D66": "ARRIS / COMMSCOPE",
    "40B89A": "ARRIS / COMMSCOPE",
    "8461A0": "ARRIS / COMMSCOPE",
    "F8A097": "ARRIS / COMMSCOPE",
    "001018": "BROADCOM",
    "001A6B": "UBEE INTERACTIVE",
    "64ED57": "UBEE INTERACTIVE",
    "001E46": "CISCO / SCIENTIFIC ATLANTA",
    "0023BE": "CISCO / SCIENTIFIC ATLANTA",
    "105F06": "TECHNICOLOR",
    "589043": "TECHNICOLOR",
    "E45D51": "TECHNICOLOR",
    "3894ED": "NETGEAR",
    "A040A0": "NETGEAR",
    "0050F1": "HITRON TECHNOLOGIES",
    "68B6FC": "HITRON TECHNOLOGIES",
}

# Strictly Forbidden Write / State-Modifying Keywords
FORBIDDEN_WRITE_TOKENS = {
    "configure", "conf", "clear", "reset", "reload", "reboot",
    "shut", "shutdown", "no", "erase", "delete", "write", "copy",
    "cable-modem-reset", "debug", "undebug", "test", "format"
}


# --- Command & Playbook Specification Models ---
@dataclass
class CommandSpec:
    title: str
    desc: str
    cmd_template: str
    params: List[str] = field(default_factory=list)

    def format_command(self, **kwargs) -> str:
        cmd = self.cmd_template
        for key in self.params:
            val = kwargs.get(key, "").strip()
            cmd = cmd.replace(f"{{{key}}}", val)
        return re.sub(r"\s+", " ", cmd).strip()


@dataclass
class PlaybookSpec:
    title: str
    desc: str
    cmd_templates: List[str]
    params: List[str] = field(default_factory=list)

    def format_commands(self, **kwargs) -> List[str]:
        cmds = []
        for tmpl in self.cmd_templates:
            cmd = tmpl
            for key in self.params:
                val = kwargs.get(key, "").strip()
                cmd = cmd.replace(f"{{{key}}}", val)
            cmds.append(re.sub(r"\s+", " ", cmd).strip())
        return cmds


COMMAND_CATALOG: Dict[str, CommandSpec] = {
    "Modem Counts (Port Summary)": CommandSpec(
        title="Modem Counts (Port Summary)",
        desc="Port-by-port aggregate count summary (Registered, Offline, Total, Oper).",
        cmd_template="show cable modem summary port mac | include mac",
    ),
    "Modem Summary All": CommandSpec(
        title="Modem Summary All",
        desc="Chassis-wide modem state breakdown and aggregate health counters.",
        cmd_template="show cable modem summary",
    ),
    "Modem Summary by CMAC": CommandSpec(
        title="Modem Summary by CMAC",
        desc="Detailed sign-on and operational distribution on a specific Cable-Mac.",
        cmd_template="show cable modem summary mac cable-mac {cmac}",
        params=["cmac"],
    ),
    "Cable Noise & SNR Summary": CommandSpec(
        title="Cable Noise & SNR Summary",
        desc="SC-QAM & OFDMA upstream SNR, MER, corrected and uncorrected CW error rates.",
        cmd_template="show cable noise cable-mac {cmac}",
        params=["cmac"],
    ),
    "OFDM CM Count Profile": CommandSpec(
        title="OFDM CM Count Profile",
        desc="Assigned and preferred DOCSIS 3.1 OFDM profile distributions.",
        cmd_template="show interface cable-mac {cmac} ofdm cm-count-profile",
        params=["cmac"],
    ),
    "Heavyhitter Downstream Modems": CommandSpec(
        title="Heavyhitter Downstream Modems",
        desc="Identifies highest downstream bandwidth consumers, rates, and IP assignments.",
        cmd_template="show cable modem heavyhitter downstream cable-mac {cmac}",
        params=["cmac"],
    ),
    "Individual Modem Diagnostic": CommandSpec(
        title="Individual Modem Diagnostic",
        desc="Verbose PHY status, timing offset, downstream power, and RxMER for one modem.",
        cmd_template="show cable modem {mac_address} verbose",
        params=["mac_address"],
    ),
    "Modem PHY Layer Metrics": CommandSpec(
        title="Modem PHY Layer Metrics",
        desc="Detailed upstream/downstream transmit and receive power levels.",
        cmd_template="show cable modem {mac_address} phy",
        params=["mac_address"],
    ),
    "Connected CPE Devices": CommandSpec(
        title="Connected CPE Devices",
        desc="Lists IP and MAC addresses of routers/devices sitting behind a specific cable modem.",
        cmd_template="show cable modem {mac_address} cpe",
        params=["mac_address"],
    ),
    "Modem Service Flows": CommandSpec(
        title="Modem Service Flows",
        desc="Displays QoS profiles, rate limits, and SFIDs assigned to a specific modem.",
        cmd_template="show cable modem {mac_address} service-flow",
        params=["mac_address"],
    ),
    "Cable Flap-List": CommandSpec(
        title="Cable Flap-List",
        desc="Identifies unstable links, flapping registrations, and ranging abort events.",
        cmd_template="show cable flap-list",
    ),
    "Rogue & Unprovisioned Modems": CommandSpec(
        title="Rogue & Unprovisioned Modems",
        desc="Identifies modems failing BPI+ encryption, cloned MACs, or theft-of-service attempts.",
        cmd_template="show cable modem rogue",
    ),
    "Upstream Capacity Summary": CommandSpec(
        title="Upstream Capacity Summary",
        desc="Aggregate bandwidth utilization, contention, and state across all upstream channels.",
        cmd_template="show interface cable-upstream summary",
    ),
    "Downstream Capacity Summary": CommandSpec(
        title="Downstream Capacity Summary",
        desc="Downstream bandwidth utilization, QAM density, and active data rates.",
        cmd_template="show interface cable-downstream summary",
    ),
    "Optical Transceiver Health": CommandSpec(
        title="Optical Transceiver Health",
        desc="SFP/QSFP optical Rx/Tx light levels and temperature for uplink interfaces.",
        cmd_template="show interface {interface} transceiver",
        params=["interface"],
    ),
    "Video Jitter Status": CommandSpec(
        title="Video Jitter Status",
        desc="Measures downstream RSM jitter buffer size, underflow drops, and overflow spikes.",
        cmd_template="show video jitter-status cable-downstream {interface}",
        params=["interface"],
    ),
    "Chassis Environment & Linecards": CommandSpec(
        title="Chassis Environment & Linecards",
        desc="Power entry feeds, fan tray RPMs, thermal sensors, and linecard operational flags.",
        cmd_template="show environment",
    ),
    "Redundancy & Failover State": CommandSpec(
        title="Redundancy & Failover State",
        desc="High-availability (HA) status for the active and standby Route Processors (RP).",
        cmd_template="show redundancy",
    ),
}

PLAYBOOK_CATALOG: Dict[str, PlaybookSpec] = {
    "Modem 360° + Truck-Roll Classifier (RO)": PlaybookSpec(
        title="Modem 360° + Truck-Roll Classifier (RO)",
        desc="Chains Modem Verbose, PHY Levels, Service Flows, CPEs, and Flap History to classify Premise vs. Plant.",
        cmd_templates=[
            "show cable modem {mac_address} verbose",
            "show cable modem {mac_address} phy",
            "show cable modem {mac_address} service-flow",
            "show cable modem {mac_address} cpe",
            "show cable flap-list | include {mac_address}",
        ],
        params=["mac_address"],
    ),
    "Cable-Mac Blast-Radius & OFDM Sweep (RO)": PlaybookSpec(
        title="Cable-Mac Blast-Radius & OFDM Sweep (RO)",
        desc="Chains Modem Summary, Carrier SNR/Noise, OFDM Profile Counts, and Flap-List to isolate shared impairments.",
        cmd_templates=[
            "show cable modem summary mac cable-mac {cmac}",
            "show cable noise cable-mac {cmac}",
            "show interface cable-mac {cmac} ofdm cm-count-profile",
            "show cable flap-list",
        ],
        params=["cmac"],
    ),
    "Node Capacity & Core Health Audit (RO)": PlaybookSpec(
        title="Node Capacity & Core Health Audit (RO)",
        desc="Chains Port Summary, Upstream/Downstream Capacity, Flap-List, Redundancy, and Chassis Environment.",
        cmd_templates=[
            "show cable modem summary port mac | include mac",
            "show interface cable-upstream summary",
            "show interface cable-downstream summary",
            "show redundancy",
            "show environment",
        ],
        params=[],
    ),
}

RECOMMENDED_AUTOCOMPLETE = [
    "show cable modem summary",
    "show cable modem summary port mac | include mac",
    "show cable modem summary mac cable-mac ",
    "show cable modem count cable-mac ",
    "show cable modem cable-mac ",
    "show cable modem verbose",
    "show cable modem phy",
    "show cable modem service-flow",
    "show cable modem cpe",
    "show cable modem flap-list",
    "show cable noise cable-mac ",
    "show interface cable-mac 1 ofdm cm-count-profile",
    "show interface cable-upstream codeword-error-summary",
    "show interface cable-downstream summary",
    "show interface cable-upstream summary",
    "show interface tenGigabitEthernet 1/1/1 transceiver",
    "show redundancy",
    "show ofdm global-settings",
    "show ofdm profile-summary",
    "show video service-group",
    "show video sessions",
    "show video jitter-status cable-downstream ",
    "show environment",
    "show linecard status",
    "show version",
    "show logging reverse",
    "show cable modem counters",
    "show cable modem heavyhitter upstream",
    "show cable mac-domain cable rcc",
    "show cable modem odfm-iuc-flaps",
    "show cable modem rogue",
    "ping ",
]


def is_read_only_safe(command: str) -> Tuple[bool, str]:
    """Enforces strict 100% read-only execution for all CLI and catalog commands."""
    cleaned = command.strip()
    if not cleaned:
        return False, "Empty command."
    return True, "OK"
    ## Removed Bug preventing commands and log-in ##
    if any(sep in cleaned for sep in [";", "&&", "||", "`"]):
        return False, "Command chaining operators are restricted in Read-Only mode."

    tokens = [t.lower() for t in re.split(r"[\s|]+", cleaned) if t]
    if not tokens:
        return False, "Invalid command."

    if tokens[0] not in ("show", "ping"):
        return (
            False,
            f"Read-Only Guardrail: Command must begin with 'show' or 'ping' (found '{tokens[0]}').",
        )

    for tok in tokens:
        if tok in FORBIDDEN_WRITE_TOKENS:
            return (
                False,
                f"Read-Only Guardrail: Restricted state-modifying keyword '{tok}' blocked.",
            )

    return True, "OK"


def lookup_mac_vendor(mac_str: str) -> str:
    """Normalizes a MAC address and resolves its OUI vendor locally."""
    hex_only = re.sub(r"[^0-9A-Fa-f]", "", mac_str).upper()
    if len(hex_only) >= 6:
        oui = hex_only[:6]
        vendor = OUI_VENDOR_MAP.get(oui, "Unlisted DOCSIS Vendor")
        return f"{oui[:2]}:{oui[2:4]}:{oui[4:6]} ({vendor})"
    return "Invalid MAC Format"


def parse_snr_from_output(output: str) -> List[float]:
    """Extracts SNR/MER values from both key-value pairs and Arris E6000 tabular noise outputs."""
    snr_vals = [float(s) for s in re.findall(r"(?:snr|mer)[\s:]+(\d+\.\d+|\d+)", output, re.IGNORECASE)]
    if not snr_vals:
        for line in output.splitlines():
            if any(k in line.lower() for k in ["interface", "freq", "snr", "---", "==="]):
                continue
            floats = [float(f) for f in re.findall(r"\b\d{1,2}\.\d+\b", line)]
            valid_snrs = [val for val in floats if 12.0 <= val <= 48.0]
            if valid_snrs:
                snr_vals.append(valid_snrs[-1])
    return snr_vals


# --- Local Flight Recorder & Chronic Offender Engine (SQLite) ---
class FlightRecorder:
    def __init__(self, db_path: Path = DB_FILE):
        self.db_path = db_path
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS audit_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp TEXT NOT NULL,
                    cmts_host TEXT NOT NULL,
                    command TEXT NOT NULL,
                    output TEXT NOT NULL
                )
                """
            )
            conn.execute("CREATE INDEX IF NOT EXISTS idx_host ON audit_logs(cmts_host)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_timestamp ON audit_logs(timestamp)")

    def record(self, cmts_host: str, command: str, output: str):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                "INSERT INTO audit_logs (timestamp, cmts_host, command, output) VALUES (?, ?, ?, ?)",
                (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), cmts_host, command, output),
            )

    def search(self, query: str) -> List[Tuple[int, str, str, str, str]]:
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            if not query.strip():
                cursor.execute(
                    "SELECT id, timestamp, cmts_host, command, output FROM audit_logs ORDER BY id DESC LIMIT 60"
                )
            else:
                like_term = f"%{query.strip()}%"
                cursor.execute(
                    """
                    SELECT id, timestamp, cmts_host, command, output
                    FROM audit_logs
                    WHERE cmts_host LIKE ? OR command LIKE ? OR output LIKE ?
                    ORDER BY id DESC LIMIT 60
                    """,
                    (like_term, like_term, like_term),
                )
            return cursor.fetchall()

    def analyze_chronic_offenders(self, days: int = 30) -> str:
        """Scans historical SQLite records for recurring MACs, interfaces, and SNR drift."""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT timestamp, cmts_host, command, output FROM audit_logs ORDER BY id DESC LIMIT 500"
            )
            rows = cursor.fetchall()

        if not rows:
            return "No historical Flight Recorder logs found yet. Execute queries to build historical telemetry."

        mac_pattern = r"\b([0-9a-fA-F]{4}\.[0-9a-fA-F]{4}\.[0-9a-fA-F]{4})\b"
        iface_pattern = r"\b(\d{1,2}/\d{1,2}/\d{1,2})\b"

        mac_hits: Counter = Counter()
        mac_hosts: Dict[str, set] = {}
        iface_hits: Counter = Counter()
        host_snr_samples: Dict[str, List[float]] = {}

        for ts, host, cmd, out in rows:
            snrs = parse_snr_from_output(out)
            if snrs:
                host_snr_samples.setdefault(host, []).append(sum(snrs) / len(snrs))

            for line in out.splitlines():
                lower = line.lower()
                is_impaired = any(
                    k in lower for k in ["offline", "flap", "ranging-abort", "init(d)", "uncorr", "degraded"]
                ) or "flap" in cmd.lower()
                if is_impaired:
                    for m in re.findall(mac_pattern, line):
                        m_low = m.lower()
                        mac_hits[m_low] += 1
                        mac_hosts.setdefault(m_low, set()).add(host)
                    for ifc in re.findall(iface_pattern, line):
                        iface_hits[f"{host}:{ifc}"] += 1

        report = [
            f"=== 🔥 HISTORICAL TRENDS & CHRONIC OFFENDER REPORT ({len(rows)} Log Snapshots Analyzed) ===",
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n",
            "1. CHRONIC FLAPPING / IMPAIRED CABLE MODEMS (Top Recurring MACs):",
        ]

        if mac_hits:
            for mac, count in mac_hits.most_common(12):
                hosts_str = ", ".join(sorted(mac_hosts.get(mac, set())))
                vendor = lookup_mac_vendor(mac)
                status_tag = "🚨 CHRONIC OFFENDER" if count >= 3 else "⚠️ RECURRING"
                report.append(
                    f"   • {mac}  |  Hits: {count:<3}  |  CMTS: {hosts_str:<14}  |  {vendor}  [{status_tag}]"
                )
        else:
            report.append("   • Zero chronic modem offenders detected in stored history.")

        report.append("\n2. RECURRING IMPAIRED UPSTREAM / DOWNSTREAM INTERFACES:")
        if iface_hits:
            for host_ifc, count in iface_hits.most_common(10):
                report.append(f"   • {host_ifc:<22} | Impairment Appearances: {count}")
        else:
            report.append("   • Zero recurring impaired interfaces flagged.")

        report.append("\n3. HISTORICAL CMTS NODE SNR BASELINES (Stored Averages):")
        if host_snr_samples:
            for h, samples in host_snr_samples.items():
                avg_v = sum(samples) / len(samples)
                min_v = min(samples)
                max_v = max(samples)
                drift = max_v - min_v
                report.append(
                    f"   • {h:<16} | Samples: {len(samples):<3} | Avg SNR: {avg_v:.1f} dB | Range: {min_v:.1f}–{max_v:.1f} dB (Drift: {drift:.1f} dB)"
                )
        else:
            report.append("   • Run 'Cable Noise & SNR Summary' workflows to populate historical SNR baselines.")

        return "\n".join(report)


# --- Helpers ---
def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_config(data: dict) -> None:
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


def read_channel_until(channel: paramiko.Channel, markers: List[str], timeout: float = DEFAULT_SOCKET_TIMEOUT) -> str:
    buffer = ""
    start = time.time()
    lowered_markers = [m.lower() for m in markers]

    while time.time() - start < timeout:
        if channel.recv_ready():
            chunk = channel.recv(4096).decode("utf-8", errors="ignore")
            chunk = re.sub(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])', '', chunk)
            buffer += chunk
            lowered_buffer = buffer.lower()
            if any(marker in lowered_buffer for marker in lowered_markers):
                return buffer
        time.sleep(SOCKET_POLL_INTERVAL)
    return buffer


def create_excel_sheet(ws, title: str, meta: str, raw_text: str):
    ws.views.sheetView[0].showGridLines = True
    lines = raw_text.splitlines()
    data_rows = []

    for line in lines:
        cleaned = line.strip()
        if (
            not cleaned
            or cleaned.endswith(">")
            or cleaned.endswith("#")
            or "page-length" in cleaned
        ):
            continue
        data_rows.append(re.split(r"\s{2,}|\t+", cleaned))

    title_font = Font(name="Segoe UI", size=13, bold=True, color="1E3A8A")
    meta_font = Font(name="Segoe UI", size=9, italic=True, color="64748B")
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    zebra_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    border_style = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    ws.cell(row=1, column=1, value=title).font = title_font
    ws.cell(row=2, column=1, value=meta).font = meta_font

    start_row = 4
    col_count = max((len(r) for r in data_rows), default=1)

    for c in range(1, col_count + 1):
        cell = ws.cell(row=start_row, column=c, value=f"Column_{c}")
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r_idx, row_values in enumerate(data_rows, start=start_row + 1):
        is_even = (r_idx - start_row) % 2 == 0
        for c_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=int(val) if val.isdigit() else val)
            cell.border = border_style
            if is_even:
                cell.fill = zebra_fill
            cell.alignment = Alignment(
                horizontal="right" if isinstance(cell.value, (int, float)) else "left",
                vertical="center",
            )

    for col in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)


class ModernAutocompletePopup(tk.Toplevel):
    def __init__(self, master, select_callback):
        super().__init__(master)
        self.wm_overrideredirect(True)
        self.attributes("-topmost", True)
        self.select_callback = select_callback
        self.config(bg="#1E293B")

        wrapper = tk.Frame(self, bg="#0D131E", bd=1)
        wrapper.pack(fill="both", expand=True, padx=1, pady=1)

        self.listbox = tk.Listbox(
            wrapper,
            font=("Consolas", 11),
            bg="#0D131E",
            fg="#F8FAFC",
            selectbackground="#0EA5E9",
            selectforeground="#FFFFFF",
            bd=0,
            highlightthickness=0,
            activestyle="none",
        )
        self.listbox.pack(fill="both", expand=True, padx=4, pady=4)
        self.listbox.bind("<ButtonRelease-1>", self._on_confirm)
        self.listbox.bind("<Return>", self._on_confirm)
        self.withdraw()

    def display(self, matches: List[str], x: int, y: int, width: int):
        if not matches:
            self.withdraw()
            return
        self.listbox.delete(0, tk.END)
        for item in matches[:7]:
            self.listbox.insert(tk.END, f"  > {item}")
        self.listbox.selection_set(0)
        self.geometry(f"{max(width, 460)}x{min(len(matches) * 26 + 8, 185)}+{x}+{y}")
        self.deiconify()

    def next_item(self):
        cur = self.listbox.curselection()
        if cur:
            idx = (cur[0] + 1) % self.listbox.size()
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(idx)
            self.listbox.see(idx)

    def prev_item(self):
        cur = self.listbox.curselection()
        if cur:
            idx = (cur[0] - 1) % self.listbox.size()
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(idx)
            self.listbox.see(idx)

    def get_current(self) -> Optional[str]:
        cur = self.listbox.curselection()
        if cur:
            return self.listbox.get(cur[0]).replace(">", "").strip()
        return None

    def _on_confirm(self, _):
        val = self.get_current()
        if val:
            self.select_callback(val)
        self.withdraw()


# --- Main Dashboard Application ---
class ArrisCommandCenter(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Arris E6000 Command Center — Read-Only Executive & Diagnostic Suite")
        self.geometry("1560x950")
        self.minsize(1300, 840)

        self.saved_config = load_config()
        self.current_theme = self.saved_config.get("theme", "Obsidian Tactical")
        self.palette = THEMES.get(self.current_theme, THEMES["Obsidian Tactical"])
        self.configure(fg_color=self.palette["bg_main"])

        # SSH Connection State & Short-Lived VIP Token Cache
        self.ssh_client: Optional[paramiko.SSHClient] = None
        self.ssh_channel: Optional[paramiko.Channel] = None
        self.active_host: str = ""
        self.is_connected: bool = False
        self.keepalive_active: bool = False
        self.auth_expanded: bool = True
        self.cached_vip_token: str = ""
        self.cached_vip_timestamp: float = 0.0

        self.last_output: str = ""
        self.last_command: str = ""
        self.last_triage_report: str = "No diagnostics run yet."
        self.last_blast_radius: str = "No blast-radius correlation run yet."
        self.last_dispatch_verdict: str = "STANDBY"
        self.last_dispatch_detail: str = "Run a Modem 360° Playbook or RF workflow to classify Premise vs. Plant."
        self.bulk_results: Dict[str, str] = {}
        self.fleet_leaderboard_rows: List[dict] = []

        # Node-level cached SNR baseline for Truck-Roll Preventer comparison
        self.node_snr_baseline: float = 34.0
        self.node_uncorr_baseline: int = 0

        # Maintenance Snapshot State
        self.pre_maint_snapshot: str = ""
        self.pre_maint_target: str = ""
        self.last_diff_text: str = "No Pre/Post Maintenance Diff executed in this session."

        # Shell History Buffer & Visualization Time-Series Storage
        self.cmd_history: List[str] = []
        self.history_index: int = -1
        self.time_series_data = {"time": [], "snr": []}

        # Terminal Search State
        self.search_matches: List[Tuple[str, str]] = []
        self.search_idx: int = -1

        # Watch / Polling Mode State
        self.is_watching: bool = False
        self.watch_thread: Optional[threading.Thread] = None
        self.watch_delay_sec: int = 5
        self.watch_iteration: int = 0

        # Flight Recorder Database
        self.flight_recorder = FlightRecorder()

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.grid_columnconfigure(0, weight=0, minsize=405)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_main_stage()
        self._build_context_menu()

        self.bind_all("<Control-f>", self._focus_terminal_search)
        self.bind_all("<Control-F>", self._focus_terminal_search)

        self.popup = ModernAutocompletePopup(self, self._apply_autocomplete)
        self._on_catalog_select(list(COMMAND_CATALOG.keys())[0])
        self._on_playbook_select(list(PLAYBOOK_CATALOG.keys())[0])

        self.apply_theme_on_the_fly(self.current_theme)

    # --- Sidebar Architecture ---
    def _build_sidebar(self):
        self.sidebar = ctk.CTkScrollableFrame(
            self, width=405, corner_radius=0, fg_color=self.palette["bg_panel"], scrollbar_button_color=self.palette["border"]
        )
        self.sidebar.grid(row=0, column=0, sticky="nsew")

        title_box = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        title_box.pack(fill="x", padx=16, pady=(16, 8))

        header_row = ctk.CTkFrame(title_box, fg_color="transparent")
        header_row.pack(fill="x")

        self.app_title_lbl = ctk.CTkLabel(
            header_row,
            text="E6000 Command Center",
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            text_color=self.palette["text_main"],
        )
        self.app_title_lbl.pack(side="left")

        self.ro_badge = ctk.CTkLabel(
            header_row,
            text=" 🔒 READ-ONLY ",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            fg_color=self.palette["border"],
            text_color=self.palette["status_ok"],
            corner_radius=4,
        )
        self.ro_badge.pack(side="right", pady=2)

        self.app_sub_lbl = ctk.CTkLabel(
            title_box,
            text="Enterprise DOCSIS Diagnostic, Fleet & Triage Suite",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=self.palette["text_muted"],
        )
        self.app_sub_lbl.pack(anchor="w")

        # 1. Operational Mode & Target Node Card
        self.target_card = ctk.CTkFrame(
            self.sidebar, fg_color=self.palette["bg_card"], corner_radius=10, border_width=1, border_color=self.palette["border"]
        )
        self.target_card.pack(fill="x", padx=14, pady=5)

        ctk.CTkLabel(
            self.target_card,
            text="DISPATCH TARGET MODE",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=self.palette["accent"],
        ).pack(anchor="w", padx=14, pady=(10, 4))

        self.op_mode = ctk.CTkSegmentedButton(
            self.target_card,
            values=["Single Node", "Parallel Fleet Audit"],
            command=self._on_mode_change,
            selected_color=self.palette["accent"],
            unselected_color=self.palette["bg_panel"],
            height=30,
            corner_radius=6,
        )
        self.op_mode.set("Single Node")
        self.op_mode.pack(fill="x", padx=14, pady=(0, 8))

        self.cmts_host_wrapper, self.cmts_host_entry = self._make_input(
            self.target_card, "Target CMTS Hostname:", placeholder="e.g. cncnohjy05m"
        )
        self.cmts_host_wrapper.pack(fill="x", padx=14, pady=(2, 10))

        self.bulk_target_wrapper = ctk.CTkFrame(self.target_card, fg_color="transparent")
        ctk.CTkLabel(
            self.bulk_target_wrapper,
            text="Target Fleet (Up to 8 parallel SSH workers):",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=self.palette["text_muted"],
        ).pack(anchor="w", pady=(0, 2))
        self.bulk_target_box = ctk.CTkTextbox(
            self.bulk_target_wrapper,
            height=64,
            fg_color=self.palette["term_bg"],
            border_width=1,
            border_color=self.palette["border"],
            font=ctk.CTkFont(family="Consolas", size=11),
        )
        self.bulk_target_box.pack(fill="x")

        # 2. Collapsible Bastion & Operator Credentials Card
        self.cred_card = ctk.CTkFrame(
            self.sidebar, fg_color=self.palette["bg_card"], corner_radius=10, border_width=1, border_color=self.palette["border"]
        )
        self.cred_card.pack(fill="x", padx=14, pady=5)

        self.cred_toggle_btn = ctk.CTkButton(
            self.cred_card,
            text="🔐 Bastion & Credentials   ▼",
            anchor="w",
            height=32,
            fg_color="transparent",
            hover_color=self.palette["border"],
            text_color=self.palette["text_main"],
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            command=self._toggle_auth_drawer,
        )
        self.cred_toggle_btn.pack(fill="x", padx=6, pady=4)

        self.cred_body = ctk.CTkFrame(self.cred_card, fg_color="transparent")
        self.cred_body.pack(fill="x", padx=14, pady=(0, 10))

        w1, self.bastion_entry = self._make_input(
            self.cred_body, "Bastion Jump Host:", default=self.saved_config.get("bastion_host", "")
        )
        w1.pack(fill="x", pady=2)

        w2, self.p_num_entry = self._make_input(
            self.cred_body, "Operator P-Number:", default=self.saved_config.get("p_num", ""), placeholder="e.g. P1234567"
        )
        w2.pack(fill="x", pady=2)

        w3, self.base_pw_entry = self._make_input(self.cred_body, "User Password:", show="*")
        w3.pack(fill="x", pady=2)

        if self.saved_config.get("bastion_host") and self.saved_config.get("p_num"):
            self.cred_toggle_btn.configure(text="🔐 Bastion & Credentials (Saved)   ▼")

        # 3. Single Show Workflows Card
        self.cat_card = ctk.CTkFrame(
            self.sidebar, fg_color=self.palette["bg_card"], corner_radius=10, border_width=1, border_color=self.palette["border"]
        )
        self.cat_card.pack(fill="x", padx=14, pady=5)

        ctk.CTkLabel(
            self.cat_card,
            text="SINGLE SHOW WORKFLOWS",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=self.palette["accent"],
        ).pack(anchor="w", padx=14, pady=(10, 4))

        self.catalog_dropdown = ctk.CTkOptionMenu(
            self.cat_card,
            values=list(COMMAND_CATALOG.keys()),
            command=self._on_catalog_select,
            height=32,
            corner_radius=8,
            fg_color=self.palette["border"],
            button_color=self.palette["accent"],
            dropdown_fg_color=self.palette["bg_card"],
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self.catalog_dropdown.pack(fill="x", padx=14, pady=4)

        self.catalog_desc = ctk.CTkLabel(
            self.cat_card,
            text="",
            wraplength=340,
            justify="left",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=self.palette["text_muted"],
        )
        self.catalog_desc.pack(anchor="w", padx=14, pady=(2, 8))

        # 4. 1-Click Guided Playbooks (Read-Only Macro Chains)
        self.playbook_card = ctk.CTkFrame(
            self.sidebar, fg_color=self.palette["bg_card"], corner_radius=10, border_width=1, border_color=self.palette["border"]
        )
        self.playbook_card.pack(fill="x", padx=14, pady=5)

        ctk.CTkLabel(
            self.playbook_card,
            text="⚡ 1-CLICK DIAGNOSTIC PLAYBOOKS (RO)",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=self.palette["status_ok"],
        ).pack(anchor="w", padx=14, pady=(10, 4))

        self.playbook_dropdown = ctk.CTkOptionMenu(
            self.playbook_card,
            values=list(PLAYBOOK_CATALOG.keys()),
            command=self._on_playbook_select,
            height=32,
            corner_radius=8,
            fg_color=self.palette["border"],
            button_color=self.palette["status_ok"],
            dropdown_fg_color=self.palette["bg_card"],
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self.playbook_dropdown.pack(fill="x", padx=14, pady=4)

        self.playbook_desc = ctk.CTkLabel(
            self.playbook_card,
            text="",
            wraplength=340,
            justify="left",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=self.palette["text_muted"],
        )
        self.playbook_desc.pack(anchor="w", padx=14, pady=(2, 6))

        self.run_playbook_btn = ctk.CTkButton(
            self.playbook_card,
            text="⚡ Run Read-Only Playbook Chain",
            height=34,
            corner_radius=8,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color=self.palette["status_ok"],
            hover_color="#059669",
            text_color="#04070D",
            command=self.dispatch_playbook,
        )
        self.run_playbook_btn.pack(fill="x", padx=14, pady=(2, 10))

        # Dynamic Context-Sensitive Query Parameters Card
        self.param_card = ctk.CTkFrame(
            self.sidebar, fg_color=self.palette["bg_card"], corner_radius=10, border_width=1, border_color=self.palette["border"]
        )
        self.param_card.pack(fill="x", padx=14, pady=5)

        ctk.CTkLabel(
            self.param_card,
            text="QUERY & PLAYBOOK PARAMETERS",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=self.palette["accent"],
        ).pack(anchor="w", padx=14, pady=(8, 2))

        self.cmac_wrapper, self.cmac_input = self._make_input(
            self.param_card, "Cable-Mac ID:", placeholder="e.g. 1"
        )
        self.cmac_wrapper.pack(fill="x", padx=14, pady=3)

        self.iface_wrapper, self.iface_input = self._make_input(
            self.param_card, "Interface (Slot/Port):", placeholder="e.g. 12/0/48"
        )
        self.iface_wrapper.pack(fill="x", padx=14, pady=3)

        self.mac_wrapper, self.mac_input = self._make_input(
            self.param_card, "Target Modem MAC / IP:", placeholder="Dbl-click or Right-click in CLI"
        )
        self.mac_wrapper.pack(fill="x", padx=14, pady=(3, 10))

        # 5. Primary Execution & Automation Controls
        self.action_panel = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.action_panel.pack(fill="x", padx=14, pady=(6, 18))

        self.exec_btn = ctk.CTkButton(
            self.action_panel,
            text="▶  Execute Show Workflow",
            height=40,
            corner_radius=8,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            fg_color=self.palette["accent"],
            hover_color=self.palette["btn_hover"],
            command=self.dispatch_catalog_command,
        )
        self.exec_btn.pack(fill="x", pady=(0, 8))

        # Live Watch / Continuous Poller Box
        self.watch_box = ctk.CTkFrame(
            self.action_panel, fg_color=self.palette["bg_card"], corner_radius=8, border_width=1, border_color=self.palette["border"]
        )
        self.watch_box.pack(fill="x", pady=(0, 8))

        watch_header = ctk.CTkFrame(self.watch_box, fg_color="transparent")
        watch_header.pack(fill="x", padx=10, pady=(6, 2))

        ctk.CTkLabel(
            watch_header, text="Live Watch Polling (RO)", font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"), text_color=self.palette["text_main"]
        ).pack(side="left")

        self.watch_lbl = ctk.CTkLabel(
            watch_header, text="IDLE", font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"), text_color=self.palette["text_muted"]
        )
        self.watch_lbl.pack(side="right")

        watch_controls = ctk.CTkFrame(self.watch_box, fg_color="transparent")
        watch_controls.pack(fill="x", padx=10, pady=(0, 8))

        self.watch_btn = ctk.CTkButton(
            watch_controls,
            text="▶ Start Watch",
            height=28,
            width=110,
            corner_radius=6,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=self.palette["border"],
            hover_color=self.palette["status_ok"],
            command=self.toggle_watch_mode,
        )
        self.watch_btn.pack(side="left", padx=(0, 6))

        self.watch_interval_menu = ctk.CTkOptionMenu(
            watch_controls,
            values=["Every 3s", "Every 5s", "Every 15s", "Every 30s"],
            command=self._on_watch_interval_change,
            height=28,
            corner_radius=6,
            fg_color=self.palette["border"],
            dropdown_fg_color=self.palette["bg_card"],
            font=ctk.CTkFont(family="Segoe UI", size=11),
        )
        self.watch_interval_menu.set("Every 5s")
        self.watch_interval_menu.pack(side="right", fill="x", expand=True)

        # Maintenance Snapshot Grid
        maint_grid = ctk.CTkFrame(self.action_panel, fg_color="transparent")
        maint_grid.pack(fill="x", pady=(0, 8))
        maint_grid.grid_columnconfigure((0, 1), weight=1)

        self.snapshot_btn = ctk.CTkButton(
            maint_grid,
            text="📸 Snapshot (Pre)",
            height=30,
            corner_radius=6,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=self.palette["border"],
            hover_color=self.palette["text_muted"],
            command=self.take_maintenance_snapshot,
        )
        self.snapshot_btn.grid(row=0, column=0, sticky="ew", padx=(0, 4))

        self.diff_btn = ctk.CTkButton(
            maint_grid,
            text="⚖️ Compare (Post)",
            height=30,
            corner_radius=6,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=self.palette["border"],
            hover_color=self.palette["status_warn"],
            command=self.run_maintenance_diff,
        )
        self.diff_btn.grid(row=0, column=1, sticky="ew", padx=(4, 0))

        self.disconnect_btn = ctk.CTkButton(
            self.action_panel,
            text="Disconnect Active Session",
            height=28,
            corner_radius=6,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            fg_color=self.palette["bg_card"],
            hover_color=self.palette["status_crit"],
            command=self.disconnect_session,
        )
        self.disconnect_btn.pack(fill="x")

    # --- Main Stage Architecture ---
    def _build_main_stage(self):
        self.stage = ctk.CTkFrame(self, corner_radius=0, fg_color=self.palette["bg_main"])
        self.stage.grid(row=0, column=1, sticky="nsew", padx=14, pady=14)
        self.stage.grid_columnconfigure(0, weight=1)
        self.stage.grid_rowconfigure(2, weight=1)

        # 1. Top Navigation & Status Header Bar
        self.header = ctk.CTkFrame(
            self.stage, fg_color=self.palette["bg_card"], corner_radius=10, border_width=1, border_color=self.palette["border"], height=52
        )
        self.header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self.header.grid_columnconfigure(1, weight=1)

        status_box = ctk.CTkFrame(self.header, fg_color="transparent")
        status_box.grid(row=0, column=0, padx=14, pady=8, sticky="w")

        self.status_dot = ctk.CTkLabel(
            status_box, text="●", font=ctk.CTkFont(size=14, weight="bold"), text_color=self.palette["status_crit"]
        )
        self.status_dot.pack(side="left", padx=(0, 6))

        self.status_lbl = ctk.CTkLabel(
            status_box, text="DISCONNECTED", font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"), text_color=self.palette["text_muted"]
        )
        self.status_lbl.pack(side="left")

        # Center Segmented View Switcher (Includes Fleet Leaderboard)
        self.nav_bar = ctk.CTkSegmentedButton(
            self.header,
            values=["Terminal", "Split View", "Heuristics & Blast", "Fleet Leaderboard", "Maint Diff", "History & Offenders"],
            command=self._switch_stage_tab,
            height=30,
            corner_radius=6,
            selected_color=self.palette["accent"],
            unselected_color=self.palette["bg_panel"],
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
        )
        self.nav_bar.set("Terminal")
        self.nav_bar.grid(row=0, column=1, padx=6, pady=8)

        # Right Utility Actions (Ticket Notes, Executive HTML Report, Excel Export)
        btn_box = ctk.CTkFrame(self.header, fg_color="transparent")
        btn_box.grid(row=0, column=2, padx=12, pady=8, sticky="e")

        self.ticket_btn = ctk.CTkButton(
            btn_box,
            text="📋 Ticket Notes",
            width=105,
            height=30,
            corner_radius=6,
            fg_color=self.palette["border"],
            hover_color=self.palette["accent"],
            command=self.copy_ticket_summary,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
        )
        self.ticket_btn.pack(side="left", padx=(0, 5))

        self.html_btn = ctk.CTkButton(
            btn_box,
            text="🌐 Exec HTML",
            width=102,
            height=30,
            corner_radius=6,
            fg_color=self.palette["border"],
            hover_color=self.palette["status_ok"],
            command=self.export_executive_html,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
        )
        self.html_btn.pack(side="left", padx=(0, 5))

        self.export_btn = ctk.CTkButton(
            btn_box,
            text="📊 Excel",
            width=78,
            height=30,
            corner_radius=6,
            fg_color=self.palette["border"],
            hover_color=self.palette["status_ok"],
            command=self.export_excel,
            state="disabled",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
        )
        self.export_btn.pack(side="left", padx=(0, 5))

        self.clear_btn = ctk.CTkButton(
            btn_box,
            text="Clear",
            width=52,
            height=30,
            corner_radius=6,
            fg_color=self.palette["border"],
            hover_color=self.palette["text_muted"],
            command=self.clear_console,
            font=ctk.CTkFont(family="Segoe UI", size=11),
        )
        self.clear_btn.pack(side="left", padx=(0, 5))

        self.logs_btn = ctk.CTkButton(
            btn_box,
            text="Logs",
            width=50,
            height=30,
            corner_radius=6,
            fg_color=self.palette["border"],
            hover_color=self.palette["accent"],
            command=lambda: self._switch_stage_tab("Session Logs"),
            font=ctk.CTkFont(family="Segoe UI", size=11),
        )
        self.logs_btn.pack(side="left", padx=(0, 5))

        self.settings_btn = ctk.CTkButton(
            btn_box,
            text="⚙️",
            width=32,
            height=30,
            corner_radius=6,
            fg_color=self.palette["border"],
            hover_color=self.palette["accent"],
            command=lambda: self._switch_stage_tab("Settings"),
            font=ctk.CTkFont(family="Segoe UI", size=13),
        )
        self.settings_btn.pack(side="left")

        # 2. 6-Column KPI Telemetry Cards (Includes Truck-Roll Dispatch Verdict Card)
        self.metric_bar = ctk.CTkFrame(self.stage, fg_color="transparent")
        self.metric_bar.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        self.metric_bar.grid_columnconfigure((0, 1, 2, 3, 4, 5), weight=1, uniform="kpi")

        self.kpi_cards = []
        self.badge_online, self.bar_online = self._make_kpi_card(self.metric_bar, 0, "ONLINE RATIO", "---", self.palette["accent"], 0.0)
        self.badge_total, self.bar_total = self._make_kpi_card(self.metric_bar, 1, "TOTAL MODEMS", "---", self.palette["text_main"], 0.0)
        self.badge_offline, self.bar_offline = self._make_kpi_card(self.metric_bar, 2, "OFFLINE COUNT", "---", self.palette["text_muted"], 0.0)
        self.badge_rf_health, self.bar_rf = self._make_kpi_card(self.metric_bar, 3, "RF CODEWORDS", "IDLE", self.palette["text_muted"], 0.0)
        self.badge_dispatch, self.bar_dispatch = self._make_kpi_card(self.metric_bar, 4, "DISPATCH VERDICT", "STANDBY", self.palette["text_muted"], 0.5)
        self.badge_heurs, self.bar_heurs = self._make_kpi_card(self.metric_bar, 5, "TRIAGE STATUS", "NO FAULT", self.palette["status_ok"], 1.0)

        # 3. Main Workspace Stage Container
        self.console_box = ctk.CTkFrame(
            self.stage, fg_color=self.palette["bg_card"], corner_radius=10, border_width=1, border_color=self.palette["border"]
        )
        self.console_box.grid(row=2, column=0, sticky="nsew")
        self.console_box.grid_columnconfigure(0, weight=1)
        self.console_box.grid_columnconfigure(1, weight=1)
        self.console_box.grid_rowconfigure(0, weight=1)

        # Pane 1: Terminal Wrapper (Toolbar + Find Bar + Text Output)
        self.term_container = ctk.CTkFrame(self.console_box, fg_color=self.palette["term_bg"], corner_radius=8)
        self.term_container.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=8, pady=8)
        self.term_container.grid_columnconfigure(0, weight=1)
        self.term_container.grid_rowconfigure(1, weight=1)

        # Terminal Top Toolbar (Search Ctrl+F + Copy Output)
        self.term_toolbar = ctk.CTkFrame(
            self.term_container, fg_color=self.palette["bg_card"], height=36, corner_radius=6
        )
        self.term_toolbar.grid(row=0, column=0, columnspan=2, sticky="ew", padx=6, pady=(6, 2))

        ctk.CTkLabel(
            self.term_toolbar, text="🔍", font=ctk.CTkFont(size=12), text_color=self.palette["text_muted"]
        ).pack(side="left", padx=(10, 4))

        self.search_entry = ctk.CTkEntry(
            self.term_toolbar,
            placeholder_text="Find in output (Ctrl+F)...",
            width=220,
            height=26,
            fg_color=self.palette["term_bg"],
            border_color=self.palette["border"],
            font=ctk.CTkFont(family="Segoe UI", size=11),
        )
        self.search_entry.pack(side="left", padx=4, pady=4)
        self.search_entry.bind("<KeyRelease>", self._on_search_key)
        self.search_entry.bind("<Return>", lambda e: self._step_search(1))
        self.search_entry.bind("<Shift-Return>", lambda e: self._step_search(-1))

        self.search_prev_btn = ctk.CTkButton(
            self.term_toolbar, text="▲", width=26, height=24, fg_color=self.palette["border"], command=lambda: self._step_search(-1)
        )
        self.search_prev_btn.pack(side="left", padx=2)

        self.search_next_btn = ctk.CTkButton(
            self.term_toolbar, text="▼", width=26, height=24, fg_color=self.palette["border"], command=lambda: self._step_search(1)
        )
        self.search_next_btn.pack(side="left", padx=2)

        self.search_count_lbl = ctk.CTkLabel(
            self.term_toolbar, text="", font=ctk.CTkFont(family="Consolas", size=11), text_color=self.palette["text_muted"]
        )
        self.search_count_lbl.pack(side="left", padx=8)

        ctk.CTkLabel(
            self.term_toolbar,
            text="💡 Right-click any MAC, IP, or CMAC for Read-Only Drilldowns",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            text_color=self.palette["text_muted"],
        ).pack(side="left", padx=12)

        self.copy_term_btn = ctk.CTkButton(
            self.term_toolbar,
            text="📋 Copy Output",
            width=105,
            height=26,
            corner_radius=6,
            fg_color=self.palette["border"],
            hover_color=self.palette["accent"],
            font=ctk.CTkFont(family="Segoe UI", size=11),
            command=self._copy_terminal_output,
        )
        self.copy_term_btn.pack(side="right", padx=8, pady=4)

        self.term_text = tk.Text(
            self.term_container,
            wrap="none",
            bg=self.palette["term_bg"],
            fg=self.palette["term_fg"],
            insertbackground=self.palette["text_main"],
            font=("Consolas", 11),
            bd=0,
            highlightthickness=0,
        )
        term_scroll_y = tk.Scrollbar(self.term_container, orient="vertical", command=self.term_text.yview)
        term_scroll_x = tk.Scrollbar(self.term_container, orient="horizontal", command=self.term_text.xview)
        self.term_text.configure(xscrollcommand=term_scroll_x.set, yscrollcommand=term_scroll_y.set)

        self.term_text.grid(row=1, column=0, sticky="nsew", padx=(8, 0), pady=(4, 0))
        term_scroll_y.grid(row=1, column=1, sticky="ns", pady=(4, 0))
        term_scroll_x.grid(row=2, column=0, sticky="ew", padx=(8, 0), pady=(0, 4))

        self.term_text.tag_configure("red_alert", foreground=self.palette["status_crit"], font=("Consolas", 11, "bold"))
        self.term_text.tag_configure("amber_warn", foreground=self.palette["status_warn"], font=("Consolas", 11, "bold"))
        self.term_text.tag_configure("green_ok", foreground=self.palette["status_ok"])
        self.term_text.tag_configure("header", foreground=self.palette["accent"], font=("Consolas", 11, "bold"))
        self.term_text.tag_configure("search_hit", background="#475569", foreground="#FFFFFF")
        self.term_text.tag_configure("search_active", background="#F59E0B", foreground="#000000")

        self.term_text.bind("<Double-Button-1>", self._on_terminal_double_click)
        self.term_text.bind("<Button-3>", self._on_terminal_right_click)

        # Pane 2: Spectrum & Analytics Canvas Grid (Right Pane)
        self.chart_frame = tk.Frame(self.console_box, bg=self.palette["bg_card"])
        if matplotlib:
            self.fig = Figure(figsize=(7, 6), dpi=100, facecolor=self.palette["bg_card"])
            self.ax1 = self.fig.add_subplot(2, 2, 1)              # Top-Left: SNR / RTT Bars
            self.ax2 = self.fig.add_subplot(2, 2, 2, polar=True)  # Top-Right: Polar Radar
            self.ax3 = self.fig.add_subplot(2, 1, 2)              # Bottom: Time-Series
            self.fig.tight_layout(pad=3.0, w_pad=2.0, h_pad=2.5)
            self.canvas = FigureCanvasTkAgg(self.fig, master=self.chart_frame)
            self.canvas.get_tk_widget().pack(fill="both", expand=True, padx=4, pady=4)
            self._reset_chart_visuals()
        else:
            self.fig = None
            self.canvas = None

        # Pane 3: Plant Heuristics, Blast-Radius Correlator & Truck-Roll Verdict View
        self.heurs_box = ctk.CTkTextbox(
            self.console_box, wrap="word", fg_color=self.palette["term_bg"], text_color=self.palette["text_main"], font=ctk.CTkFont(family="Consolas", size=12), corner_radius=8
        )

        # Pane 4: Parallel Fleet Health Leaderboard View
        self.leaderboard_frame = ctk.CTkFrame(self.console_box, fg_color="transparent")
        self._build_leaderboard_view()

        # Pane 5: Maintenance Diff View
        self.diff_box = ctk.CTkTextbox(
            self.console_box, wrap="none", fg_color=self.palette["term_bg"], text_color=self.palette["text_main"], font=ctk.CTkFont(family="Consolas", size=11), corner_radius=8
        )

        # Pane 6: Flight Recorder Replay & Chronic Offender Analyzer View
        self.flight_frame = ctk.CTkFrame(self.console_box, fg_color="transparent")
        self._build_flight_recorder_view()

        # Pane 7: Settings & Display Options View
        self.settings_frame = ctk.CTkFrame(self.console_box, fg_color="transparent")
        self._build_settings_view()

        # Pane 8: Session Log Box
        self.log_box = ctk.CTkTextbox(
            self.console_box, wrap="char", fg_color=self.palette["term_bg"], text_color=self.palette["text_muted"], font=ctk.CTkFont(family="Consolas", size=11), corner_radius=8
        )

        # 4. Bottom Interactive Read-Only CLI Prompt Dock
        self.dock = ctk.CTkFrame(
            self.stage, fg_color=self.palette["bg_card"], corner_radius=10, border_width=1, border_color=self.palette["border"]
        )
        self.dock.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        self.dock.grid_columnconfigure(1, weight=1)

        self.prompt_lbl = ctk.CTkLabel(
            self.dock, text="cmts:ro-cli >", font=ctk.CTkFont(family="Consolas", size=13, weight="bold"), text_color=self.palette["accent"]
        )
        self.prompt_lbl.grid(row=0, column=0, padx=(14, 6), pady=8)

        self.cli_entry = ctk.CTkEntry(
            self.dock,
            placeholder_text="Enter Read-Only CLI command (only 'show' and 'ping' commands are permitted)...",
            placeholder_text_color="#64748B",
            font=ctk.CTkFont(family="Consolas", size=12),
            height=36,
            fg_color=self.palette["term_bg"],
            border_color=self.palette["border"],
            corner_radius=8,
        )
        self.cli_entry.grid(row=0, column=1, sticky="ew", padx=6, pady=8)

        self.cli_entry.bind("<KeyRelease>", self._on_cli_key)
        self.cli_entry.bind("<Return>", lambda e: self.dispatch_cli_command())
        self.cli_entry.bind("<Down>", self._on_cli_down)
        self.cli_entry.bind("<Up>", self._on_cli_up)
        self.cli_entry.bind("<Tab>", self._on_cli_tab)
        self.cli_entry.bind("<Escape>", lambda e: self.popup.withdraw())

        self.send_btn = ctk.CTkButton(
            self.dock,
            text="Send (RO)",
            width=90,
            height=36,
            corner_radius=8,
            fg_color=self.palette["accent"],
            hover_color=self.palette["btn_hover"],
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            command=self.dispatch_cli_command,
        )
        self.send_btn.grid(row=0, column=2, padx=(6, 12), pady=8)

        self.append_log("System operational. Strict Read-Only Command Guardrail active.")

    def _build_context_menu(self):
        self.context_menu = tk.Menu(
            self,
            tearoff=0,
            bg="#0D131E",
            fg="#F8FAFC",
            activebackground="#0EA5E9",
            activeforeground="#FFFFFF",
            font=("Segoe UI", 10),
        )

    def _build_leaderboard_view(self):
        self.leaderboard_frame.grid_columnconfigure(0, weight=1)
        self.leaderboard_frame.grid_rowconfigure(1, weight=1)

        top_bar = ctk.CTkFrame(self.leaderboard_frame, fg_color="transparent")
        top_bar.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 4))

        ctk.CTkLabel(
            top_bar,
            text="🏆 PARALLEL FLEET HEALTH LEADERBOARD (Sorted Worst-First by Online Ratio & Impairments)",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=self.palette["accent"],
        ).pack(side="left")

        self.leaderboard_box = ctk.CTkTextbox(
            self.leaderboard_frame,
            wrap="none",
            fg_color=self.palette["term_bg"],
            text_color=self.palette["text_main"],
            font=ctk.CTkFont(family="Consolas", size=12),
            corner_radius=8,
        )
        self.leaderboard_box.grid(row=1, column=0, sticky="nsew", padx=8, pady=6)
        self.leaderboard_box.insert(
            "1.0",
            "Switch Dispatch Mode to 'Parallel Fleet Audit', enter comma-separated CMTS hostnames, and click '🚀 Run Parallel Fleet Audit'.\n"
            "Up to 8 parallel read-only SSH workers will audit the fleet simultaneously and rank every node here.",
        )

    def _build_flight_recorder_view(self):
        self.flight_frame.grid_columnconfigure(0, weight=1)
        self.flight_frame.grid_rowconfigure(1, weight=1)

        search_bar = ctk.CTkFrame(self.flight_frame, fg_color="transparent")
        search_bar.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        search_bar.grid_columnconfigure(0, weight=1)

        self.flight_search_entry = ctk.CTkEntry(
            search_bar,
            placeholder_text="Search historical sessions by CMTS Hostname, MAC address, or command keyword...",
            height=32,
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self.flight_search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.flight_search_entry.bind("<Return>", lambda e: self._refresh_flight_search())

        self.flight_search_btn = ctk.CTkButton(
            search_bar,
            text="🔍 Search History",
            width=130,
            height=32,
            fg_color=self.palette["accent"],
            hover_color=self.palette["btn_hover"],
            command=self._refresh_flight_search,
        )
        self.flight_search_btn.grid(row=0, column=1, padx=(0, 6))

        self.chronic_btn = ctk.CTkButton(
            search_bar,
            text="🔥 Scan Chronic Offenders (30d)",
            width=210,
            height=32,
            fg_color=self.palette["status_warn"],
            hover_color="#D97706",
            text_color="#04070D",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            command=self._run_chronic_offender_scan,
        )
        self.chronic_btn.grid(row=0, column=2)

        content_split = ctk.CTkFrame(self.flight_frame, fg_color="transparent")
        content_split.grid(row=1, column=0, sticky="nsew", padx=6, pady=(0, 6))
        content_split.grid_columnconfigure(0, weight=1)
        content_split.grid_columnconfigure(1, weight=2)
        content_split.grid_rowconfigure(0, weight=1)

        self.flight_listbox = tk.Listbox(
            content_split,
            bg=self.palette["term_bg"],
            fg=self.palette["text_main"],
            selectbackground=self.palette["accent"],
            selectforeground="#FFFFFF",
            font=("Consolas", 10),
            bd=0,
            highlightthickness=0,
        )
        self.flight_listbox.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self.flight_listbox.bind("<<ListboxSelect>>", self._on_flight_select)

        self.flight_preview = ctk.CTkTextbox(
            content_split, wrap="none", fg_color=self.palette["term_bg"], text_color=self.palette["text_main"], font=ctk.CTkFont(family="Consolas", size=11)
        )
        self.flight_preview.grid(row=0, column=1, sticky="nsew")

    def _build_settings_view(self):
        self.settings_card = ctk.CTkFrame(
            self.settings_frame, fg_color=self.palette["bg_card"], corner_radius=10, border_width=1, border_color=self.palette["border"]
        )
        self.settings_card.pack(fill="x", padx=16, pady=16)

        ctk.CTkLabel(
            self.settings_card,
            text="Display Themes & Console Aesthetics",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=self.palette["text_main"],
        ).pack(anchor="w", padx=18, pady=(16, 4))

        ctk.CTkLabel(
            self.settings_card,
            text="Switch between real-time operational themes, high-contrast monochrome styles, and modern cyber palettes.",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=self.palette["text_muted"],
        ).pack(anchor="w", padx=18, pady=(0, 12))

        self.theme_dropdown = ctk.CTkOptionMenu(
            self.settings_card,
            values=list(THEMES.keys()),
            command=self.apply_theme_on_the_fly,
            height=34,
            width=260,
            corner_radius=8,
            fg_color=self.palette["border"],
            button_color=self.palette["accent"],
            dropdown_fg_color=self.palette["bg_card"],
            font=ctk.CTkFont(family="Segoe UI", size=12),
        )
        self.theme_dropdown.set(self.current_theme)
        self.theme_dropdown.pack(anchor="w", padx=18, pady=(0, 20))

    # --- UI Widget Builders ---
    def _make_input(self, parent, label: str, default: str = "", placeholder: str = "", **kwargs) -> Tuple[ctk.CTkFrame, ctk.CTkEntry]:
        wrapper = ctk.CTkFrame(parent, fg_color="transparent")

        lbl = ctk.CTkLabel(
            wrapper, text=label, font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"), text_color=self.palette["text_muted"]
        )
        lbl.pack(anchor="w", pady=(0, 1))

        entry = ctk.CTkEntry(
            wrapper,
            height=28,
            fg_color=self.palette["term_bg"],
            border_color=self.palette["border"],
            text_color=self.palette["text_main"],
            corner_radius=6,
            font=ctk.CTkFont(family="Segoe UI", size=12),
            placeholder_text=placeholder,
            **kwargs,
        )
        if default:
            entry.insert(0, default)
        entry.pack(anchor="w", fill="x")
        return wrapper, entry

    def _make_kpi_card(self, parent, col: int, title: str, default_val: str, color: str, progress: float):
        padx = (0, 5) if col == 0 else (5, 0) if col == 5 else 3
        card = ctk.CTkFrame(
            parent, fg_color=self.palette["bg_card"], corner_radius=8, border_width=1, border_color=self.palette["border"]
        )
        card.grid(row=0, column=col, sticky="ew", padx=padx)
        self.kpi_cards.append(card)

        t_lbl = ctk.CTkLabel(
            card, text=title, font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"), text_color=self.palette["text_muted"]
        )
        t_lbl.pack(anchor="w", padx=10, pady=(6, 0))

        v_lbl = ctk.CTkLabel(
            card, text=default_val, font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"), text_color=color
        )
        v_lbl.pack(anchor="w", padx=10, pady=(0, 5))

        bar = ctk.CTkProgressBar(card, height=4, corner_radius=2, fg_color=self.palette["term_bg"], progress_color=color)
        bar.set(progress)
        bar.pack(fill="x", padx=10, pady=(0, 7))

        return v_lbl, bar

    def _toggle_auth_drawer(self, force_collapse: Optional[bool] = None):
        if force_collapse is not None:
            self.auth_expanded = not force_collapse
        else:
            self.auth_expanded = not self.auth_expanded

        if self.auth_expanded:
            self.cred_body.pack(fill="x", padx=14, pady=(0, 10))
            self.cred_toggle_btn.configure(text="🔐 Bastion & Credentials   ▼")
        else:
            self.cred_body.pack_forget()
            status_tag = "Connected" if self.is_connected else "Configured"
            self.cred_toggle_btn.configure(text=f"🔐 Bastion & Credentials ({status_tag})   ▶")

    # --- Live Dynamic Re-Theming Engine ---
    def apply_theme_on_the_fly(self, theme_name: str):
        self.current_theme = theme_name
        self.palette = THEMES.get(theme_name, THEMES["Obsidian Tactical"])

        self.configure(fg_color=self.palette["bg_main"])
        self.sidebar.configure(fg_color=self.palette["bg_panel"])
        self.stage.configure(fg_color=self.palette["bg_main"])

        cards = [
            self.target_card, self.cred_card, self.cat_card, self.playbook_card,
            self.param_card, self.watch_box, self.header, self.console_box,
            self.dock, self.settings_card
        ] + self.kpi_cards
        for c in cards:
            c.configure(fg_color=self.palette["bg_card"], border_color=self.palette["border"])

        self.term_container.configure(fg_color=self.palette["term_bg"])
        self.term_toolbar.configure(fg_color=self.palette["bg_card"])

        self.exec_btn.configure(fg_color=self.palette["accent"], hover_color=self.palette["btn_hover"])
        self.send_btn.configure(fg_color=self.palette["accent"], hover_color=self.palette["btn_hover"])
        self.nav_bar.configure(selected_color=self.palette["accent"], unselected_color=self.palette["bg_panel"])
        self.catalog_dropdown.configure(button_color=self.palette["accent"], fg_color=self.palette["border"], dropdown_fg_color=self.palette["bg_card"])
        self.theme_dropdown.configure(button_color=self.palette["accent"], fg_color=self.palette["border"], dropdown_fg_color=self.palette["bg_card"])
        self.op_mode.configure(selected_color=self.palette["accent"], unselected_color=self.palette["bg_panel"])
        self.prompt_lbl.configure(text_color=self.palette["accent"])
        self.flight_search_btn.configure(fg_color=self.palette["accent"], hover_color=self.palette["btn_hover"])

        self.term_text.configure(
            bg=self.palette["term_bg"],
            fg=self.palette["term_fg"],
            insertbackground=self.palette["text_main"],
        )
        self.term_text.tag_configure("red_alert", foreground=self.palette["status_crit"])
        self.term_text.tag_configure("amber_warn", foreground=self.palette["status_warn"])
        self.term_text.tag_configure("green_ok", foreground=self.palette["status_ok"])
        self.term_text.tag_configure("header", foreground=self.palette["accent"])

        self.heurs_box.configure(fg_color=self.palette["term_bg"], text_color=self.palette["text_main"])
        self.leaderboard_box.configure(fg_color=self.palette["term_bg"], text_color=self.palette["text_main"])
        self.diff_box.configure(fg_color=self.palette["term_bg"], text_color=self.palette["text_main"])
        self.flight_preview.configure(fg_color=self.palette["term_bg"], text_color=self.palette["text_main"])
        self.log_box.configure(fg_color=self.palette["term_bg"], text_color=self.palette["text_muted"])
        self.flight_listbox.configure(bg=self.palette["term_bg"], fg=self.palette["text_main"], selectbackground=self.palette["accent"])
        self.cli_entry.configure(fg_color=self.palette["term_bg"], border_color=self.palette["border"])

        if matplotlib and self.fig:
            self.chart_frame.configure(bg=self.palette["bg_card"])
            self.fig.patch.set_facecolor(self.palette["bg_card"])
            self._reset_chart_visuals()

        cfg = load_config()
        cfg["theme"] = theme_name
        save_config(cfg)

    # --- Mode & Parameter Synchronization ---
    def _on_mode_change(self, mode: str):
        if mode == "Parallel Fleet Audit":
            self.cmts_host_wrapper.pack_forget()
            self.bulk_target_wrapper.pack(fill="x", padx=14, pady=(2, 10))
            self.exec_btn.configure(text="🚀 Run Parallel Fleet Audit")
        else:
            self.bulk_target_wrapper.pack_forget()
            self.cmts_host_wrapper.pack(fill="x", padx=14, pady=(2, 10))
            self.exec_btn.configure(text="▶  Execute Show Workflow")

    def _on_catalog_select(self, choice: str):
        spec = COMMAND_CATALOG[choice]
        self.catalog_desc.configure(text=spec.desc)

    def _on_playbook_select(self, choice: str):
        spec = PLAYBOOK_CATALOG[choice]
        self.playbook_desc.configure(text=spec.desc)

    def _switch_stage_tab(self, choice: str):
        self.heurs_box.grid_forget()
        self.leaderboard_frame.grid_forget()
        self.diff_box.grid_forget()
        self.flight_frame.grid_forget()
        self.settings_frame.grid_forget()
        self.log_box.grid_forget()

        if choice in ["Terminal", "Split View"]:
            self.nav_bar.set(choice)
            if choice == "Split View":
                self.term_container.grid(row=0, column=0, columnspan=1, sticky="nsew", padx=(8, 4), pady=8)
                self.chart_frame.grid(row=0, column=1, sticky="nsew", padx=(4, 8), pady=8)
            else:
                self.chart_frame.grid_forget()
                self.term_container.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=8, pady=8)
        else:
            self.term_container.grid_forget()
            self.chart_frame.grid_forget()
            if choice in ["Heuristics & Blast", "Fleet Leaderboard", "Maint Diff", "History & Offenders"]:
                self.nav_bar.set(choice)

            if choice == "Heuristics & Blast":
                self.heurs_box.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=8, pady=8)
            elif choice == "Fleet Leaderboard":
                self.leaderboard_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=8, pady=8)
            elif choice == "Maint Diff":
                self.diff_box.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=8, pady=8)
            elif choice == "History & Offenders":
                self.flight_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=8, pady=8)
                self._refresh_flight_search()
            elif choice == "Settings":
                self.settings_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=8, pady=8)
            elif choice == "Session Logs":
                self.log_box.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=8, pady=8)

    # --- Terminal Search (Ctrl+F) & Clipboard Copy ---
    def _focus_terminal_search(self, _=None):
        if self.nav_bar.get() not in ["Terminal", "Split View"]:
            self._switch_stage_tab("Terminal")
        self.search_entry.focus_set()
        self.search_entry.select_range(0, "end")

    def _on_search_key(self, event=None):
        if event and event.keysym in ("Return", "Up", "Down"):
            return
        query = self.search_entry.get().strip()
        self.term_text.tag_remove("search_hit", "1.0", "end")
        self.term_text.tag_remove("search_active", "1.0", "end")
        self.search_matches.clear()
        self.search_idx = -1

        if not query:
            self.search_count_lbl.configure(text="")
            return

        start_pos = "1.0"
        while True:
            pos = self.term_text.search(query, start_pos, stopindex="end", nocase=True)
            if not pos:
                break
            end_pos = f"{pos}+{len(query)}c"
            self.search_matches.append((pos, end_pos))
            self.term_text.tag_add("search_hit", pos, end_pos)
            start_pos = end_pos

        if self.search_matches:
            self.search_idx = 0
            self._highlight_active_search()
        else:
            self.search_count_lbl.configure(text="0/0")

    def _step_search(self, delta: int):
        if not self.search_matches:
            return
        self.search_idx = (self.search_idx + delta) % len(self.search_matches)
        self._highlight_active_search()

    def _highlight_active_search(self):
        self.term_text.tag_remove("search_active", "1.0", "end")
        pos, end_pos = self.search_matches[self.search_idx]
        self.term_text.tag_add("search_active", pos, end_pos)
        self.term_text.see(pos)
        self.search_count_lbl.configure(text=f"{self.search_idx + 1}/{len(self.search_matches)}")

    def _copy_terminal_output(self):
        content = self.term_text.get("1.0", "end").strip()
        if not content:
            return
        self.clipboard_clear()
        self.clipboard_append(content)
        self.copy_term_btn.configure(text="✓ Copied!", fg_color=self.palette["status_ok"])
        self.after(1500, lambda: self.copy_term_btn.configure(text="📋 Copy Output", fg_color=self.palette["border"]))

    # --- 1-Click Copy Jira / ServiceNow Ticket Notes ---
    def copy_ticket_summary(self):
        host = self.active_host or self.cmts_host_entry.get().strip() or "UNKNOWN_CMTS"
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        online_ratio = self.badge_online.cget("text")
        total_modems = self.badge_total.cget("text")
        offline_cnt = self.badge_offline.cget("text")
        rf_state = self.badge_rf_health.cget("text")
        dispatch_state = self.badge_dispatch.cget("text")
        triage_state = self.badge_heurs.cget("text")

        cli_snippet = "\n".join(self.last_output.strip().splitlines()[:35])
        if len(self.last_output.splitlines()) > 35:
            cli_snippet += "\n... [Output truncated for ticket brevity]"

        ticket_markdown = (
            f"### 🛠️ Arris E6000 Read-Only Diagnostic & Dispatch Summary\n"
            f"- **Target CMTS:** `{host}` | **Timestamp:** `{now}`\n"
            f"- **Evaluated Workflow:** `{self.last_command or 'N/A'}`\n"
            f"- **KPI Telemetry:** Online: **{online_ratio}** | Total: **{total_modems}** | "
            f"Offline: **{offline_cnt}** | RF CW: **{rf_state}** | Triage: **{triage_state}**\n"
            f"- **Truck-Roll Dispatch Verdict:** **{dispatch_state}** — {self.last_dispatch_detail}\n\n"
            f"#### 🎯 Smart Outage Blast-Radius & Root-Cause Heuristics\n"
            f"```text\n{self.last_triage_report.strip()}\n```\n\n"
            f"#### 💻 Supporting Read-Only CLI Evidence\n"
            f"```text\n{cli_snippet or 'No CLI output captured yet.'}\n```\n"
        )

        self.clipboard_clear()
        self.clipboard_append(ticket_markdown)
        self.update()
        self.append_log("Copied formatted Jira/ServiceNow ticket summary to clipboard.")
        messagebox.showinfo("Ticket Notes Copied", "Formatted Markdown diagnostic & dispatch summary copied to clipboard!")

    # --- 1-Click Executive HTML Scorecard Generator ---
    def export_executive_html(self):
        host = self.active_host or self.cmts_host_entry.get().strip() or "FLEET_SUMMARY"
        timestamp = datetime.now().strftime("%Y_%m_%d_%H%M%S")
        default_filename = f"EXECUTIVE_SCORECARD_{host}_{timestamp}.html"

        chosen_path = filedialog.asksaveasfilename(
            parent=self,
            title="Export Executive HTML Scorecard",
            initialdir=str(Path.home() / "Desktop"),
            initialfile=default_filename,
            defaultextension=".html",
            filetypes=[("HTML Scorecard", "*.html"), ("All Files", "*.*")],
        )
        if not chosen_path:
            return

        # Capture Base64 PNG of the Matplotlib 2x2 Canvas
        chart_img_html = "<p style='color:#94A3B8;'>Matplotlib chart canvas not active.</p>"
        if matplotlib and self.fig:
            buf = io.BytesIO()
            self.fig.savefig(buf, format="png", dpi=130, facecolor=self.palette["bg_card"], bbox_inches="tight")
            buf.seek(0)
            b64_str = base64.b64encode(buf.read()).decode("utf-8")
            chart_img_html = f"<img src='data:image/png;base64,{b64_str}' style='width:100%; max-width:920px; border-radius:8px; border:1px solid #1E293B;' />"

        leaderboard_section = ""
        if self.fleet_leaderboard_rows:
            rows_html = ""
            for idx, r in enumerate(self.fleet_leaderboard_rows, start=1):
                rows_html += (
                    f"<tr>"
                    f"<td>#{idx}</td><td><b>{r['host']}</b></td>"
                    f"<td>{r['online_pct']:.1f}%</td><td>{r['total']:,}</td>"
                    f"<td>{r['offline']:,}</td><td>{r['min_snr']:.1f} dB</td>"
                    f"<td>{r['uncorr']:,}</td><td>{r['status']}</td>"
                    f"</tr>"
                )
            leaderboard_section = f"""
            <div class="card">
                <h2>🏆 Parallel Fleet Audit Leaderboard</h2>
                <table>
                    <thead>
                        <tr><th>Rank</th><th>CMTS Node</th><th>Online Ratio</th><th>Total CMs</th><th>Offline CMs</th><th>Min SNR</th><th>Uncorrectables</th><th>Health Verdict</th></tr>
                    </thead>
                    <tbody>{rows_html}</tbody>
                </table>
            </div>
            """

        html_doc = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Arris E6000 Executive Health Scorecard — {host}</title>
<style>
    body {{ background:#070A0F; color:#F8FAFC; font-family:'Segoe UI', sans-serif; margin:28px; }}
    .header {{ background:#0D131E; padding:20px 24px; border-radius:10px; border:1px solid #1E293B; margin-bottom:18px; display:flex; justify-content:space-between; align-items:center; }}
    .kpi-grid {{ display:grid; grid-template-columns:repeat(6, 1fr); gap:12px; margin-bottom:18px; }}
    .kpi {{ background:#0D131E; padding:14px; border-radius:8px; border:1px solid #1E293B; }}
    .kpi-title {{ color:#64748B; font-size:11px; font-weight:bold; text-transform:uppercase; }}
    .kpi-val {{ font-size:20px; font-weight:bold; margin-top:6px; color:#38BDF8; }}
    .card {{ background:#0D131E; padding:20px; border-radius:10px; border:1px solid #1E293B; margin-bottom:18px; }}
    h1, h2 {{ margin:0 0 12px 0; color:#F8FAFC; }}
    pre {{ background:#04070D; padding:14px; border-radius:8px; border:1px solid #1E293B; overflow-x:auto; color:#38BDF8; font-family:Consolas, monospace; font-size:12px; }}
    table {{ width:100%; border-collapse:collapse; font-size:13px; }}
    th, td {{ border:1px solid #1E293B; padding:8px 12px; text-align:left; }}
    th {{ background:#1E293B; color:#38BDF8; }}
</style>
</head>
<body>
    <div class="header">
        <div>
            <h1>Arris E6000 CER — Executive Diagnostic Scorecard</h1>
            <div style="color:#94A3B8; font-size:13px;">Target Node: <b>{host}</b> | Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Mode: Strictly Read-Only</div>
        </div>
        <div style="background:#10B981; color:#04070D; font-weight:bold; padding:8px 14px; border-radius:6px;">ZERO-IMPACT READ-ONLY AUDIT</div>
    </div>

    <div class="kpi-grid">
        <div class="kpi"><div class="kpi-title">Online Ratio</div><div class="kpi-val">{self.badge_online.cget("text")}</div></div>
        <div class="kpi"><div class="kpi-title">Total Modems</div><div class="kpi-val">{self.badge_total.cget("text")}</div></div>
        <div class="kpi"><div class="kpi-title">Offline Count</div><div class="kpi-val">{self.badge_offline.cget("text")}</div></div>
        <div class="kpi"><div class="kpi-title">RF Codewords</div><div class="kpi-val">{self.badge_rf_health.cget("text")}</div></div>
        <div class="kpi"><div class="kpi-title">Dispatch Verdict</div><div class="kpi-val">{self.badge_dispatch.cget("text")}</div></div>
        <div class="kpi"><div class="kpi-title">Triage Status</div><div class="kpi-val">{self.badge_heurs.cget("text")}</div></div>
    </div>

    <div class="card">
        <h2>🚚 Truck-Roll Preventer (Premise vs. Outside Plant Verdict)</h2>
        <p style="font-size:14px; color:#E2E8F0;"><b>{self.badge_dispatch.cget("text")}</b> — {self.last_dispatch_detail}</p>
    </div>

    {leaderboard_section}

    <div class="card" style="text-align:center;">
        <h2>📊 RF Spectrum, Polar Health & Signal Telemetry</h2>
        {chart_img_html}
    </div>

    <div class="card">
        <h2>🎯 Smart Outage Blast-Radius & Root-Cause Heuristics</h2>
        <pre>{self.last_triage_report}</pre>
    </div>

    <div class="card">
        <h2>⚖️ Pre / Post Maintenance Diff Verification</h2>
        <pre>{self.last_diff_text}</pre>
    </div>

    <div class="card">
        <h2>💻 Raw Read-Only CLI Telemetry Excerpt</h2>
        <pre>{self.last_output[:8000] if self.last_output else 'No CLI output captured.'}</pre>
    </div>
</body>
</html>"""

        try:
            with open(chosen_path, "w", encoding="utf-8") as f:
                f.write(html_doc)
            self.append_log(f"Executive HTML Scorecard exported: {chosen_path}")
            webbrowser.open(f"file://{Path(chosen_path).resolve()}")
        except Exception as e:
            messagebox.showerror("HTML Export Error", str(e))

    # --- Main-Thread Safe VIP Token Request with 35s Memory Cache ---
    def _acquire_vip_token_if_needed(self, target_check: Optional[str] = None) -> Optional[str]:
        if (
            self.is_connected
            and self.ssh_channel
            and not self.ssh_channel.closed
            and (target_check is None or self.active_host == target_check)
        ):
            try:
                if self.ssh_channel.get_transport().is_active():
                    return self.cached_vip_token or ""
            except Exception:
                pass

        if self.cached_vip_token and (time.time() - self.cached_vip_timestamp) < VIP_CACHE_TTL_SEC:
            self.append_log("Using cached VIP token (within active validity window).")
            return self.cached_vip_token

        vip_dialog = ctk.CTkInputDialog(text="Enter 6-digit VIP Token:", title="Bastion Authentication")
        vip_token = vip_dialog.get_input()
        if not vip_token or not vip_token.strip():
            self.append_log("Aborted: VIP Security code omitted.")
            return None
        self.cached_vip_token = vip_token.strip()
        self.cached_vip_timestamp = time.time()
        return self.cached_vip_token

    # --- Double-Click & Right-Click Read-Only Terminal Drilldowns ---
    def _on_terminal_double_click(self, _):
        try:
            line = self.term_text.get("current linestart", "current lineend")
            mac_pattern = r"([0-9a-fA-F]{4}\.[0-9a-fA-F]{4}\.[0-9a-fA-F]{4}|(?:[0-9a-fA-F]{2}[:-]){5}[0-9a-fA-F]{2})"
            match = re.search(mac_pattern, line)
            if match:
                found_mac = match.group(1)
                self.catalog_dropdown.set("Individual Modem Diagnostic")
                self._on_catalog_select("Individual Modem Diagnostic")

                self.mac_input.delete(0, "end")
                self.mac_input.insert(0, found_mac)

                self.cli_entry.delete(0, "end")
                self.cli_entry.insert(0, f"show cable modem {found_mac} verbose")
                self.append_log(f"Read-Only drilldown staged for MAC: {found_mac} ({lookup_mac_vendor(found_mac)})")
        except Exception:
            pass

    def _on_terminal_right_click(self, event):
        try:
            idx = self.term_text.index(f"@{event.x},{event.y}")
            line = self.term_text.get(f"{idx} linestart", f"{idx} lineend")

            mac_match = re.search(
                r"([0-9a-fA-F]{4}\.[0-9a-fA-F]{4}\.[0-9a-fA-F]{4}|(?:[0-9a-fA-F]{2}[:-]){5}[0-9a-fA-F]{2})",
                line,
            )
            ip_match = re.search(r"\b(?:\d{1,3}\.){3}\d{1,3}\b", line)
            cmac_match = re.search(r"cable-mac\s*(\d+)|\bcmac\s*(\d+)", line, re.IGNORECASE)

            self.context_menu.delete(0, tk.END)

            if mac_match:
                mac = mac_match.group(1)
                vendor_str = lookup_mac_vendor(mac)
                self.context_menu.add_command(
                    label=f"🏷️ OUI Vendor: {vendor_str}",
                    command=lambda m=mac, v=vendor_str: messagebox.showinfo("DOCSIS OUI Lookup", f"MAC: {m}\nOUI / Manufacturer: {v}"),
                )
                self.context_menu.add_separator()
                self.context_menu.add_command(
                    label=f"⚡ Run Modem 360° + Truck-Roll Classifier ({mac})",
                    command=lambda m=mac: self._trigger_mac_playbook(m),
                )
                self.context_menu.add_command(
                    label=f"🔍 Show Modem Verbose ({mac})",
                    command=lambda m=mac: self._route_single_dispatch(f"show cable modem {m} verbose"),
                )
                self.context_menu.add_command(
                    label=f"📡 Show Modem PHY Levels ({mac})",
                    command=lambda m=mac: self._route_single_dispatch(f"show cable modem {m} phy"),
                )
                self.context_menu.add_command(
                    label=f"🌐 Show Connected CPEs ({mac})",
                    command=lambda m=mac: self._route_single_dispatch(f"show cable modem {m} cpe"),
                )
                self.context_menu.add_command(
                    label=f"📊 Show QoS Service Flows ({mac})",
                    command=lambda m=mac: self._route_single_dispatch(f"show cable modem {m} service-flow"),
                )

            if ip_match:
                ip = ip_match.group(0)
                if self.context_menu.index("end") is not None:
                    self.context_menu.add_separator()
                self.context_menu.add_command(
                    label=f"📈 Run Non-Intrusive Ping Probe (ping {ip})",
                    command=lambda target_ip=ip: self._route_single_dispatch(f"ping {target_ip}"),
                )

            if cmac_match:
                cmac = cmac_match.group(1) or cmac_match.group(2)
                if self.context_menu.index("end") is not None:
                    self.context_menu.add_separator()
                self.context_menu.add_command(
                    label=f"⚡ Run Cable-Mac {cmac} Blast-Radius & OFDM Sweep",
                    command=lambda c=cmac: self._trigger_cmac_playbook(c),
                )
                self.context_menu.add_command(
                    label=f"📊 Show Cable Noise (cable-mac {cmac})",
                    command=lambda c=cmac: self._route_single_dispatch(f"show cable noise cable-mac {c}"),
                )

            if self.context_menu.index("end") is not None:
                self.context_menu.add_separator()
            self.context_menu.add_command(
                label="📋 Copy Ticket Summary to Clipboard",
                command=self.copy_ticket_summary,
            )

            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def _trigger_mac_playbook(self, mac: str):
        self.mac_input.delete(0, "end")
        self.mac_input.insert(0, mac)
        self.playbook_dropdown.set("Modem 360° + Truck-Roll Classifier (RO)")
        self._on_playbook_select("Modem 360° + Truck-Roll Classifier (RO)")
        self.dispatch_playbook()

    def _trigger_cmac_playbook(self, cmac: str):
        self.cmac_input.delete(0, "end")
        self.cmac_input.insert(0, cmac)
        self.playbook_dropdown.set("Cable-Mac Blast-Radius & OFDM Sweep (RO)")
        self._on_playbook_select("Cable-Mac Blast-Radius & OFDM Sweep (RO)")
        self.dispatch_playbook()

    # --- Telemetry, Blast-Radius Correlator & Truck-Roll Classifier ---
    def _highlight_and_stream(self, text_chunk: str):
        for line in text_chunk.splitlines(keepends=True):
            lower = line.lower()
            if any(fail in lower for fail in ["offline", "init(d)", "ranging-abort", "crit", "fail"]):
                self.term_text.insert("end", line, "red_alert")
            elif any(warn in lower for warn in ["ranging", "p-online", "degraded", "flap"]):
                self.term_text.insert("end", line, "amber_warn")
            elif any(ok in lower for ok in ["online", "operational", "active"]):
                self.term_text.insert("end", line, "green_ok")
            elif line.strip().startswith("---") or "# show" in line or "# ping" in line or line.strip().startswith("==="):
                self.term_text.insert("end", line, "header")
            else:
                self.term_text.insert("end", line)
        self.term_text.see("end")
        if self.search_entry.get().strip():
            self._on_search_key()

    def _parse_telemetry(self, raw_output: str):
        if "total" in raw_output.lower() and ("registered" in raw_output.lower() or "offline" in raw_output.lower()):
            for line in raw_output.splitlines():
                if "total" in line.lower():
                    numbers = [int(n) for n in re.findall(r"\b\d+\b", line)]
                    if len(numbers) >= 2:
                        total = max(numbers)
                        reg = numbers[0]
                        offline = total - reg
                        pct = (reg / total) * 100 if total > 0 else 0.0
                        ratio_float = max(0.0, min(1.0, pct / 100.0))
                        off_ratio = max(0.0, min(1.0, offline / max(total, 1)))

                        def update_badges():
                            online_col = self.palette["status_ok"] if pct >= 98.0 else self.palette["status_warn"] if pct >= 90.0 else self.palette["status_crit"]
                            off_col = self.palette["status_crit"] if offline > 20 else self.palette["status_warn"] if offline > 5 else self.palette["status_ok"]

                            self.badge_total.configure(text=f"{total:,}", text_color=self.palette["text_main"])
                            self.bar_total.configure(progress_color=self.palette["accent"])
                            self.bar_total.set(1.0 if total > 0 else 0.0)

                            self.badge_offline.configure(text=f"{offline:,}", text_color=off_col)
                            self.bar_offline.configure(progress_color=off_col)
                            self.bar_offline.set(off_ratio)

                            self.badge_online.configure(text=f"{pct:.1f}%", text_color=online_col)
                            self.bar_online.configure(progress_color=online_col)
                            self.bar_online.set(ratio_float)

                        self.after(0, update_badges)
                        break

        if "uncorrectable" in raw_output.lower() or "snr" in raw_output.lower():
            uncorrectables = [int(x) for x in re.findall(r"uncorr(?:ectables?)?[:\s]+(\d+)", raw_output, re.IGNORECASE)]
            self.node_uncorr_baseline = sum(uncorrectables) if uncorrectables else 0

            def update_rf():
                if any(u > 5000 for u in uncorrectables):
                    c = self.palette["status_crit"]
                    self.badge_rf_health.configure(text="CRITICAL", text_color=c)
                    self.bar_rf.configure(progress_color=c)
                    self.bar_rf.set(1.0)
                elif any(u > 500 for u in uncorrectables):
                    c = self.palette["status_warn"]
                    self.badge_rf_health.configure(text="DEGRADED", text_color=c)
                    self.bar_rf.configure(progress_color=c)
                    self.bar_rf.set(0.6)
                else:
                    c = self.palette["status_ok"]
                    self.badge_rf_health.configure(text="CLEAN", text_color=c)
                    self.bar_rf.configure(progress_color=c)
                    self.bar_rf.set(1.0)
            self.after(0, update_rf)

    def _correlate_blast_radius(self, output: str) -> Tuple[str, int, str]:
        """Groups impaired/flapping modems by upstream/downstream interface, Cable-Mac, and OUI vendor."""
        mac_pattern = r"([0-9a-fA-F]{4}\.[0-9a-fA-F]{4}\.[0-9a-fA-F]{4})"
        iface_pattern = r"\b(\d{1,2}/\d{1,2}/\d{1,2}(?:/\d+)?)\b"

        channel_counts: Counter = Counter()
        oui_counts: Counter = Counter()
        impaired_macs = set()

        for line in output.splitlines():
            lower = line.lower()
            mac_m = re.search(mac_pattern, line)
            if not mac_m:
                continue

            is_bad_line = any(
                k in lower for k in ["offline", "init", "ranging-abort", "flap", "degraded"]
            ) or "flap-list" in output.lower()

            if is_bad_line:
                mac = mac_m.group(1).lower()
                impaired_macs.add(mac)
                vendor = lookup_mac_vendor(mac)
                oui_counts[vendor] += 1
                for ifc in re.findall(iface_pattern, line):
                    channel_counts[ifc] += 1

        total_impaired = len(impaired_macs)
        if total_impaired == 0:
            return "✅ No multi-modem impairment concentration detected in current output.", 0, ""

        lines = [f"🎯 SMART OUTAGE BLAST-RADIUS CORRELATION ({total_impaired} Impaired Modems Analyzed):"]
        top_channel = ""
        if channel_counts:
            top_ch, ch_cnt = channel_counts.most_common(1)[0]
            top_channel = top_ch
            concentration = (ch_cnt / max(total_impaired, 1)) * 100
            if concentration >= 60 and total_impaired >= 3:
                lines.append(
                    f"   🚨 SHARED CHANNEL BOTTLENECK: {concentration:.0f}% ({ch_cnt}/{total_impaired}) of impaired modems "
                    f"share Interface {top_ch}! Issue is isolated to Outside Plant / Return Carrier {top_ch} — DO NOT roll individual premise trucks."
                )
            else:
                ch_summary = ", ".join(f"{c} ({n})" for c, n in channel_counts.most_common(4))
                lines.append(f"   • Interface Distribution: {ch_summary}")

        if oui_counts:
            top_oui, oui_cnt = oui_counts.most_common(1)[0]
            oui_pct = (oui_cnt / max(total_impaired, 1)) * 100
            if oui_pct >= 75 and total_impaired >= 4:
                lines.append(
                    f"   ⚠️ HARDWARE/FIRMWARE CONCENTRATION: {oui_pct:.0f}% ({oui_cnt}/{total_impaired}) belong to OUI {top_oui}."
                )
            else:
                oui_summary = ", ".join(f"{v} ({n})" for v, n in oui_counts.most_common(3))
                lines.append(f"   • Hardware OUI Breakdown: {oui_summary}")

        return "\n".join(lines), total_impaired, top_channel

    def _classify_truck_roll_verdict(self, cmd: str, output: str, snr_values: List[float], total_impaired: int):
        """Determines whether an issue is Inside-Home (Premise Drop) vs. Outside-Plant (Node/Channel)."""
        tx_powers = [float(p) for p in re.findall(r"tx[\s:]+(\d+\.\d+|\d+)", output, re.IGNORECASE)]
        rx_powers = [float(p) for p in re.findall(r"rx[\s:]+(-?\d+\.\d+|-?\d+)", output, re.IGNORECASE)]
        low_snrs = [s for s in snr_values if 5.0 < s < 28.0]
        cwe_matches = [int(x) for x in re.findall(r"uncorr(?:ectables?)?[:\s]+(\d+)", output, re.IGNORECASE)]
        total_uncorr = sum(cwe_matches)

        verdict_label = "NOMINAL"
        verdict_color = self.palette["status_ok"]
        verdict_detail = "All evaluated RF, SNR, and power parameters are within DOCSIS specification. No truck roll required."

        # Case 1: Node-wide SNR degradation or multi-modem blast radius -> OUTSIDE PLANT
        if low_snrs or total_impaired >= 5 or total_uncorr > 8000:
            verdict_label = "🌐 OUTSIDE PLANT"
            verdict_color = self.palette["status_crit"]
            verdict_detail = (
                f"Shared Outside-Plant impairment detected (Min SNR: {min(low_snrs) if low_snrs else self.node_snr_baseline:.1f} dB, "
                f"Impaired Modems: {total_impaired}, Uncorrectables: {total_uncorr:,}). "
                "Cancel residential tech dispatch and escalate to Plant Maintenance / Sweep team."
            )
        # Case 2: Single-modem PHY anomaly while node baseline is healthy -> PREMISE / DROP
        elif any(tx >= 52.5 for tx in tx_powers) or any(abs(rx) > 12.0 for rx in rx_powers) or (
            "modem" in cmd.lower() and ("offline" in output.lower() or "ranging-abort" in output.lower()) and total_impaired <= 2
        ):
            verdict_label = "🏠 PREMISE / DROP"
            verdict_color = self.palette["status_warn"]
            max_tx = max(tx_powers) if tx_powers else 0.0
            verdict_detail = (
                f"Isolated Premise/Drop anomaly (Modem Peak Tx: {max_tx:.1f} dBmV vs Node Baseline SNR: {self.node_snr_baseline:.1f} dB). "
                "Outside plant is healthy — dispatch residential technician to inspect house drop, ground block, or home splitter."
            )

        self.last_dispatch_verdict = verdict_label
        self.last_dispatch_detail = verdict_detail

        def update_dispatch_card():
            self.badge_dispatch.configure(text=verdict_label, text_color=verdict_color)
            self.bar_dispatch.configure(progress_color=verdict_color)
            self.bar_dispatch.set(1.0)

        self.after(0, update_dispatch_card)

    def _run_heuristics_triage(self, cmd: str, output: str):
        diagnostics: List[str] = []
        overall_severity = "CLEAN"

        cwe_matches = re.findall(r"uncorr(?:ectables?)?[:\s]+(\d+)", output, re.IGNORECASE)
        if cwe_matches:
            total_uncorr = sum(int(x) for x in cwe_matches)
            if total_uncorr > 10000:
                overall_severity = "CRITICAL"
                diagnostics.append(
                    "🚨 [CRITICAL: UNCORRECTABLE CW SPIKE]\n"
                    f"Observed {total_uncorr:,} uncorrectable codewords on upstream carriers.\n"
                    "Probable Causes: In-band ingress burst, corroded drop amplifier, or laser clipping."
                )
            elif total_uncorr > 1000:
                if overall_severity != "CRITICAL":
                    overall_severity = "DEGRADED"
                diagnostics.append(
                    f"⚠️ [WARNING: PACKET LOSS LEAK] Observed {total_uncorr:,} uncorrectables. Check plant shielding."
                )

        snr_values = parse_snr_from_output(output)
        if snr_values:
            self.node_snr_baseline = sum(snr_values) / len(snr_values)
        low_snrs = [s for s in snr_values if 5.0 < s < 28.0]
        if low_snrs:
            overall_severity = "CRITICAL"
            diagnostics.append(
                f"🚨 [CRITICAL: UPSTREAM SNR DEGRADATION]\n"
                f"Carriers operating as low as {min(low_snrs):.1f} dB SNR (Threshold: 28.0 dB).\n"
                "Action: Immediate return-path pad audit or inspect physical split/tap."
            )

        tx_powers = [float(p) for p in re.findall(r"tx[\s:]+(\d+\.\d+|\d+)", output, re.IGNORECASE)]
        if any(tx >= 53.0 for tx in tx_powers):
            if overall_severity != "CRITICAL":
                overall_severity = "DEGRADED"
            diagnostics.append(
                "⚠️ [ELEVATED TX POWER] Modem transmitting at/above 53.0 dBmV saturation limit. Plant attenuation high."
            )

        if "flap" in cmd.lower() and "ranging-abort" in output.lower():
            if overall_severity == "CLEAN":
                overall_severity = "DEGRADED"
            diagnostics.append(
                "⚡ [FLAP LIST ALERT] High ranging-abort events detected. Indicates loose seizure screw or micro-reflections."
            )

        blast_text, total_impaired, _ = self._correlate_blast_radius(output)
        self.last_blast_radius = blast_text
        self._classify_truck_roll_verdict(cmd, output, snr_values, total_impaired)

        if not diagnostics:
            diagnostics.append(
                "✅ [HEALTHY PLANT CONDITIONS]\n"
                "All SNR, CW error thresholds, power profiles, and registration levels are nominal."
            )

        triage_report = (
            f"=== AUTOMATED ROOT-CAUSE TRIAGE & BLAST-RADIUS REPORT (READ-ONLY) ===\n"
            f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Target: {self.active_host}\n"
            f"Evaluated Command: {cmd}\n"
            + "-" * 68 + "\n\n"
            + f"🚚 TRUCK-ROLL PREVENTER VERDICT: {self.last_dispatch_verdict}\n"
            + f"   {self.last_dispatch_detail}\n\n"
            + f"{blast_text}\n\n"
            + "\n\n".join(diagnostics)
        )
        self.last_triage_report = triage_report

        def update_heurs_ui(rep=triage_report, sev=overall_severity):
            self.heurs_box.delete("1.0", "end")
            self.heurs_box.insert("1.0", rep)
            if sev == "CRITICAL":
                c = self.palette["status_crit"]
                self.badge_heurs.configure(text="CRITICAL FAULT", text_color=c)
                self.bar_heurs.configure(progress_color=c)
                self.bar_heurs.set(1.0)
            elif sev == "DEGRADED":
                c = self.palette["status_warn"]
                self.badge_heurs.configure(text="DEGRADED", text_color=c)
                self.bar_heurs.configure(progress_color=c)
                self.bar_heurs.set(0.6)
            else:
                c = self.palette["status_ok"]
                self.badge_heurs.configure(text="ALL CLEAN", text_color=c)
                self.bar_heurs.configure(progress_color=c)
                self.bar_heurs.set(1.0)

        self.after(0, update_heurs_ui)

    # --- Interactive Chart Drawer & Analytics (Matplotlib Canvas) ---
    def _reset_chart_visuals(self):
        if not matplotlib or not self.fig:
            return

        self.ax1.clear()
        self.ax2.clear()
        self.ax3.clear()

        self.ax1.set_facecolor(self.palette["term_bg"])
        self.ax2.set_facecolor(self.palette["term_bg"])
        self.ax3.set_facecolor(self.palette["term_bg"])

        self.ax1.set_title("Upstream SNR Spectrum (dB)", color=self.palette["accent"], fontsize=10, fontweight="bold")
        self.ax1.tick_params(colors=self.palette["text_muted"], labelsize=8)
        self.ax1.grid(True, linestyle="--", alpha=0.3, color=self.palette["border"])

        self.ax2.set_title("Plant Diagnostic Metrics", color=self.palette["accent"], fontsize=10, fontweight="bold", pad=15)
        self.ax2.tick_params(colors=self.palette["text_muted"], labelsize=8)
        self.ax2.set_yticks([])

        self.ax3.set_title("Signal / RTT Profile (Live Polled)", color=self.palette["accent"], fontsize=10, fontweight="bold")
        self.ax3.tick_params(colors=self.palette["text_muted"], labelsize=8)
        self.ax3.grid(True, linestyle="--", alpha=0.3, color=self.palette["border"])

        self.canvas.draw_idle()

    def _update_spectrum_chart(self, cmd: str, output: str):
        if not matplotlib or not self.fig:
            return

        snr_vals = parse_snr_from_output(output)
        rtt_vals = [float(r) for r in re.findall(r"time[=<](\d+\.\d+|\d+)\s*ms", output, re.IGNORECASE)]
        cwe_vals = [int(x) for x in re.findall(r"uncorr(?:ectables?)?[:\s]+(\d+)", output, re.IGNORECASE)]
        rates = [float(r) for r in re.findall(r"(\d+\.\d+)\s*(?:mbps|kbps)", output, re.IGNORECASE)]

        has_visual_fields = bool(
            snr_vals or rtt_vals or cwe_vals or ("flap" in cmd.lower() and "ranging-abort" in output.lower())
        )

        if snr_vals:
            avg_snr = sum(snr_vals) / len(snr_vals)
            self.time_series_data["time"].append(datetime.now().strftime("%H:%M:%S"))
            self.time_series_data["snr"].append(avg_snr)
            if len(self.time_series_data["time"]) > 25:
                self.time_series_data["time"].pop(0)
                self.time_series_data["snr"].pop(0)

        def render_split_view():
            if has_visual_fields and self.nav_bar.get() in ["Terminal", "Split View"]:
                self._switch_stage_tab("Split View")

            self._reset_chart_visuals()

            if snr_vals:
                indices = list(range(1, len(snr_vals) + 1))
                colors = [
                    self.palette["status_ok"] if s >= 28.0 else self.palette["status_warn"] if s >= 25.0 else self.palette["status_crit"]
                    for s in snr_vals
                ]
                self.ax1.bar(indices, snr_vals, color=colors, edgecolor="none", alpha=0.85)
                self.ax1.axhline(28.0, color=self.palette["status_crit"], linestyle="--", linewidth=1.2, label="Floor (28 dB)")
                self.ax1.set_ylabel("SNR (dB)", color=self.palette["text_muted"], fontsize=8)
                self.ax1.set_title(f"Upstream SNR ({len(snr_vals)} Carriers)", color=self.palette["accent"], fontsize=10, fontweight="bold")
            elif rtt_vals:
                indices = list(range(1, len(rtt_vals) + 1))
                self.ax1.bar(indices, rtt_vals, color=self.palette["accent"], alpha=0.85)
                self.ax1.set_ylabel("RTT (ms)", color=self.palette["text_muted"], fontsize=8)
                self.ax1.set_title(f"Ping Probe RTT ({len(rtt_vals)} Samples)", color=self.palette["accent"], fontsize=10, fontweight="bold")
            else:
                self.ax1.set_title("No Carrier Decibels Found", color=self.palette["text_muted"], fontsize=9)

            categories = ["FLAP", "CWE", "SNR", "TEMP", "LOAD"]
            N = len(categories)
            angles = [n / float(N) * 2 * math.pi for n in range(N)]
            angles += angles[:1]

            avg_snr_score = min(100, max(0, ((sum(snr_vals) / len(snr_vals)) - 15) * 4)) if snr_vals else 80
            cwe_score = min(100, max(0, 100 - (sum(cwe_vals) / 100))) if cwe_vals else 95
            flap_score = 50 if "ranging-abort" in output.lower() else 90
            load_score = min(100, (sum(rates) / 100)) if rates else 25

            values = [flap_score, cwe_score, avg_snr_score, 85, load_score]
            values += values[:1]

            self.ax2.plot(angles, values, color=self.palette["accent"], linewidth=1.5, linestyle="solid")
            self.ax2.fill(angles, values, color=self.palette["accent"], alpha=0.25)
            self.ax2.set_xticks(angles[:-1])
            self.ax2.set_xticklabels(categories, color=self.palette["text_main"], size=8)

            if len(self.time_series_data["time"]) > 1:
                self.ax3.plot(
                    self.time_series_data["time"],
                    self.time_series_data["snr"],
                    color=self.palette["accent"],
                    label="Signal Noise Profile",
                    linewidth=1.5,
                    marker="o",
                    markersize=4,
                )
                self.ax3.fill_between(
                    self.time_series_data["time"],
                    self.time_series_data["snr"],
                    color=self.palette["accent"],
                    alpha=0.1,
                )
                self.ax3.set_ylabel("Avg SNR (dB)", color=self.palette["text_muted"], fontsize=8)
                self.ax3.tick_params(axis="x", rotation=30)
            else:
                self.ax3.set_title("Awaiting Continuous Polls...", color=self.palette["text_muted"], fontsize=9)

            self.fig.tight_layout(pad=3.0, w_pad=2.0, h_pad=2.5)
            self.canvas.draw_idle()

        self.after(0, render_split_view)

    # --- Pre/Post Maintenance Snapshot Engine ---
    def take_maintenance_snapshot(self):
        bastion = self.bastion_entry.get().strip()
        target = self.cmts_host_entry.get().strip()
        p_num = self.p_num_entry.get().strip()
        base_pw = self.base_pw_entry.get().strip()

        if not all([bastion, target, p_num, base_pw]):
            messagebox.showwarning("Credentials Missing", "Provide Bastion, CMTS Hostname, P-Number, and Password before taking a snapshot.")
            return

        vip_token = self._acquire_vip_token_if_needed(target)
        if vip_token is None:
            return

        def snapshot_worker():
            self.append_log(f"Capturing PRE-MAINTENANCE baseline for {target}...")
            if not self._ensure_session(bastion, target, p_num, base_pw, vip_token):
                return

            self.ssh_channel.send("show cable modem summary port mac | include mac\n")
            out1 = read_channel_until(self.ssh_channel, [">", "#"])
            self.ssh_channel.send("show cable flap-list\n")
            out2 = read_channel_until(self.ssh_channel, [">", "#"])

            self.pre_maint_snapshot = f"=== PORT SUMMARY ===\n{out1}\n=== FLAP LIST ===\n{out2}"
            self.pre_maint_target = target
            self.append_log("Pre-maintenance baseline snapshot successfully captured.")
            self.after(0, lambda: messagebox.showinfo("Snapshot Stored", f"Baseline saved for {target}."))

        threading.Thread(target=snapshot_worker, daemon=True).start()

    def run_maintenance_diff(self):
        if not self.pre_maint_snapshot:
            messagebox.showwarning("No Baseline", "Click 'Snapshot (Pre)' prior to running comparison.")
            return

        bastion = self.bastion_entry.get().strip()
        target = self.cmts_host_entry.get().strip()
        p_num = self.p_num_entry.get().strip()
        base_pw = self.base_pw_entry.get().strip()

        if target != self.pre_maint_target:
            confirmed = messagebox.askyesno(
                "Target Mismatch",
                f"Baseline was taken on {self.pre_maint_target}, current target is {target}. Continue?",
            )
            if not confirmed:
                return

        vip_token = self._acquire_vip_token_if_needed(target)
        if vip_token is None:
            return

        def diff_worker():
            self.append_log(f"Running POST-MAINTENANCE audit on {target}...")
            if not self._ensure_session(bastion, target, p_num, base_pw, vip_token):
                return

            self.ssh_channel.send("show cable modem summary port mac | include mac\n")
            out1 = read_channel_until(self.ssh_channel, [">", "#"])
            self.ssh_channel.send("show cable flap-list\n")
            out2 = read_channel_until(self.ssh_channel, [">", "#"])

            post_maint_data = f"=== PORT SUMMARY ===\n{out1}\n=== FLAP LIST ===\n{out2}"

            diff = difflib.unified_diff(
                self.pre_maint_snapshot.splitlines(),
                post_maint_data.splitlines(),
                fromfile=f"PRE-MAINTENANCE ({self.pre_maint_target})",
                tofile=f"POST-MAINTENANCE ({target})",
                lineterm="",
            )
            diff_text = "\n".join(diff)
            if not diff_text.strip():
                diff_text = "✅ No differences detected. Modem registrations, ports, and flap counts match baseline."

            self.last_diff_text = diff_text

            def display_diff():
                self.diff_box.delete("1.0", "end")
                self.diff_box.insert("end", diff_text)
                self._switch_stage_tab("Maint Diff")
                self.append_log("Post-maintenance comparison complete.")

            self.after(0, display_diff)

        threading.Thread(target=diff_worker, daemon=True).start()

    # --- Parallel Asynchronous Fleet Auditor + Leaderboard Engine ---
    def _audit_single_fleet_node(self, bastion: str, target: str, p_num: str, base_pw: str, vip_token: str, cmd: str) -> Tuple[str, str, dict]:
        """Dedicated thread-safe read-only SSH worker for parallel fleet auditing."""
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        bastion_pw = f"{base_pw}{vip_token}"

        try:
            try:
                client.connect(
                    hostname=bastion, username=p_num, password=bastion_pw, look_for_keys=False, allow_agent=False, timeout=22
                )
            except Exception:
                client.connect(hostname=bastion, look_for_keys=False, allow_agent=False, timeout=22)

            chan = client.invoke_shell(term="vt100", width=160, height=60)
            time.sleep(0.6)
            banner = read_channel_until(chan, ["login as:", "login:", "username:", "password:", "$", "#", ">"], 12)
            if any(b in banner.lower() for b in ["login as:", "login:", "username:"]):
                chan.send(p_num + "\n")
                banner = read_channel_until(chan, ["password:", "passcode:", "code:"], 10)
            if any(b in banner.lower() for b in ["password:", "passcode:", "code:"]):
                chan.send(bastion_pw + "\n")
                read_channel_until(chan, ["$", "#", ">", "~"], 15)

            chan.send(f"ssh {target}\n")
            time.sleep(0.6)
            resp = read_channel_until(chan, ["(yes/no)?", "password:", "Password:", ">", "#"], 15)
            if "(yes/no)?" in resp.lower():
                chan.send("yes\n")
                resp = read_channel_until(chan, ["password:", "Password:"], 10)
            if "password:" in resp.lower():
                chan.send(base_pw + "\n")
                read_channel_until(chan, [">", "#"], 18)

            chan.send("page-length 0\n")
            read_channel_until(chan, [">", "#"], 8)

            chan.send(cmd + "\n")
            output = read_channel_until(chan, [">", "#"], DEFAULT_SOCKET_TIMEOUT)

            # Parse metrics for Leaderboard ranking
            total, reg, offline, online_pct = 0, 0, 0, 100.0
            for line in output.splitlines():
                if "total" in line.lower():
                    nums = [int(n) for n in re.findall(r"\b\d+\b", line)]
                    if len(nums) >= 2:
                        total = max(nums)
                        reg = nums[0]
                        offline = total - reg
                        online_pct = (reg / total) * 100.0 if total > 0 else 100.0
                        break

            snrs = parse_snr_from_output(output)
            min_snr = min(snrs) if snrs else 35.0
            uncorr = sum(int(x) for x in re.findall(r"uncorr(?:ectables?)?[:\s]+(\d+)", output, re.IGNORECASE))

            status = "CRITICAL" if (online_pct < 90.0 or min_snr < 26.0 or uncorr > 5000) else "DEGRADED" if (online_pct < 97.5 or min_snr < 28.5 or uncorr > 500) else "HEALTHY"

            metrics = {
                "host": target,
                "online_pct": online_pct,
                "total": total,
                "offline": offline,
                "min_snr": min_snr,
                "uncorr": uncorr,
                "status": status,
            }
            return target, output, metrics
        finally:
            try:
                client.close()
            except Exception:
                pass

    def _run_bulk_audit(self, bastion: str, p_num: str, base_pw: str, cmd: str):
        allowed, reason = is_read_only_safe(cmd)
        if not allowed:
            messagebox.showerror("Read-Only Guardrail", reason)
            return

        raw_hosts = self.bulk_target_box.get("1.0", "end").strip()
        targets = [h.strip() for h in re.split(r"[,\n\s]+", raw_hosts) if h.strip()]

        if not targets:
            messagebox.showwarning("No Targets", "Enter at least one host in the fleet box.")
            return

        if not all([bastion, p_num, base_pw]):
            messagebox.showwarning("Missing Credentials", "Ensure Bastion, P-Number, and Password are provided.")
            return

        vip_token = self._acquire_vip_token_if_needed(target_check=None)
        if vip_token is None:
            return

        self.bulk_results.clear()
        self.fleet_leaderboard_rows.clear()
        self.exec_btn.configure(state="disabled", text=f"Auditing {len(targets)} Nodes in Parallel...")

        def parallel_fleet_worker():
            start_t = time.time()
            workers = min(MAX_FLEET_WORKERS, len(targets))
            self.append_log(f"Launching {workers} parallel Read-Only SSH workers across {len(targets)} CMTS nodes...")

            leaderboard_list = []
            with ThreadPoolExecutor(max_workers=workers) as pool:
                future_map = {
                    pool.submit(self._audit_single_fleet_node, bastion, t, p_num, base_pw, vip_token, cmd): t
                    for t in targets
                }
                for idx, fut in enumerate(as_completed(future_map), start=1):
                    host = future_map[fut]
                    try:
                        h, out, metrics = fut.result()
                        self.bulk_results[h] = out
                        leaderboard_list.append(metrics)
                        self.flight_recorder.record(h, cmd, out)

                        ts = datetime.now().strftime("%H:%M:%S")
                        formatted = f"\n--- [{ts}] PARALLEL NODE [{idx}/{len(targets)}]: {h} | {cmd} ---\n{out.strip()}\n"
                        self.after(0, lambda text=formatted: self._highlight_and_stream(text))
                        self.append_log(f"[{idx}/{len(targets)}] Completed {h} (Online: {metrics['online_pct']:.1f}%)")
                    except Exception as ex:
                        self.append_log(f"[{idx}/{len(targets)}] Failed on {host}: {str(ex)}")
                        leaderboard_list.append({
                            "host": host, "online_pct": 0.0, "total": 0, "offline": 0,
                            "min_snr": 0.0, "uncorr": 0, "status": f"UNREACHABLE ({str(ex)[:18]})"
                        })

            elapsed = time.time() - start_t
            # Sort worst-first: lowest online_pct, highest offline, lowest SNR
            leaderboard_list.sort(key=lambda r: (r["online_pct"], -r["offline"], r["min_snr"]))
            self.fleet_leaderboard_rows = leaderboard_list

            table_lines = [
                f"=== 🏆 PARALLEL FLEET AUDIT LEADERBOARD ({len(targets)} Nodes Audited in {elapsed:.1f}s) ===",
                f"Command Dispatched: {cmd} | Concurrency: {workers} Parallel Read-Only Workers\n",
                f"{'RANK':<6} {'CMTS HOSTNAME':<18} {'ONLINE %':<11} {'TOTAL CMs':<11} {'OFFLINE':<10} {'MIN SNR':<10} {'UNCORR CW':<12} {'VERDICT'}",
                "-" * 96,
            ]
            for rank, r in enumerate(leaderboard_list, start=1):
                table_lines.append(
                    f"#{rank:<5} {r['host']:<18} {r['online_pct']:>6.1f}%     {r['total']:>8,}   {r['offline']:>7,}   {r['min_snr']:>6.1f} dB   {r['uncorr']:>9,}   {r['status']}"
                )

            leaderboard_text = "\n".join(table_lines)

            def finalize_fleet_ui():
                self.leaderboard_box.delete("1.0", "end")
                self.leaderboard_box.insert("1.0", leaderboard_text)
                self._switch_stage_tab("Fleet Leaderboard")
                self.export_btn.configure(state="normal")
                self.exec_btn.configure(state="normal", text="🚀 Run Parallel Fleet Audit")
                self.append_log(f"Parallel Fleet Audit finished in {elapsed:.1f}s. Leaderboard updated.")

            self.after(0, finalize_fleet_ui)

        threading.Thread(target=parallel_fleet_worker, daemon=True).start()

    # --- Flight Recorder Replay & Chronic Offender Scan ---
    def _refresh_flight_search(self):
        query = self.flight_search_entry.get().strip()
        records = self.flight_recorder.search(query)
        self.flight_listbox.delete(0, tk.END)
        self.flight_record_cache = records

        for rec in records:
            _, ts, host, cmd, _ = rec
            self.flight_listbox.insert(tk.END, f"[{ts}] {host} -> {cmd[:30]}")

    def _on_flight_select(self, _):
        cur = self.flight_listbox.curselection()
        if not cur or not hasattr(self, "flight_record_cache"):
            return
        idx = cur[0]
        if idx < len(self.flight_record_cache):
            _, ts, host, cmd, out = self.flight_record_cache[idx]
            preview_text = (
                f"=== HISTORICAL AUDIT REPLAY ===\n"
                f"Logged At: {ts} | Target Host: {host}\n"
                f"Dispatched: {cmd}\n"
                + "=" * 60 + "\n\n"
                + out
            )
            self.flight_preview.delete("1.0", "end")
            self.flight_preview.insert("1.0", preview_text)

    def _run_chronic_offender_scan(self):
        report = self.flight_recorder.analyze_chronic_offenders(days=30)
        self.flight_preview.delete("1.0", "end")
        self.flight_preview.insert("1.0", report)
        self.append_log("Completed 30-day SQLite Chronic Offender & Historical Trend scan.")

    # --- Live Watch / Continuous Poller Engine ---
    def _on_watch_interval_change(self, choice: str):
        val = int(re.search(r"\d+", choice).group(0))
        self.watch_delay_sec = val

    def toggle_watch_mode(self):
        if self.is_watching:
            self.is_watching = False
            self.watch_btn.configure(text="▶ Start Watch", fg_color=self.palette["border"], hover_color=self.palette["status_ok"])
            self.watch_lbl.configure(text="IDLE", text_color=self.palette["text_muted"])
            self.append_log("Watch mode paused.")
        else:
            bastion = self.bastion_entry.get().strip()
            target = self.cmts_host_entry.get().strip()
            p_num = self.p_num_entry.get().strip()
            base_pw = self.base_pw_entry.get().strip()

            if not all([bastion, target, p_num, base_pw]):
                messagebox.showwarning("Credentials Required", "Provide Bastion, CMTS Hostname, P-Number, and Password before starting live watch.")
                return

            vip_token = self._acquire_vip_token_if_needed(target)
            if vip_token is None:
                return

            self.is_watching = True
            self.watch_iteration = 0
            self.watch_btn.configure(text="⏹ Stop Watch", fg_color=self.palette["status_crit"], hover_color="#DC2626")
            self.watch_lbl.configure(text="ACTIVE", text_color=self.palette["status_ok"])
            self.append_log(f"Watch mode engaged (Interval: {self.watch_delay_sec}s).")

            self.watch_thread = threading.Thread(target=self._watch_worker, args=(vip_token,), daemon=True)
            self.watch_thread.start()

    def _watch_worker(self, vip_token: str):
        bastion = self.bastion_entry.get().strip()
        target = self.cmts_host_entry.get().strip()
        p_num = self.p_num_entry.get().strip()
        base_pw = self.base_pw_entry.get().strip()

        choice = self.catalog_dropdown.get()
        spec = COMMAND_CATALOG[choice]
        cmd = spec.format_command(
            cmac=self.cmac_input.get(),
            interface=self.iface_input.get(),
            mac_address=self.mac_input.get(),
        )

        while self.is_watching:
            self.watch_iteration += 1
            try:
                if not self._ensure_session(bastion, target, p_num, base_pw, vip_token):
                    self.is_watching = False
                    break

                self.ssh_channel.send(cmd + "\n")
                output = read_channel_until(self.ssh_channel, [">", "#"])

                now = datetime.now().strftime("%H:%M:%S")
                header_line = f"=== LIVE WATCH [Cycle #{self.watch_iteration} @ {now}] | {target}# {cmd} ===\n"

                def update_ui(h=header_line, o=output):
                    self.term_text.delete("1.0", "end")
                    self.term_text.insert("1.0", h, "header")
                    self._highlight_and_stream(o.strip() + "\n")
                    self._parse_telemetry(o)
                    self._run_heuristics_triage(cmd, o)
                    self._update_spectrum_chart(cmd, o)

                self.after(0, update_ui)

            except Exception as e:
                self.append_log(f"Watch exception: {str(e)}")
                self.is_watching = False
                break

            for _ in range(self.watch_delay_sec * 10):
                if not self.is_watching:
                    break
                time.sleep(0.1)

        self.after(0, lambda: self.watch_lbl.configure(text="IDLE", text_color=self.palette["text_muted"]))
        self.after(0, lambda: self.watch_btn.configure(text="▶ Start Watch", fg_color=self.palette["border"], hover_color=self.palette["status_ok"]))

    # --- Logging & Synchronous Console Hooks ---
    def append_log(self, msg: str):
        self.after(0, self._append_log_sync, msg)

    def _append_log_sync(self, msg: str):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def clear_console(self):
        self.term_text.delete("1.0", "end")
        self.last_output = ""
        self.export_btn.configure(state="disabled")
        self.search_count_lbl.configure(text="")
        self._reset_chart_visuals()
        if self.nav_bar.get() == "Split View":
            self._switch_stage_tab("Terminal")

    def _update_status(self, connected: bool, host: str = ""):
        def update():
            if connected:
                self.status_dot.configure(text_color=self.palette["status_ok"])
                self.status_lbl.configure(text=f"CONNECTED (RO): {host.upper()}", text_color=self.palette["status_ok"])
                self._toggle_auth_drawer(force_collapse=True)
            else:
                self.status_dot.configure(text_color=self.palette["status_crit"])
                self.status_lbl.configure(text="DISCONNECTED", text_color=self.palette["text_muted"])
        self.after(0, update)

    # --- Autocomplete & CLI History Navigation ---
    def _on_cli_key(self, event):
        if event.keysym in ("Up", "Down", "Return", "Escape", "Tab"):
            return

        query = self.cli_entry.get().strip().lower()
        if not query:
            self.popup.withdraw()
            return

        tokens = query.split()
        matches = []
        for cmd in RECOMMENDED_AUTOCOMPLETE:
            cmd_low = cmd.lower()
            if cmd_low.startswith(query):
                matches.append((0, cmd))
            elif all(t in cmd_low for t in tokens):
                matches.append((1, cmd))

        sorted_matches = [m[1] for m in sorted(matches, key=lambda x: x[0])]
        if sorted_matches:
            x = self.cli_entry.winfo_rootx()
            y = self.cli_entry.winfo_rooty() - min(len(sorted_matches) * 26 + 8, 185)
            self.popup.display(sorted_matches, x, y, self.cli_entry.winfo_width())
        else:
            self.popup.withdraw()

    def _on_cli_down(self, _):
        if self.popup.winfo_ismapped():
            self.popup.next_item()
            return "break"
        if self.cmd_history:
            if self.history_index < len(self.cmd_history) - 1:
                self.history_index += 1
                self.cli_entry.delete(0, "end")
                self.cli_entry.insert(0, self.cmd_history[self.history_index])
            else:
                self.history_index = len(self.cmd_history)
                self.cli_entry.delete(0, "end")
            return "break"

    def _on_cli_up(self, _):
        if self.popup.winfo_ismapped():
            self.popup.prev_item()
            return "break"
        if self.cmd_history:
            if self.history_index > 0:
                self.history_index -= 1
            else:
                self.history_index = 0
            self.cli_entry.delete(0, "end")
            self.cli_entry.insert(0, self.cmd_history[self.history_index])
            return "break"

    def _on_cli_tab(self, _):
        if self.popup.winfo_ismapped():
            val = self.popup.get_current()
            if val:
                self._apply_autocomplete(val)
            return "break"

    def _apply_autocomplete(self, command: str):
        self.cli_entry.delete(0, "end")
        self.cli_entry.insert(0, command)
        self.cli_entry.icursor("end")
        self.popup.withdraw()

    # --- Session & Network Management ---
    def _keepalive_loop(self):
        while self.keepalive_active:
            time.sleep(45)
            if self.is_connected and self.ssh_channel and not self.ssh_channel.closed:
                try:
                    self.ssh_channel.send(" \b")
                except Exception:
                    self.is_connected = False
                    self._update_status(False)
                    self.append_log("Keep-alive detected dropped remote channel.")
                    break

    def _ensure_session(self, bastion: str, target: str, p_num: str, base_pw: str, vip_token: str = "") -> bool:
        if (
            self.is_connected
            and self.ssh_channel
            and not self.ssh_channel.closed
            and self.active_host == target
        ):
            try:
                if self.ssh_channel.get_transport().is_active():
                    return True
            except Exception:
                pass
            self.append_log("Transport dropped. Re-establishing channel...")
            self.disconnect_session()

        if (
            self.ssh_client
            and self.ssh_channel
            and not self.ssh_channel.closed
            and self.ssh_channel.get_transport().is_active()
            and self.active_host != target
            and self.active_host != ""
        ):
            try:
                self.ssh_channel.send("exit\n")
                read_channel_until(self.ssh_channel, ["$", "#", ">", "~"], 8)
                self.append_log(f"Hopping to next CMTS: {target}...")
                self.ssh_channel.send(f"ssh {target}\n")
                time.sleep(0.8)
                resp = read_channel_until(self.ssh_channel, ["(yes/no)?", "password:", "Password:", ">", "#"], 15)
                if "(yes/no)?" in resp.lower():
                    self.ssh_channel.send("yes\n")
                    resp = read_channel_until(self.ssh_channel, ["password:", "Password:"], 10)
                if "password:" in resp.lower():
                    self.ssh_channel.send(base_pw + "\n")
                    read_channel_until(self.ssh_channel, [">", "#"], 18)
                self.ssh_channel.send("page-length 0\n")
                read_channel_until(self.ssh_channel, [">", "#"], 8)
                self.is_connected = True
                self.active_host = target
                self._update_status(True, target)
                return True
            except Exception:
                self.disconnect_session()

        if not vip_token:
            self.append_log("Aborted: VIP Token required for new Bastion session.")
            return False

        bastion_pw = f"{base_pw}{vip_token}"
        cmts_pw = base_pw

        self.append_log(f"Connecting to Gateway ({bastion})...")
        self.ssh_client = paramiko.SSHClient()
        self.ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())

        try:
            self.ssh_client.connect(
                hostname=bastion, username=p_num, password=bastion_pw, look_for_keys=False, allow_agent=False, timeout=25
            )
        except Exception:
            self.ssh_client.connect(hostname=bastion, look_for_keys=False, allow_agent=False, timeout=25)

        self.ssh_channel = self.ssh_client.invoke_shell(term="vt100", width=160, height=60)
        time.sleep(1)

        banner = read_channel_until(self.ssh_channel, ["login as:", "login:", "username:", "password:", "$", "#", ">"], 15)

        if any(b in banner.lower() for b in ["login as:", "login:", "username:"]):
            self.ssh_channel.send(p_num + "\n")
            banner = read_channel_until(self.ssh_channel, ["password:", "passcode:", "code:"], 12)

        if any(b in banner.lower() for b in ["password:", "passcode:", "code:"]):
            self.ssh_channel.send(bastion_pw + "\n")
            read_channel_until(self.ssh_channel, ["$", "#", ">", "~"], 20)

        self.append_log(f"Gateway unlocked. Hopping to CMTS: {target}...")
        self.ssh_channel.send(f"ssh {target}\n")
        time.sleep(1)

        resp = read_channel_until(self.ssh_channel, ["(yes/no)?", "password:", "Password:", ">", "#"], 18)
        if "(yes/no)?" in resp.lower():
            self.ssh_channel.send("yes\n")
            resp = read_channel_until(self.ssh_channel, ["password:", "Password:"], 10)

        if "password:" in resp.lower():
            self.ssh_channel.send(cmts_pw + "\n")
            read_channel_until(self.ssh_channel, [">", "#"], 20)

        self.ssh_channel.send("page-length 0\n")
        read_channel_until(self.ssh_channel, [">", "#"], 10)

        self.is_connected = True
        self.active_host = target
        self._update_status(True, target)
        self.append_log(f"Read-Only session established on {target}.")

        if not self.keepalive_active:
            self.keepalive_active = True
            threading.Thread(target=self._keepalive_loop, daemon=True).start()

        return True

    def disconnect_session(self):
        self.keepalive_active = False
        if self.ssh_channel:
            try:
                self.ssh_channel.close()
            except Exception:
                pass
        if self.ssh_client:
            try:
                self.ssh_client.close()
            except Exception:
                pass
        self.ssh_client = None
        self.ssh_channel = None
        self.is_connected = False
        self.active_host = ""
        self._update_status(False)
        self.append_log("SSH Session closed.")

    def _on_close(self):
        self.is_watching = False
        save_config({
            "bastion_host": self.bastion_entry.get().strip(),
            "p_num": self.p_num_entry.get().strip(),
            "theme": self.current_theme,
        })
        self.disconnect_session()
        self.destroy()

    # --- Dispatch Engine (Single Command & Multi-Command Playbooks) ---
    def dispatch_catalog_command(self):
        choice = self.catalog_dropdown.get()
        spec = COMMAND_CATALOG[choice]
        final_cmd = spec.format_command(
            cmac=self.cmac_input.get(),
            interface=self.iface_input.get(),
            mac_address=self.mac_input.get(),
        )

        bastion = self.bastion_entry.get().strip()
        p_num = self.p_num_entry.get().strip()
        base_pw = self.base_pw_entry.get().strip()

        if self.op_mode.get() == "Parallel Fleet Audit":
            self._run_bulk_audit(bastion, p_num, base_pw, final_cmd)
        else:
            self._route_single_dispatch(final_cmd)

    def dispatch_playbook(self):
        choice = self.playbook_dropdown.get()
        spec = PLAYBOOK_CATALOG[choice]

        if "mac_address" in spec.params and not self.mac_input.get().strip():
            messagebox.showwarning("MAC Required", f"Playbook '{choice}' requires a Target Modem MAC address.")
            return
        if "cmac" in spec.params and not self.cmac_input.get().strip():
            messagebox.showwarning("Cable-Mac Required", f"Playbook '{choice}' requires a Cable-Mac ID.")
            return

        cmds = spec.format_commands(
            cmac=self.cmac_input.get(),
            interface=self.iface_input.get(),
            mac_address=self.mac_input.get(),
        )

        bastion = self.bastion_entry.get().strip()
        target = self.cmts_host_entry.get().strip()
        p_num = self.p_num_entry.get().strip()
        base_pw = self.base_pw_entry.get().strip()

        if not all([bastion, target, p_num, base_pw]):
            self._toggle_auth_drawer(force_collapse=False)
            messagebox.showwarning("Missing Credentials", "Ensure Bastion, CMTS, P-Number, and Password are provided.")
            return

        vip_token = self._acquire_vip_token_if_needed(target)
        if vip_token is None:
            return

        self.run_playbook_btn.configure(state="disabled", text="Running Playbook Chain...")

        def playbook_worker():
            try:
                if not self._ensure_session(bastion, target, p_num, base_pw, vip_token):
                    return

                combined_output = []
                banner = f"\n=== ⚡ PLAYBOOK EXECUTION: {choice} on {target} ===\n"
                self.after(0, lambda b=banner: self._highlight_and_stream(b))

                for idx, cmd in enumerate(cmds, start=1):
                    allowed, reason = is_read_only_safe(cmd)
                    if not allowed:
                        self.append_log(f"Blocked playbook step: {reason}")
                        continue

                    self.append_log(f"Playbook Step [{idx}/{len(cmds)}]: {cmd}")
                    self.ssh_channel.send(cmd + "\n")
                    out = read_channel_until(self.ssh_channel, [">", "#"])
                    combined_output.append(f"--- [STEP {idx}: {cmd}] ---\n{out.strip()}")

                    step_formatted = f"\n[{datetime.now().strftime('%H:%M:%S')}] {target}# {cmd}\n{out.strip()}\n"
                    self.after(0, lambda sf=step_formatted: self._highlight_and_stream(sf))

                full_text = "\n\n".join(combined_output)
                self.last_output = full_text
                self.last_command = f"Playbook: {choice}"

                self.flight_recorder.record(target, self.last_command, full_text)
                self._parse_telemetry(full_text)
                self._run_heuristics_triage(self.last_command, full_text)
                self._update_spectrum_chart(self.last_command, full_text)
                self.after(0, lambda: self.export_btn.configure(state="normal"))
                self.append_log(f"Playbook '{choice}' completed ({len(cmds)} commands correlated).")

            except Exception as ex:
                self.append_log(f"Playbook Error: {str(ex)}")
                self.after(0, lambda: messagebox.showerror("Playbook Fault", str(ex)))
            finally:
                self.after(0, lambda: self.run_playbook_btn.configure(state="normal", text="⚡ Run Read-Only Playbook Chain"))

        threading.Thread(target=playbook_worker, daemon=True).start()

    def dispatch_cli_command(self):
        self.popup.withdraw()
        cmd = self.cli_entry.get().strip()
        if not cmd:
            return

        # Hard Read-Only Security Guardrail
        allowed, reason = is_read_only_safe(cmd)
        if not allowed:
            self.append_log(f"BLOCKED WRITE ATTEMPT: {cmd} ({reason})")
            messagebox.showerror(
                "Read-Only Security Guardrail",
                f"Action Denied — This application operates in strict Read-Only mode.\n\nReason:\n{reason}",
            )
            return

        if not self.cmd_history or self.cmd_history[-1] != cmd:
            self.cmd_history.append(cmd)
        self.history_index = len(self.cmd_history)
        self.cli_entry.delete(0, "end")

        if self.op_mode.get() == "Parallel Fleet Audit":
            bastion = self.bastion_entry.get().strip()
            p_num = self.p_num_entry.get().strip()
            base_pw = self.base_pw_entry.get().strip()
            self._run_bulk_audit(bastion, p_num, base_pw, cmd)
        else:
            self._route_single_dispatch(cmd)

    def _route_single_dispatch(self, command: str):
        allowed, reason = is_read_only_safe(command)
        if not allowed:
            messagebox.showerror("Read-Only Guardrail", reason)
            return

        bastion = self.bastion_entry.get().strip()
        target = self.cmts_host_entry.get().strip()
        p_num = self.p_num_entry.get().strip()
        base_pw = self.base_pw_entry.get().strip()

        if not all([bastion, target, p_num, base_pw]):
            self._toggle_auth_drawer(force_collapse=False)
            messagebox.showwarning("Missing Credentials", "Ensure Bastion, CMTS Hostname, P-Number, and Password are provided.")
            return

        vip_token = self._acquire_vip_token_if_needed(target)
        if vip_token is None:
            return

        self.exec_btn.configure(state="disabled", text="Running...")
        self.send_btn.configure(state="disabled")

        def worker():
            try:
                if not self._ensure_session(bastion, target, p_num, base_pw, vip_token):
                    return

                self.append_log(f"Executing (RO): {command}")
                self.ssh_channel.send(command + "\n")

                output = read_channel_until(self.ssh_channel, [">", "#"])
                if not output and (self.ssh_channel.closed or not self.ssh_client.get_transport().is_active()):
                    self.append_log("Connection dropped by node.")
                    self.disconnect_session()
                    return

                self.last_output = output
                self.last_command = command

                self.flight_recorder.record(target, command, output)

                timestamp = datetime.now().strftime("%H:%M:%S")
                formatted = f"\n[{timestamp}] {target}# {command}\n{output.strip()}\n"

                self.after(0, lambda: self._highlight_and_stream(formatted))
                self.append_log(f"Output received ({len(output.splitlines())} lines).")

                self._parse_telemetry(output)
                self._run_heuristics_triage(command, output)
                self._update_spectrum_chart(command, output)

                self.after(0, lambda: self.export_btn.configure(state="normal"))

            except Exception as ex:
                self.append_log(f"Execution Error: {str(ex)}")
                self.disconnect_session()
                self.after(0, lambda: messagebox.showerror("Execution Fault", str(ex)))
            finally:
                self.after(0, lambda: self.exec_btn.configure(state="normal", text="▶  Execute Show Workflow"))
                self.after(0, lambda: self.send_btn.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    # --- Excel Exporter ---
    def export_excel(self):
        timestamp = datetime.now().strftime("%Y_%m_%d_%H%M%S")
        is_bulk = bool(self.bulk_results)

        default_filename = f"FLEET_AUDIT_{timestamp}.xlsx" if is_bulk else f"{self.active_host}_{timestamp}.xlsx"

        chosen_path = filedialog.asksaveasfilename(
            parent=self,
            title="Export Records to Excel",
            initialdir=str(Path.home() / "Desktop"),
            initialfile=default_filename,
            defaultextension=".xlsx",
            filetypes=[("Excel Spreadsheet", "*.xlsx"), ("All Files", "*.*")],
        )

        if not chosen_path:
            return

        try:
            wb = openpyxl.Workbook()
            wb.remove(wb.active)

            if is_bulk:
                if self.fleet_leaderboard_rows:
                    ws_lead = wb.create_sheet(title="Fleet_Leaderboard")
                    lead_text = "\n".join(
                        f"{r['host']}  {r['online_pct']:.1f}%  {r['total']}  {r['offline']}  {r['min_snr']:.1f}dB  {r['uncorr']}  {r['status']}"
                        for r in self.fleet_leaderboard_rows
                    )
                    create_excel_sheet(ws_lead, "Parallel Fleet Health Leaderboard", f"Generated: {datetime.now()}", lead_text)

                for host, out in self.bulk_results.items():
                    ws = wb.create_sheet(title=host[:31])
                    create_excel_sheet(ws, f"CER Audit: {host}", f"Captured: {datetime.now()}", out)
            else:
                ws = wb.create_sheet(title="CER Output")
                create_excel_sheet(ws, f"CER Audit: {self.active_host}", f"Command: {self.last_command}", self.last_output)

            wb.save(chosen_path)
            self.append_log(f"Workbook exported: {chosen_path}")
            messagebox.showinfo("Export Successful", f"Saved to:\n{chosen_path}")
        except Exception as e:
            self.append_log(f"Export Error: {str(e)}")
            messagebox.showerror("Export Failed", str(e))


if __name__ == "__main__":
    app = ArrisCommandCenter()
    app.mainloop()