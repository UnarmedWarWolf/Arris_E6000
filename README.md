# Arris_E6000

Projects for the ARRIS E6000 CMTS chassis.

## CMTS Command Center (v0.3.0)

Operations console for ARRIS / CommScope E6000 CER chassis, reached directly over SSH or through a bastion (jump) host with a one-time VIP token.

### Run

```
pip install customtkinter paramiko matplotlib openpyxl
python CMTS_Command_Center.py
```

Settings and history are stored in `~/.e6000_command_center`.

### Build a Windows .exe

Double-click `build.bat` (or run it from a command prompt in the repo folder). It creates a clean `.venv`, installs everything, and builds.

Manual equivalent:

```
pip install pyinstaller customtkinter paramiko matplotlib openpyxl
pyinstaller --clean --noconfirm CMTS_Command_Center.spec
```

If Windows Defender deletes the exe right after it's built, add an exclusion for the repo's `dist\` folder.

The executable is written to `dist/`. Build output isn't committed to the repo. Publish the .exe as a GitHub Release instead.

### What's new in 0.3.0

**Fixed from 0.2.x**
- Commands no longer get cut off or mixed up. Reads wait for the real CMTS prompt, including config mode.
- `--More--` pagers are answered automatically. The wide terminal keeps tables from wrapping.
- The VIP token is asked for once. One bastion session stays open and CMTS hops reuse it.
- The UI no longer freezes. SSH runs on a single background worker and parsing happens off the UI thread.
- Read-only guardrail removed. Disruptive commands ask for confirmation, which can be turned off in Settings.
- HTML export is escaped. Modem counts are read from the correct columns.

**New**
- Threshold-based coloring for SNR, MER, Tx/Rx power, codewords, utilization, temperature, flaps, optical Rx, modem state and port online %. Thresholds are editable in Settings.
- Sortable, filterable Table view with an "issues only" filter.
- Redesigned charts, triage (outside-plant detection, blast radius), color-coded fleet leaderboard, maintenance diff, history replay, Stop (Ctrl-C) button.
- Direct SSH when no bastion is set.

### Needs verification on a real E6000
- Paging command (`terminal length 0` / `page-length 0`).
- Column headers for `show cable noise`, `show cable modem ... phy` and summary outputs.
