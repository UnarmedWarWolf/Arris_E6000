# Arris_E6000

Projects for the ARRIS E6000 CMTS chassis.

## CMTS Command Center (v0.3.2)

Operations console for ARRIS / CommScope E6000 CER chassis, reached directly over SSH or through a bastion (jump) host with a one-time VIP token.

### Run

```
pip install paramiko matplotlib openpyxl
python CMTS_Command_Center.py
```

Settings and history are stored in `~/.e6000_command_center`.

### Install on Windows

Download `CMTS_Command_Center_Setup_<version>.exe` from [Releases](https://github.com/UnarmedWarWolf/Arris_E6000/releases) and run it. The installer:

- Installs to `C:\Program Files\CMTS Command Center` by default. You can choose another folder.
- Adds a Start Menu entry and, unless you untick it, a desktop shortcut.
- Can install for all users (needs admin) or for just you (no admin; goes to `%LOCALAPPDATA%\Programs`).
- Upgrades an existing install in place. Your settings and history are kept.
- Appears in **Settings → Apps** for uninstall. Uninstall asks whether to also delete your settings and history (default: keep).

### Build the installer

Needs Python 3.10+ and [Inno Setup 6](https://jrsoftware.org/isdl.php) (`winget install JRSoftware.InnoSetup`).

```
build_installer.bat
```

This installs the build dependencies, runs PyInstaller (`dist\CMTS_Command_Center\`), then Inno Setup. The installer is written to `dist\installer\`.

To build it on GitHub instead, push a version tag. The **Build Windows installer** workflow builds the setup .exe and attaches it to a Release:

```
git tag v0.3.2
git push origin v0.3.2
```

Build output isn't committed to the repo.

### What's new in 0.3.2
- Windows installer (Inno Setup) with install-location choice, desktop and Start Menu shortcuts, and an uninstaller.
- App icon on the .exe, window and taskbar.
- PyInstaller now produces a folder build, which starts faster than the old one-file .exe.
- Save dialogs for the HTML report and Excel workbook now open in your last export folder or Documents. Before, they could open in the install folder, which isn't writable under Program Files.

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
