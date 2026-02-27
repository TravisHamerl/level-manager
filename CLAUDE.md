# Mastercam Level Manager

## Overview

A Windows GUI application for toggling Mastercam level (layer) visibility using global hotkeys. Connects to Mastercam's docked Levels panel via Windows UI Automation and provides keyboard shortcuts for rapid level toggling without switching focus.

## Project Structure

```
level-manager/
├── level_manager.py          # Main GUI app (tkinter + pynput + pywinauto)
├── level_toggle.py           # CLI utility for single-level hotkey toggle
├── Level Manager.bat         # Windows launcher (finds Python automatically)
├── inspect_level_items.py    # Dev tool: deep-inspect Levels tree items and UIA patterns
├── inspect_levels_panel.py   # Dev tool: discover Levels panel via Win32 enumeration
├── test_level_connect.py     # Dev tool: quick connection test (find panel → toggle level 202)
├── .gitignore
└── CLAUDE.md                 # This file
```

## Prerequisites

```bash
pip install pywinauto pynput sv-ttk
```

Python 3.10+ required. Python 3.14 installed at `C:\Program Files\Python314\python.exe`.

## Usage

### GUI App
```
python level_manager.py
# or double-click "Level Manager.bat"
```

1. Open and dock the **Levels panel** in Mastercam
2. **Scroll the levels you need into view** — the tree is virtualized and only visible items are accessible
3. Click **Connect** to discover the panel and scan levels
4. **Double-click** a level to toggle its visibility
5. **Right-click** or double-click the Hotkey column to assign a global hotkey
6. Use assigned hotkeys from anywhere (no need to focus the app)
7. If hotkeys stop working, the app will notify you — click **Refresh** to reconnect

### CLI Toggle
```bash
python level_toggle.py              # Toggle level 202 with Ctrl+Alt+I
python level_toggle.py 210          # Toggle level 210
python level_toggle.py 202 --key f9 # Toggle level 202 with F9
```
Press Escape to quit.

## Architecture

### Mastercam Connection

Two-phase discovery with GUID caching for speed:

1. **Cached GUID fast path** — On first successful connect, the Levels panel's HwndWrapper GUID is saved to settings. Subsequent connects use this GUID for instant Win32 lookup (~4s vs ~11s)
2. **Slow probe fallback** — If no cached GUID (first run) or cache is stale (Mastercam update), enumerates all HwndWrapper child windows and tests each with UIA until finding one containing `LevelTreeListBox`. Caches the discovered GUID for next time

The GUID is stored in `%LOCALAPPDATA%` (per-machine), so each computer discovers and caches its own GUID independently.

### Virtualized Levels Tree

The Mastercam Levels tree is **virtualized** — only items currently visible in the viewport have UIA elements. Items scrolled off-screen are not accessible. Users must scroll the levels they need into view before connecting.

### Levels Panel Structure (UIA)
```
LevelTreeListBox (Tree, auto_id="LevelTreeListBox")
  ├─ GroupTreeItem (TreeItem — group header, expandable)
  │   ├─ LevelTreeItem (TreeItem)
  │   │   ├─ Edit (level number, e.g. "202")
  │   │   ├─ Edit (level name, e.g. "Geometry")
  │   │   └─ Button (auto_id="IsLevelVisibleButton" — Invoke pattern only)
  │   └─ ... more levels
  └─ LevelTreeItem (ungrouped levels at root)
```

### Visibility Toggle

Uses UIA `Invoke` pattern on `IsLevelVisibleButton` — no mouse movement required. The Toggle pattern is **not available** on these buttons; only Invoke is supported.

**Stale reference notification:** If `Invoke()` throws an exception (stale COM reference after Mastercam changes levels), the app shows a notification with a bell/flash and prompts the user to click Refresh. No automatic reconnection (which would cause a ~12s freeze).

### Hotkey System

- **Library:** pynput (`keyboard.Listener` in a daemon thread)
- **Modifier tracking:** Ctrl, Alt, Shift state tracked via `on_press`/`on_release`
- **Thread marshalling:** Hotkey callbacks fire via `root.after(0, callback)` to run in tkinter thread
- **Recording mode:** Separate `_rec_listener` captures new hotkey combos; main listener paused during recording to avoid conflicts
- Supports letters, numbers, and F-keys with any modifier combination
- Conflict detection prevents duplicate bindings

### Groups

- Select 2+ levels and click **Group** to create a named group
- Group hotkeys toggle all member levels in one keypress
- Groups display as collapsible tree nodes in the GUI
- Auto-generated names from member level names

### Settings Persistence

Stored at `%LOCALAPPDATA%\LevelManager\settings.json` (per-machine, not in Dropbox):
- Hotkey assignments (per level number)
- Group definitions and group hotkeys
- Cached Levels panel GUID (for fast reconnect)
- Window geometry

## Key Functions

### Connection (`level_manager.py`)

| Function | Description |
|----------|-------------|
| `_find_mastercam_hwnd()` | Win32 fast discovery of main Mastercam window |
| `_find_levels_by_guid(mc_hwnd, guid)` | Win32 search for HwndWrapper with specific GUID |
| `connect_levels_panel()` | Full connect: returns `(panel, tree, error_msg)` with GUID caching |
| `scan_levels(tree)` | Scan visible tree items, return list of `{number, name, item, vis_btn}` dicts |

### Toggle (`level_manager.py`)

| Function | Description |
|----------|-------------|
| `toggle_visibility(level)` | Invoke button; returns `True` on success, `False` on exception |
| `_notify_stale()` | Show stale connection notification with bell/flash |
| `_hotkey_toggle(number)` | Single level toggle; notifies on stale reference |
| `_hotkey_toggle_group(numbers, name)` | Group toggle with partial failure reporting |

### Hotkey Helpers (`level_manager.py`)

| Function | Description |
|----------|-------------|
| `hotkey_to_str(hk)` | Convert hotkey dict to display string (e.g. "Ctrl+F9") |
| `hotkey_matches(hk, key, ctrl, alt, shift)` | Check if pressed key matches a hotkey config |

## Troubleshooting

### "Levels panel not found"
- Ensure the Levels panel is open/docked in Mastercam (not just the tab — the panel itself must be visible)
- Click **Connect** / **Refresh** to retry

### Hotkeys stop working / "Connection stale"
- The app notifies you when references go stale — click **Refresh** to reconnect
- This happens when levels are added, deleted, or reordered in Mastercam
- If Mastercam was restarted, click **Connect** to re-establish the connection

### Missing levels after connect
- The Levels tree is virtualized: only levels visible in Mastercam's panel are found
- Scroll the levels you need into view in Mastercam, then click **Refresh**
- Collapsed group children are not accessible via UIA until expanded

### First connect is slow (~11s)
- Normal: the app is discovering the Levels panel GUID for the first time
- Subsequent connects use the cached GUID and are faster (~4s)
- If slow every time, the GUID may have changed (Mastercam update) — the app auto-discovers and re-caches

## Development Tools

| Script | Purpose |
|--------|---------|
| `inspect_levels_panel.py` | Enumerate Mastercam child windows, find Levels panel candidates, dump UIA tree |
| `inspect_level_items.py` | Deep-inspect each TreeItem: controls, UIA patterns (Toggle, Invoke, Value), and hierarchy |
| `test_level_connect.py` | Quick smoke test: find panel → find level 202 → toggle visibility |

## Technical Notes

- **No mouse movement** — all operations use UIA `Invoke` pattern, not simulated clicks
- **Settings path** — `%LOCALAPPDATA%\LevelManager\settings.json` avoids Dropbox sync conflicts across machines
- **GUID caching** — each machine discovers and caches its own Levels panel GUID independently
- **Dark theme** — uses `sv_ttk` (Sun Valley theme) for modern dark appearance
- **Launcher** — `Level Manager.bat` tries common Python install paths, py launcher, then PATH
- **Virtualized tree** — Mastercam only creates UIA elements for visible items (~25% of tree at a time); `tree.children()` takes ~4.6s, per-level extraction ~150ms each
- **Toggle pattern unavailable** — `IsLevelVisibleButton` only supports `Invoke`, not `Toggle`; state read-back is not possible
