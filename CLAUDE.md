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

1. Open and dock the **Levels panel** in Mastercam (groups expanded)
2. Click **Connect** to discover the panel and scan levels
3. **Double-click** a level to toggle its visibility
4. **Right-click** or double-click the Hotkey column to assign a global hotkey
5. Use assigned hotkeys from anywhere (no need to focus the app)

### CLI Toggle
```bash
python level_toggle.py              # Toggle level 202 with Ctrl+Alt+I
python level_toggle.py 210          # Toggle level 210
python level_toggle.py 202 --key f9 # Toggle level 202 with F9
```
Press Escape to quit.

## Architecture

### Mastercam Connection

Two-phase discovery for speed:

1. **Win32 fast path** — `EnumWindows` + `EnumChildWindows` to find the Levels panel by its HwndWrapper GUID (`732c3493` in the class name)
2. **UIA fallback** — If GUID not found, enumerate all HwndWrapper windows and test each with UIA until finding one containing `LevelTreeListBox`

### Levels Panel Structure (UIA)
```
LevelTreeListBox (Tree, auto_id="LevelTreeListBox")
  ├─ GroupTreeItem (TreeItem — group header, expandable)
  │   ├─ LevelTreeItem (TreeItem)
  │   │   ├─ Edit (level number, e.g. "202")
  │   │   ├─ Edit (level name, e.g. "Geometry")
  │   │   └─ Button (auto_id="IsLevelVisibleButton" — Toggle + Invoke patterns)
  │   └─ ... more levels
  └─ LevelTreeItem (ungrouped levels at root)
```

### Visibility Toggle

Uses UIA `Invoke` pattern on `IsLevelVisibleButton` — no mouse movement required.

**Stale reference detection:** After `Invoke()`, the app reads the button's `Toggle` state before and after. If unchanged (stale COM reference silently did nothing), it uses a tiered retry: targeted re-lookup for just that level number (fast), then full reconnect only as a last resort.

**Nested levels:** Uses `tree.descendants()` (not `children()`) to find levels inside expanded groups.

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

Stored at `%LOCALAPPDATA%\LevelManager\settings.json` (not in Dropbox):
- Hotkey assignments (per level number)
- Group definitions and group hotkeys
- Window geometry

## Key Functions

### Connection (`level_manager.py`)

| Function | Description |
|----------|-------------|
| `find_levels_hwnd()` | Win32 fast discovery of Levels panel handle |
| `connect_levels_panel()` | UIA connect to panel; returns `(wrapper, error_msg)` |
| `find_level_tree(panel)` | Find `LevelTreeListBox` Tree control |
| `scan_levels(tree)` | Scan tree items, return list of `{number, name, item, vis_btn}` dicts |

### Toggle (`level_manager.py`)

| Function | Description |
|----------|-------------|
| `toggle_visibility(level)` | Invoke button + read-back verify; returns `True` only if state changed |
| `_ensure_fresh()` | Check tree child count; reconnect if structure changed |
| `_find_level_fresh(number)` | Targeted single-level re-lookup using existing tree ref |
| `_hotkey_toggle(number)` | Single level toggle with tiered retry (cache → re-lookup → reconnect) |
| `_hotkey_toggle_group(numbers, name)` | Group toggle with tiered retry and partial failure handling |

### Hotkey Helpers (`level_manager.py`)

| Function | Description |
|----------|-------------|
| `hotkey_to_str(hk)` | Convert hotkey dict to display string (e.g. "Ctrl+F9") |
| `hotkey_matches(hk, key, ctrl, alt, shift)` | Check if pressed key matches a hotkey config |

## Troubleshooting

### "Levels panel not found"
- Ensure the Levels panel is open/docked in Mastercam (not just the tab — the panel itself must be visible)
- Click **Connect** / **Refresh** to retry

### Hotkeys stop working
- The app auto-detects stale references and reconnects. If issues persist, click **Refresh**
- If Mastercam was restarted, click **Connect** to re-establish the connection

### Levels not appearing after connect
- Expand all groups in Mastercam's Levels panel before connecting
- Collapsed group children are not accessible via UIA until expanded

## Development Tools

| Script | Purpose |
|--------|---------|
| `inspect_levels_panel.py` | Enumerate Mastercam child windows, find Levels panel candidates, dump UIA tree |
| `inspect_level_items.py` | Deep-inspect each TreeItem: controls, UIA patterns (Toggle, Invoke, Value), and hierarchy |
| `test_level_connect.py` | Quick smoke test: find panel → find level 202 → toggle visibility |

## Technical Notes

- **No mouse movement** — all operations use UIA patterns (`Invoke`, `Toggle`, `SelectionItem`), not simulated clicks
- **Settings path** — `%LOCALAPPDATA%\LevelManager\settings.json` avoids Dropbox sync conflicts across machines
- **Dark theme** — uses `sv_ttk` (Sun Valley theme) for modern dark appearance
- **Launcher** — `Level Manager.bat` tries common Python install paths, py launcher, then PATH
