"""Mastercam Level Manager — toggle level visibility with GUI + hotkeys.

Usage:
    python tools/level_manager.py

Requires the Levels panel to be open/docked in Mastercam with groups expanded.
"""

import json
import os
import sys
import ctypes
import ctypes.wintypes
import threading
import tkinter as tk
from tkinter import ttk
from pathlib import Path
from pynput import keyboard
from pywinauto import Application
import sv_ttk

# ---------- Settings ----------

def _settings_path():
    """Store settings locally per machine, not in shared Dropbox folder."""
    local = os.environ.get("LOCALAPPDATA")
    if local:
        d = Path(local) / "LevelManager"
        d.mkdir(exist_ok=True)
        return d / "settings.json"
    # Fallback: next to script (non-Windows or missing env var)
    return Path(__file__).parent / ".level_manager_settings.json"

SETTINGS_FILE = _settings_path()


def load_settings():
    try:
        if SETTINGS_FILE.exists():
            return json.loads(SETTINGS_FILE.read_text())
    except Exception:
        pass
    return {}


def save_settings(settings):
    try:
        SETTINGS_FILE.write_text(json.dumps(settings, indent=2))
    except Exception:
        pass


# ---------- Mastercam connection (Win32 fast path) ----------

EnumWindows = ctypes.windll.user32.EnumWindows
EnumChildWindows = ctypes.windll.user32.EnumChildWindows
GetClassName = ctypes.windll.user32.GetClassNameW
GetWindowText = ctypes.windll.user32.GetWindowTextW
IsWindowVisible = ctypes.windll.user32.IsWindowVisible
WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)


def _get_class(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    GetClassName(hwnd, buf, 256)
    return buf.value


def _get_title(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    GetWindowText(hwnd, buf, 256)
    return buf.value


def _find_mastercam_hwnd():
    """Find the main Mastercam window handle via Win32 (fast)."""
    mc_hwnd = None
    def find_mc(hwnd, _):
        nonlocal mc_hwnd
        title = _get_title(hwnd)
        if 'Mastercam' in title and IsWindowVisible(hwnd) and '.mcam' in title.lower():
            mc_hwnd = hwnd
        return True
    EnumWindows(WNDENUMPROC(find_mc), 0)
    if not mc_hwnd:
        def find_mc2(hwnd, _):
            nonlocal mc_hwnd
            title = _get_title(hwnd)
            if 'Mastercam' in title and IsWindowVisible(hwnd):
                mc_hwnd = hwnd
            return True
        EnumWindows(WNDENUMPROC(find_mc2), 0)
    return mc_hwnd


def _find_levels_by_guid(mc_hwnd, guid_substr):
    """Search child windows for a HwndWrapper with the given GUID substring."""
    levels_hwnd = None
    def find_levels(hwnd, _):
        nonlocal levels_hwnd
        cls = _get_class(hwnd)
        if guid_substr in cls:
            levels_hwnd = hwnd
        return True
    EnumChildWindows(mc_hwnd, WNDENUMPROC(find_levels), 0)
    return levels_hwnd


# GUID cache — persisted to settings so only the first-ever connect is slow
_cached_guid = None


def connect_levels_panel():
    """Connect to Levels panel. Returns (panel_wrapper, tree, error_msg).

    Uses cached GUID for fast Win32 lookup. On first run (no cache),
    falls back to probing HwndWrapper windows via UIA and caches the
    discovered GUID for future connects.
    """
    global _cached_guid

    mc_hwnd = _find_mastercam_hwnd()
    if not mc_hwnd:
        return None, None, "Mastercam not found."

    # Load cached GUID from settings if we don't have it in memory
    if not _cached_guid:
        settings = load_settings()
        _cached_guid = settings.get("levels_guid")

    # Fast path: try cached GUID first
    if _cached_guid:
        levels_hwnd = _find_levels_by_guid(mc_hwnd, _cached_guid)
        if levels_hwnd:
            try:
                app = Application(backend='uia').connect(handle=levels_hwnd)
                panel = app.window(handle=levels_hwnd)
                tree = panel.child_window(auto_id="LevelTreeListBox", control_type="Tree")
                if tree.exists(timeout=0.5):
                    return panel, tree, None
            except Exception:
                pass
        # Cached GUID failed — clear it and fall through to probe
        _cached_guid = None

    # Slow path: probe all HwndWrapper children for one with LevelTreeListBox
    hwnd_wrappers = []
    def find_wrappers(hwnd, _):
        cls = _get_class(hwnd)
        if 'HwndWrapper' in cls and IsWindowVisible(hwnd):
            hwnd_wrappers.append((hwnd, cls))
        return True
    EnumChildWindows(mc_hwnd, WNDENUMPROC(find_wrappers), 0)

    for hw, cls in hwnd_wrappers:
        try:
            app = Application(backend='uia').connect(handle=hw)
            panel = app.window(handle=hw)
            tree = panel.child_window(auto_id="LevelTreeListBox", control_type="Tree")
            if tree.exists(timeout=1):
                # Extract GUID from class name and cache it
                # Class looks like "HwndWrapper[DefaultDomain;;ee99ee78-...]"
                for part in cls.split(';'):
                    part = part.strip(']').strip()
                    if len(part) >= 8 and '-' in part:
                        _cached_guid = part[:8]
                        break
                if _cached_guid:
                    settings = load_settings()
                    settings["levels_guid"] = _cached_guid
                    save_settings(settings)
                return panel, tree, None
        except Exception:
            pass

    return None, None, "Levels panel not found. Is it open in Mastercam?"


def scan_levels(tree):
    """Scan the tree and return a list of level dicts.

    Note: The Mastercam Levels tree is virtualized — only items currently
    visible in the viewport have UIA elements. Scroll the levels you need
    into view in Mastercam before connecting.
    """
    levels = []
    items = tree.children(control_type="TreeItem")

    for item in items:
        if "LevelTreeItem" not in item.window_text():
            continue
        try:
            # Single children() call — extract edits and button from same list
            all_children = item.children()
            number = None
            name = None
            vis_btn = None
            for child in all_children:
                try:
                    ct = child.element_info.control_type
                    if ct == "Edit":
                        val = child.iface_value
                        if val:
                            v = val.CurrentValue
                            if v and v.isdigit() and number is None:
                                number = v
                            elif v and not v.isdigit():
                                name = v
                    elif ct == "Button" and child.automation_id() == "IsLevelVisibleButton":
                        vis_btn = child
                except Exception:
                    pass

            if number is None:
                continue

            levels.append({
                "number": number,
                "name": name or f"Level {number}",
                "item": item,
                "vis_btn": vis_btn,
            })
        except Exception:
            pass

    return levels


def toggle_visibility(level):
    """Toggle a level's visibility. Returns True on success, False if stale.

    Re-finds the visibility button fresh from the cached TreeItem on each
    call (~150ms overhead). This avoids stale cached button wrappers that
    silently do nothing on Invoke() on some machines.
    """
    item = level.get("item")
    if not item:
        return False
    try:
        # Re-fetch children from the TreeItem — live UIA call.
        # If the item was destroyed (tree rebuilt), this throws or
        # returns empty, which we catch as stale.
        for child in item.children():
            try:
                if (child.element_info.control_type == "Button"
                        and child.automation_id() == "IsLevelVisibleButton"):
                    child.iface_invoke.Invoke()
                    return True
            except Exception:
                pass
        return False  # button not found → stale
    except Exception:
        return False


# ---------- Hotkey helpers ----------

def _is_ctrl(key):
    return key in (keyboard.Key.ctrl_l, keyboard.Key.ctrl_r, keyboard.Key.ctrl)

def _is_alt(key):
    return key in (keyboard.Key.alt_l, keyboard.Key.alt_r, keyboard.Key.alt, keyboard.Key.alt_gr)

def _is_shift(key):
    return key in (keyboard.Key.shift, keyboard.Key.shift_l, keyboard.Key.shift_r)

def _is_modifier(key):
    return _is_ctrl(key) or _is_alt(key) or _is_shift(key)


def hotkey_to_str(hk):
    """Convert hotkey dict to display string."""
    if not hk:
        return "(none)"
    parts = []
    for mod in hk.get("modifiers", []):
        parts.append(mod.capitalize())
    parts.append(hk.get("key", "?").upper())
    return "+".join(parts)


def hotkey_matches(hk, key, ctrl_held, alt_held, shift_held):
    """Check if a pressed key matches a hotkey config."""
    if not hk:
        return False
    mods = set(hk.get("modifiers", []))
    if ("ctrl" in mods) != ctrl_held:
        return False
    if ("alt" in mods) != alt_held:
        return False
    if ("shift" in mods) != shift_held:
        return False

    trigger = hk.get("key", "")
    # F-key
    if trigger.startswith("f") and trigger[1:].isdigit():
        return key == getattr(keyboard.Key, trigger, None)
    # Letter/char
    if hasattr(key, 'char') and key.char:
        return key.char.lower() == trigger.lower()
    if hasattr(key, 'vk') and key.vk and len(trigger) == 1:
        return key.vk == ord(trigger.upper())
    return False


# ---------- GUI ----------

class LevelManagerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Level Manager")
        self.root.minsize(650, 400)
        sv_ttk.set_theme("dark")
        default_font = ("Segoe UI", 11)
        self.root.option_add("*Font", default_font)

        self.levels = []
        self.panel = None
        self.tree = None
        self.hotkeys = {}  # level_number -> hotkey dict
        self.groups = {}   # group_name -> {"levels": [num, ...], "hotkey": {...} or None}
        self.row_widgets = {}  # level_number -> dict of widgets

        # Hotkey listener state
        self._ctrl_held = False
        self._alt_held = False
        self._shift_held = False
        self._listener = None

        # Recording state
        self._recording_target = None  # level_number or "grp:name"
        self._recording_mods = set()

        # Auto-rescan guard
        self._rescanning = False
        self._health_poll_id = None

        self._load_settings()
        self._build_ui()
        self._restore_geometry()
        self._start_hotkey_listener()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _load_settings(self):
        settings = load_settings()
        self.hotkeys = settings.get("hotkeys", {})
        self.groups = settings.get("groups", {})
        self._always_on_top = settings.get("always_on_top", True)
        self._geometry = settings.get("geometry", "700x450")

    def _save_settings(self):
        settings = load_settings()
        settings["hotkeys"] = self.hotkeys
        settings["groups"] = self.groups
        settings["always_on_top"] = self._always_on_top
        settings["geometry"] = self.root.geometry()
        save_settings(settings)

    def _restore_geometry(self):
        try:
            geo = self._geometry
            # Parse geometry string to validate bounds before applying
            # Format: WxH+X+Y or WxH-X-Y
            import re
            m = re.match(r'(\d+)x(\d+)([+-]\d+)([+-]\d+)', geo)
            if m:
                w, h, x, y = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
                sw = self.root.winfo_screenwidth()
                sh = self.root.winfo_screenheight()
                if x < -50 or y < -50 or x > sw - 50 or y > sh - 50:
                    geo = f"{w}x{h}+100+100"
            self.root.geometry(geo)
        except Exception:
            pass

    def _build_ui(self):
        self.root.attributes('-topmost', self._always_on_top)

        # Top bar
        top = ttk.Frame(self.root, padding=(8, 4))
        top.pack(fill=tk.X)

        self._connect_btn = ttk.Button(top, text="Connect", command=self._on_connect_click)
        self._connect_btn.pack(side=tk.LEFT)

        self._topmost_var = tk.BooleanVar(value=self._always_on_top)
        ttk.Checkbutton(top, text="Always on top", variable=self._topmost_var,
                        command=self._toggle_topmost).pack(side=tk.RIGHT)

        # Level list (Treeview) — extended select for multi-select grouping
        list_frame = ttk.Frame(self.root, padding=(8, 0, 8, 0))
        list_frame.pack(fill=tk.BOTH, expand=True)

        style = ttk.Style()
        style.configure("Treeview", rowheight=28, font=("Segoe UI", 11), foreground="#32CD32",
                        indent=20)
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"), foreground="#32CD32")
        style.configure("TButton", font=("Segoe UI", 11), foreground="#32CD32")
        style.configure("TCheckbutton", font=("Segoe UI", 11), foreground="#32CD32")
        style.configure("TLabel", font=("Segoe UI", 10), foreground="#32CD32")

        columns = ("number", "name", "hotkey")
        self.treeview = ttk.Treeview(list_frame, columns=columns, show="tree headings",
                                     selectmode="extended")
        self.treeview.heading("#0", text="", anchor=tk.W)
        self.treeview.heading("number", text="Level")
        self.treeview.heading("name", text="Name")
        self.treeview.heading("hotkey", text="Hotkey")
        self.treeview.column("#0", width=50, minwidth=50, stretch=False)
        self.treeview.column("number", width=60, minwidth=50, stretch=False)
        self.treeview.column("name", width=200, minwidth=100)
        self.treeview.column("hotkey", width=120, minwidth=80)

        # Tag for group header rows
        self.treeview.tag_configure("group", background="#2a3a4a", foreground="#32CD32", font=("Segoe UI", 11, "bold"))

        scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.treeview.yview)
        self.treeview.configure(yscrollcommand=scroll.set)

        self.treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Double-click to toggle visibility
        self.treeview.bind("<Double-1>", self._on_double_click)
        # Right-click to assign hotkey
        self.treeview.bind("<Button-3>", self._on_right_click)

        # Buttons row
        btn_frame = ttk.Frame(self.root, padding=(8, 4))
        btn_frame.pack(fill=tk.X)

        ttk.Button(btn_frame, text="Toggle", command=self._toggle_selected).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(btn_frame, text="Set Hotkey", command=self._set_hotkey_selected).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(btn_frame, text="Clear Hotkey", command=self._clear_hotkey_selected).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Separator(btn_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=6)
        ttk.Button(btn_frame, text="Group", command=self._create_group).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(btn_frame, text="Ungroup", command=self._ungroup_selected).pack(side=tk.LEFT)

        # Status bar
        self._status_var = tk.StringVar(value="Press Connect to get started.")
        status = ttk.Label(self.root, textvariable=self._status_var, padding=(8, 4),
                          font=("Segoe UI", 10))
        status.pack(fill=tk.X, side=tk.BOTTOM)

    def _toggle_topmost(self):
        self._always_on_top = self._topmost_var.get()
        self.root.attributes('-topmost', self._always_on_top)
        self._save_settings()

    def _set_status(self, msg):
        self._status_var.set(msg)

    def _notify_stale(self):
        """Show a prominent notification that levels need refreshing."""
        self._set_status("⚠ Connection stale — click Refresh to reconnect")
        self._connect_btn.configure(text="⟳ Refresh")
        # Flash the window to get attention
        try:
            self.root.attributes('-topmost', True)
            self.root.bell()
            if not self._always_on_top:
                self.root.after(500, lambda: self.root.attributes('-topmost', False))
        except Exception:
            pass

    # ---------- Connection ----------

    def _on_connect_click(self):
        self._connect_btn.configure(text="Refresh")
        self._connect_and_scan()

    def _connect_and_scan(self):
        # Fast path: try re-scanning existing tree connection first
        if self.tree:
            self._set_status("Refreshing levels...")
            self.root.update_idletasks()
            if self._rescan_levels():
                self._populate_treeview()
                self._set_status(f"Refreshed — {len(self.levels)} level(s) found")
                self._start_health_poll()
                return

        # Slow path: full reconnect
        self._set_status("Connecting to Mastercam Levels panel...")
        self.root.update_idletasks()

        panel, tree, err = connect_levels_panel()
        if err:
            self._set_status(f"Error: {err}")
            return

        self.panel = panel
        self.tree = tree
        old_levels = self.levels
        self.levels = scan_levels(tree)
        remapped = self._reconcile_numbers(old_levels)
        self._populate_treeview()
        if remapped:
            self._set_status(f"Connected — {len(self.levels)} level(s), remapped {remapped} hotkey(s)")
        else:
            self._set_status(f"Connected — {len(self.levels)} level(s) found")
        self._start_health_poll()

    def _reconcile_numbers(self, old_levels):
        """Remap hotkeys and groups when level numbers change.

        Compares old and new levels by name. If a name matches but the
        number changed, updates self.hotkeys keys and self.groups member
        lists to use the new number. Returns count of remapped entries.
        """
        if not old_levels:
            return 0

        # Build name→number maps, skipping duplicate names
        def _unique_name_map(levels):
            name_counts = {}
            for lvl in levels:
                name_counts[lvl["name"]] = name_counts.get(lvl["name"], 0) + 1
            return {lvl["name"]: lvl["number"] for lvl in levels
                    if name_counts.get(lvl["name"]) == 1}

        old_map = _unique_name_map(old_levels)   # name → old number
        new_map = _unique_name_map(self.levels)   # name → new number

        # Find renumbered levels: same name, different number
        remap = {}  # old_number → new_number
        for name, old_num in old_map.items():
            new_num = new_map.get(name)
            if new_num and new_num != old_num:
                remap[old_num] = new_num

        if not remap:
            return 0

        # Remap hotkeys dict keys
        for old_num, new_num in remap.items():
            if old_num in self.hotkeys:
                self.hotkeys[new_num] = self.hotkeys.pop(old_num)

        # Remap group member lists
        for grp in self.groups.values():
            members = grp.get("levels", [])
            grp["levels"] = [remap.get(n, n) for n in members]

        self._save_settings()
        return len(remap)

    def _rescan_levels(self):
        """Re-scan levels using existing tree connection (skips reconnect).

        Returns True if successful, False if the tree itself is stale.
        """
        if not self.tree:
            return False
        try:
            old_levels = self.levels
            new_levels = scan_levels(self.tree)
            if not new_levels:
                return False  # tree probably stale too
            self.levels = new_levels
            self._reconcile_numbers(old_levels)
            return True
        except Exception:
            return False

    def _auto_rescan_and_retry(self, level_number=None, group_numbers=None, group_name=None):
        """Auto-rescan in background thread, then retry the failed toggle."""
        if self._rescanning:
            return
        self._rescanning = True
        self._set_status("Auto-refreshing...")

        def _worker():
            try:
                success = self._rescan_levels()
            except Exception:
                success = False
            finally:
                self._rescanning = False

            if not success:
                self.root.after(0, self._notify_stale)
                return

            # Update UI and retry toggle on the tkinter thread
            def _finish():
                self._populate_treeview()
                if level_number:
                    self._hotkey_toggle(level_number, retry=False)
                elif group_numbers:
                    self._hotkey_toggle_group(group_numbers, group_name, retry=False)

            self.root.after(0, _finish)

        threading.Thread(target=_worker, daemon=True).start()

    def _start_health_poll(self):
        """Start periodic health check that detects stale wrappers proactively."""
        self._stop_health_poll()
        self._health_poll_id = self.root.after(15000, self._health_check)

    def _stop_health_poll(self):
        if self._health_poll_id:
            self.root.after_cancel(self._health_poll_id)
            self._health_poll_id = None

    def _health_check(self):
        """Probe one cached TreeItem to detect stale wrappers. ~150ms."""
        self._health_poll_id = None
        if not self.levels or self._rescanning:
            self._start_health_poll()
            return

        # Pick first level's TreeItem and try to find its button
        item = self.levels[0].get("item")
        if not item:
            self._start_health_poll()
            return

        def _probe():
            try:
                children = item.children()
                for ch in children:
                    try:
                        if (ch.element_info.control_type == "Button"
                                and ch.automation_id() == "IsLevelVisibleButton"):
                            return True  # still fresh
                    except Exception:
                        pass
                return False  # button not found → stale
            except Exception:
                return False  # item stale

        def _on_result(fresh):
            if not fresh and not self._rescanning:
                self._set_status("Detected change — refreshing...")
                self._rescanning = True
                def _rescan_worker():
                    try:
                        success = self._rescan_levels()
                    except Exception:
                        success = False
                    finally:
                        self._rescanning = False
                    def _update_ui():
                        if success:
                            self._populate_treeview()
                            self._set_status(f"Refreshed — {len(self.levels)} level(s)")
                    self.root.after(0, _update_ui)
                threading.Thread(target=_rescan_worker, daemon=True).start()
            self._start_health_poll()

        # Run probe in background to avoid blocking UI
        def _bg():
            fresh = _probe()
            self.root.after(0, lambda: _on_result(fresh))
        threading.Thread(target=_bg, daemon=True).start()

    def _populate_treeview(self):
        """Rebuild the treeview with groups and ungrouped levels."""
        self.treeview.delete(*self.treeview.get_children())

        # Tags for grouped member levels
        self.treeview.tag_configure("grouped", foreground="#32CD32")

        # Track which levels are in groups
        grouped_nums = set()
        for grp_name, grp in self.groups.items():
            grouped_nums.update(grp.get("levels", []))

        # Insert groups first
        for grp_name, grp in sorted(self.groups.items()):
            grp_iid = f"grp_{grp_name}"
            hk_str = hotkey_to_str(grp.get("hotkey"))
            member_nums = grp.get("levels", [])
            self.treeview.insert("", tk.END, iid=grp_iid, text="",
                                values=("", f"\u25A0 {grp_name}", hk_str),
                                tags=("group",), open=True)
            # Insert member levels as children with indent marker in Name column
            for i, num in enumerate(member_nums):
                name = f"Level {num}"
                for lvl in self.levels:
                    if lvl["number"] == num:
                        name = lvl["name"]
                        break
                lv_hk_str = hotkey_to_str(self.hotkeys.get(num))
                iid = f"lv_{num}"
                connector = "\u2514\u2500 " if i == len(member_nums) - 1 else "\u251C\u2500 "
                self.treeview.insert(grp_iid, tk.END, iid=iid, text="",
                                     values=(num, f"{connector}{name}", lv_hk_str),
                                     tags=("grouped",))

        # Insert ungrouped levels
        for lvl in self.levels:
            num = lvl["number"]
            if num in grouped_nums:
                continue
            hk_str = hotkey_to_str(self.hotkeys.get(num))
            iid = f"lv_{num}"
            self.treeview.insert("", tk.END, iid=iid, values=(num, lvl["name"], hk_str))

    # ---------- Level actions ----------

    def _iid(self, num):
        return f"lv_{num}"

    def _grp_iid(self, name):
        return f"grp_{name}"

    def _get_selected_level(self):
        """Get the first selected level (ignoring group headers)."""
        sel = self.treeview.selection()
        if not sel:
            return None
        iid = sel[0]
        if iid.startswith("grp_"):
            return None
        num = iid[3:] if iid.startswith("lv_") else iid
        for lvl in self.levels:
            if lvl["number"] == num:
                return lvl
        return None

    def _get_selected_target(self):
        """Get the first selected item as either a level number or group name.
        Returns (target_key, is_group) where target_key is level_number or group_name."""
        sel = self.treeview.selection()
        if not sel:
            return None, False
        iid = sel[0]
        if iid.startswith("grp_"):
            return iid[4:], True
        num = iid[3:] if iid.startswith("lv_") else iid
        return num, False

    def _get_selected_level_numbers(self):
        """Get all selected level numbers (resolving groups to their members)."""
        sel = self.treeview.selection()
        nums = []
        for iid in sel:
            if iid.startswith("grp_"):
                grp_name = iid[4:]
                grp = self.groups.get(grp_name, {})
                nums.extend(grp.get("levels", []))
            elif iid.startswith("lv_"):
                nums.append(iid[3:])
        return list(dict.fromkeys(nums))  # dedupe preserving order

    def _toggle_selected(self):
        nums = self._get_selected_level_numbers()
        if not nums:
            return
        toggled = []
        for num in nums:
            for lvl in self.levels:
                if lvl["number"] == num:
                    if toggle_visibility(lvl):
                        toggled.append(num)
                    break
        if toggled:
            self._set_status(f"Toggled {len(toggled)} level(s): {', '.join(toggled)}")
        else:
            self._set_status("Failed to toggle")

    def _on_double_click(self, event):
        region = self.treeview.identify_region(event.x, event.y)
        # Ignore double-click on the tree expand/collapse arrow
        if region == "tree":
            return
        if region == "cell":
            col = self.treeview.identify_column(event.x)
            if col == "#3":  # Hotkey column
                self._set_hotkey_selected()
                return
        self._toggle_selected()

    def _on_right_click(self, event):
        item = self.treeview.identify_row(event.y)
        if item:
            self.treeview.selection_set(item)
            self._set_hotkey_selected()

    # ---------- Group management ----------

    def _create_group(self):
        """Group the currently selected levels."""
        sel = self.treeview.selection()
        # Collect level numbers from selection (skip group headers, resolve children)
        nums = []
        for iid in sel:
            if iid.startswith("lv_"):
                nums.append(iid[3:])
            elif iid.startswith("grp_"):
                # Add the group's members
                grp = self.groups.get(iid[4:], {})
                nums.extend(grp.get("levels", []))
        nums = list(dict.fromkeys(nums))  # dedupe
        if len(nums) < 2:
            self._set_status("Select 2+ levels to group (Ctrl+click or Shift+click)")
            return

        # Generate group name from first level names
        names = []
        for num in nums[:3]:
            for lvl in self.levels:
                if lvl["number"] == num:
                    names.append(lvl["name"])
                    break
        grp_name = " + ".join(names)
        if len(nums) > 3:
            grp_name += f" +{len(nums) - 3}"

        # Check if any selected levels are already in another group — remove them
        for existing_name, existing_grp in list(self.groups.items()):
            remaining = [n for n in existing_grp.get("levels", []) if n not in nums]
            if len(remaining) < 2:
                # Group too small, dissolve it
                del self.groups[existing_name]
            elif remaining != existing_grp["levels"]:
                existing_grp["levels"] = remaining

        self.groups[grp_name] = {"levels": nums, "hotkey": None}
        self._save_settings()
        self._populate_treeview()
        self._set_status(f"Created group \"{grp_name}\" with {len(nums)} levels")

    def _ungroup_selected(self):
        """Dissolve the selected group."""
        target, is_group = self._get_selected_target()
        if not is_group:
            # If a level inside a group is selected, find its parent group
            sel = self.treeview.selection()
            if sel:
                parent = self.treeview.parent(sel[0])
                if parent and parent.startswith("grp_"):
                    target = parent[4:]
                    is_group = True
        if not is_group or target not in self.groups:
            self._set_status("Select a group to ungroup")
            return
        del self.groups[target]
        self._save_settings()
        self._populate_treeview()
        self._set_status(f"Ungrouped \"{target}\"")

    # ---------- Hotkey assignment ----------

    def _set_hotkey_selected(self):
        target, is_group = self._get_selected_target()
        if not target:
            self._set_status("Select a level or group first")
            return
        if is_group:
            self._start_recording(f"grp:{target}")
        else:
            self._start_recording(target)

    def _clear_hotkey_selected(self):
        target, is_group = self._get_selected_target()
        if not target:
            return
        if is_group:
            grp = self.groups.get(target)
            if grp:
                grp["hotkey"] = None
                self._save_settings()
                self.treeview.set(self._grp_iid(target), "hotkey", "(none)")
                self._set_status(f"Cleared hotkey for group \"{target}\"")
        else:
            if target in self.hotkeys:
                del self.hotkeys[target]
                self._save_settings()
                self.treeview.set(self._iid(target), "hotkey", "(none)")
                self._set_status(f"Cleared hotkey for level {target}")

    def _start_recording(self, target):
        """Start recording a hotkey. target is level_number or 'grp:name'."""
        self._stop_hotkey_listener()
        self._recording_target = target
        self._recording_mods = set()

        if target.startswith("grp:"):
            label = f"group \"{target[4:]}\""
            iid = self._grp_iid(target[4:])
        else:
            label = f"level {target}"
            iid = self._iid(target)
        self._set_status(f"Press a key combo for {label}... (Esc to cancel)")
        self.treeview.set(iid, "hotkey", "(press key...)")

        # Start a recording-only listener
        self._rec_listener = keyboard.Listener(
            on_press=self._on_rec_press,
            on_release=self._on_rec_release,
        )
        self._rec_listener.daemon = True
        self._rec_listener.start()

    def _on_rec_press(self, key):
        if _is_ctrl(key):
            self._recording_mods.add("ctrl")
        elif _is_alt(key):
            self._recording_mods.add("alt")
        elif _is_shift(key):
            self._recording_mods.add("shift")
        elif key == keyboard.Key.esc:
            self.root.after(0, self._cancel_recording)
            return False
        else:
            # Non-modifier key pressed — this is the trigger
            trigger = None
            if hasattr(key, 'char') and key.char:
                trigger = key.char.lower()
            elif hasattr(key, 'name'):
                trigger = key.name  # e.g., "f4", "space"
            elif hasattr(key, 'vk') and key.vk:
                if 65 <= key.vk <= 90:
                    trigger = chr(key.vk).lower()
                elif 48 <= key.vk <= 57:
                    trigger = chr(key.vk)

            if trigger:
                hk = {
                    "key": trigger,
                    "modifiers": sorted(self._recording_mods),
                }
                self.root.after(0, lambda: self._finish_recording(hk))
                return False

    def _on_rec_release(self, key):
        pass

    def _finish_recording(self, hk):
        target = self._recording_target
        if target is None:
            return
        self._recording_target = None
        if self._rec_listener:
            try:
                self._rec_listener.stop()
            except Exception:
                pass
            self._rec_listener = None

        is_group = target.startswith("grp:")

        # Check for conflicts against all level hotkeys and group hotkeys
        for existing_num, existing_hk in self.hotkeys.items():
            if not is_group and existing_num == target:
                continue
            if existing_hk == hk:
                self._set_status(f"Conflict: {hotkey_to_str(hk)} already assigned to level {existing_num}")
                self._restore_recording_display(target)
                self._start_hotkey_listener()
                return
        for grp_name, grp in self.groups.items():
            grp_hk = grp.get("hotkey")
            if is_group and grp_name == target[4:]:
                continue
            if grp_hk == hk:
                self._set_status(f"Conflict: {hotkey_to_str(hk)} already assigned to group \"{grp_name}\"")
                self._restore_recording_display(target)
                self._start_hotkey_listener()
                return

        # Assign the hotkey
        if is_group:
            grp_name = target[4:]
            self.groups[grp_name]["hotkey"] = hk
            self._save_settings()
            self.treeview.set(self._grp_iid(grp_name), "hotkey", hotkey_to_str(hk))
            self._set_status(f"Group \"{grp_name}\": hotkey set to {hotkey_to_str(hk)}")
        else:
            self.hotkeys[target] = hk
            self._save_settings()
            self.treeview.set(self._iid(target), "hotkey", hotkey_to_str(hk))
            self._set_status(f"Level {target}: hotkey set to {hotkey_to_str(hk)}")
        self._start_hotkey_listener()

    def _restore_recording_display(self, target):
        """Restore the hotkey column display after a cancelled/conflicted recording."""
        if target.startswith("grp:"):
            grp_name = target[4:]
            grp = self.groups.get(grp_name, {})
            self.treeview.set(self._grp_iid(grp_name), "hotkey", hotkey_to_str(grp.get("hotkey")))
        else:
            self.treeview.set(self._iid(target), "hotkey", hotkey_to_str(self.hotkeys.get(target)))

    def _cancel_recording(self):
        target = self._recording_target
        self._recording_target = None
        if self._rec_listener:
            try:
                self._rec_listener.stop()
            except Exception:
                pass
            self._rec_listener = None
        if target:
            self._restore_recording_display(target)
        self._set_status("Hotkey recording cancelled")
        self._start_hotkey_listener()

    # ---------- Global hotkey listener ----------

    def _start_hotkey_listener(self):
        if self._listener:
            return
        self._ctrl_held = False
        self._alt_held = False
        self._shift_held = False
        self._listener = keyboard.Listener(
            on_press=self._on_hotkey_press,
            on_release=self._on_hotkey_release,
        )
        self._listener.daemon = True
        self._listener.start()

    def _stop_hotkey_listener(self):
        if self._listener:
            try:
                self._listener.stop()
            except Exception:
                pass
            self._listener = None

    def _on_hotkey_press(self, key):
        if _is_ctrl(key):
            self._ctrl_held = True
        elif _is_alt(key):
            self._alt_held = True
        elif _is_shift(key):
            self._shift_held = True
        elif not _is_modifier(key):
            # Check group hotkeys first
            for grp_name, grp in self.groups.items():
                grp_hk = grp.get("hotkey")
                if grp_hk and hotkey_matches(grp_hk, key, self._ctrl_held, self._alt_held, self._shift_held):
                    nums = grp.get("levels", [])
                    self.root.after(0, lambda ns=nums, gn=grp_name: self._hotkey_toggle_group(ns, gn))
                    return
            # Check individual level hotkeys
            for num, hk in self.hotkeys.items():
                if hotkey_matches(hk, key, self._ctrl_held, self._alt_held, self._shift_held):
                    self.root.after(0, lambda n=num: self._hotkey_toggle(n))
                    break

    def _on_hotkey_release(self, key):
        if _is_ctrl(key):
            self._ctrl_held = False
        elif _is_alt(key):
            self._alt_held = False
        elif _is_shift(key):
            self._shift_held = False

    def _hotkey_toggle(self, level_number, retry=True):
        """Toggle a level by its number (called from hotkey)."""
        for lvl in self.levels:
            if lvl["number"] == level_number:
                if toggle_visibility(lvl):
                    self._set_status(f"Toggled level {level_number} ({lvl['name']})")
                elif retry:
                    self._auto_rescan_and_retry(level_number=level_number)
                else:
                    self._notify_stale()
                return
        self._set_status(f"Level {level_number} not found — try Refresh")

    def _hotkey_toggle_group(self, level_numbers, group_name, retry=True):
        """Toggle all levels in a group (called from hotkey)."""
        toggled = 0
        failed = False
        for num in level_numbers:
            for lvl in self.levels:
                if lvl["number"] == num:
                    if toggle_visibility(lvl):
                        toggled += 1
                    else:
                        failed = True
                    break
        if failed and toggled == 0:
            if retry:
                self._auto_rescan_and_retry(group_numbers=level_numbers, group_name=group_name)
            else:
                self._notify_stale()
        elif failed:
            self._set_status(f"Toggled group \"{group_name}\" ({toggled}/{len(level_numbers)} levels) — some stale, click Refresh")
        else:
            self._set_status(f"Toggled group \"{group_name}\" ({toggled}/{len(level_numbers)} levels)")

    # ---------- Lifecycle ----------

    def _on_close(self):
        self._save_settings()
        self._stop_health_poll()
        self._stop_hotkey_listener()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = LevelManagerApp()
    app.run()
