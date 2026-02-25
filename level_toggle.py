"""Toggle Mastercam level visibility with a global hotkey.

Usage:
    python tools/level_toggle.py              # Toggle level 202 with Ctrl+Shift+L
    python tools/level_toggle.py 210          # Toggle level 210
    python tools/level_toggle.py 202 --key f9 # Toggle level 202 with F9

Press Escape to quit.
"""

import sys
import argparse
import ctypes
import ctypes.wintypes
from pynput import keyboard
from pywinauto import Application, findwindows


# ---------- Fast panel discovery via Win32 ----------

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


def find_levels_hwnd():
    """Find the Levels panel HwndWrapper handle via Win32 (fast)."""
    # First find Mastercam
    mc_hwnd = None
    def find_mc(hwnd, _):
        nonlocal mc_hwnd
        title = _get_title(hwnd)
        if 'Mastercam' in title and IsWindowVisible(hwnd) and 'mcam' in title.lower():
            mc_hwnd = hwnd
        return True
    EnumWindows(WNDENUMPROC(find_mc), 0)

    if not mc_hwnd:
        # Broader search
        def find_mc2(hwnd, _):
            nonlocal mc_hwnd
            title = _get_title(hwnd)
            if 'Mastercam' in title and IsWindowVisible(hwnd):
                mc_hwnd = hwnd
            return True
        EnumWindows(WNDENUMPROC(find_mc2), 0)

    if not mc_hwnd:
        return None

    print(f"Mastercam: {_get_title(mc_hwnd)} (hwnd={mc_hwnd})")

    # Find the HwndWrapper with the known GUID for Levels panel
    levels_hwnd = None
    def find_levels(hwnd, _):
        nonlocal levels_hwnd
        cls = _get_class(hwnd)
        if '732c3493' in cls:
            levels_hwnd = hwnd
        return True
    EnumChildWindows(mc_hwnd, WNDENUMPROC(find_levels), 0)

    return levels_hwnd


# ---------- Level operations ----------

def connect_levels_panel():
    """Connect to the Levels panel via its HwndWrapper handle."""
    hwnd = find_levels_hwnd()
    if not hwnd:
        print("ERROR: Levels panel not found. Is it open/docked in Mastercam?")
        sys.exit(1)
    print(f"Levels panel hwnd: {hwnd}")
    app = Application(backend='uia').connect(handle=hwnd)
    win = app.window(handle=hwnd)
    return win


def find_level_tree(panel):
    """Find the LevelTreeListBox within the panel."""
    tree = panel.child_window(auto_id="LevelTreeListBox", control_type="Tree")
    if not tree.exists(timeout=5):
        print("ERROR: LevelTreeListBox not found in panel!")
        sys.exit(1)
    return tree


def find_level_item(tree, level_number):
    """Find a specific level by its number. Returns (TreeItem, name)."""
    level_str = str(level_number)
    items = tree.children(control_type="TreeItem")

    for item in items:
        if "LevelTreeItem" not in item.window_text():
            continue
        try:
            edits = item.children(control_type="Edit")
            for edit in edits:
                try:
                    val = edit.iface_value
                    if val and val.CurrentValue == level_str:
                        name = "?"
                        for e2 in edits:
                            try:
                                v2 = e2.iface_value
                                if v2 and v2.CurrentValue != level_str:
                                    name = v2.CurrentValue
                                    break
                            except Exception:
                                pass
                        return item, name
                except Exception:
                    pass
        except Exception:
            pass

    return None, None


def toggle_level_visibility(item):
    """Click the IsLevelVisibleButton on a level tree item."""
    for child in item.children():
        try:
            if child.automation_id() == "IsLevelVisibleButton":
                child.iface_invoke.Invoke()
                return
        except Exception:
            pass
    raise RuntimeError("IsLevelVisibleButton not found")


# ---------- Hotkey listener ----------

def _is_ctrl(key):
    return key in (keyboard.Key.ctrl_l, keyboard.Key.ctrl_r, keyboard.Key.ctrl)

def _is_alt(key):
    return key in (keyboard.Key.alt_l, keyboard.Key.alt_r, keyboard.Key.alt, keyboard.Key.alt_gr)

def _is_shift(key):
    return key in (keyboard.Key.shift, keyboard.Key.shift_l, keyboard.Key.shift_r)


def parse_hotkey(key_str):
    """Parse a hotkey string like 'ctrl+alt+i' or 'f9' into components."""
    parts = key_str.lower().split('+')
    need_ctrl = False
    need_alt = False
    need_shift = False
    trigger_char = None
    trigger_fkey = None

    for p in parts:
        p = p.strip()
        if p in ('ctrl', 'control'):
            need_ctrl = True
        elif p == 'shift':
            need_shift = True
        elif p == 'alt':
            need_alt = True
        elif p.startswith('f') and p[1:].isdigit():
            trigger_fkey = getattr(keyboard.Key, p)
        else:
            trigger_char = p

    return need_ctrl, need_alt, need_shift, trigger_char, trigger_fkey


def main():
    parser = argparse.ArgumentParser(description="Toggle Mastercam level visibility with a hotkey")
    parser.add_argument("level", nargs="?", default="202", help="Level number to toggle (default: 202)")
    parser.add_argument("--key", default="ctrl+alt+i", help="Hotkey combo (default: ctrl+alt+i)")
    args = parser.parse_args()

    level_num = args.level
    hotkey_str = args.key

    # Connect directly to the Levels panel (fast)
    print("Connecting to Levels panel...")
    panel = connect_levels_panel()
    tree = find_level_tree(panel)

    # Find the level
    item, name = find_level_item(tree, level_num)
    if item is None:
        print(f"ERROR: Level {level_num} not found in the tree!")
        print("Make sure the level's group is expanded in the Levels panel.")
        sys.exit(1)

    print(f"Found level {level_num}: {name!r}")

    # Parse hotkey
    need_ctrl, need_alt, need_shift, trigger_char, trigger_fkey = parse_hotkey(hotkey_str)
    print(f"Hotkey: {hotkey_str.upper()}")
    print(f"Press {hotkey_str.upper()} to toggle level {level_num} visibility")
    print(f"Press Escape to quit\n")

    # Track modifier state
    ctrl_held = False
    alt_held = False
    shift_held = False
    state = {"item": item, "panel": panel, "tree": tree}

    def do_toggle():
        try:
            toggle_level_visibility(state["item"])
            print(f"  Toggled level {level_num} ({name})")
        except Exception as e:
            print(f"  Error toggling: {e}")
            try:
                state["panel"] = connect_levels_panel()
                state["tree"] = find_level_tree(state["panel"])
                new_item, _ = find_level_item(state["tree"], level_num)
                if new_item:
                    state["item"] = new_item
                    toggle_level_visibility(new_item)
                    print(f"  Reconnected and toggled level {level_num}")
            except Exception as e2:
                print(f"  Reconnect failed: {e2}")

    def check_hotkey(key):
        nonlocal ctrl_held, alt_held, shift_held

        # Check modifiers
        if need_ctrl and not ctrl_held:
            return False
        if need_alt and not alt_held:
            return False
        if need_shift and not shift_held:
            return False

        # Check trigger key
        if trigger_fkey:
            return key == trigger_fkey
        if trigger_char:
            # Match by char or vk code
            if hasattr(key, 'char') and key.char:
                return key.char.lower() == trigger_char
            if hasattr(key, 'vk') and key.vk:
                # vk code for letters: A=65, B=66, etc.
                expected_vk = ord(trigger_char.upper())
                return key.vk == expected_vk
        return False

    def on_press(key):
        nonlocal ctrl_held, alt_held, shift_held

        if _is_ctrl(key):
            ctrl_held = True
        elif _is_alt(key):
            alt_held = True
        elif _is_shift(key):
            shift_held = True

        if key == keyboard.Key.esc:
            print("Exiting.")
            return False

        if check_hotkey(key):
            do_toggle()

    def on_release(key):
        nonlocal ctrl_held, alt_held, shift_held

        if _is_ctrl(key):
            ctrl_held = False
        elif _is_alt(key):
            alt_held = False
        elif _is_shift(key):
            shift_held = False

    with keyboard.Listener(on_press=on_press, on_release=on_release) as listener:
        listener.join()


if __name__ == "__main__":
    main()
