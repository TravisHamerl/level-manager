"""Inspect the Mastercam Levels panel to find automation targets.

Run with the Levels panel open/docked in Mastercam.
Uses win32 to enumerate child windows first (fast), then UIA for controls.
"""

import sys
import ctypes
import ctypes.wintypes
from pywinauto import Application

EnumWindows = ctypes.windll.user32.EnumWindows
EnumChildWindows = ctypes.windll.user32.EnumChildWindows
GetClassName = ctypes.windll.user32.GetClassNameW
GetWindowText = ctypes.windll.user32.GetWindowTextW
IsWindowVisible = ctypes.windll.user32.IsWindowVisible
GetParent = ctypes.windll.user32.GetParent
WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)


def get_class(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    GetClassName(hwnd, buf, 256)
    return buf.value


def get_title(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    GetWindowText(hwnd, buf, 256)
    return buf.value


def find_mastercam():
    """Find the main Mastercam window handle."""
    result = []
    def cb(hwnd, _):
        title = get_title(hwnd)
        if 'Mastercam' in title and IsWindowVisible(hwnd):
            result.append((hwnd, title))
        return True
    EnumWindows(WNDENUMPROC(cb), 0)
    return result


def list_all_children(parent_hwnd, max_count=500):
    """List all child windows of a parent."""
    children = []
    def cb(hwnd, _):
        if len(children) < max_count:
            cls = get_class(hwnd)
            title = get_title(hwnd)
            vis = IsWindowVisible(hwnd)
            children.append((hwnd, cls, title, vis))
        return True
    EnumChildWindows(parent_hwnd, WNDENUMPROC(cb), 0)
    return children


def dump_uia(hwnd, max_depth=4):
    """Connect via UIA and dump the control tree."""
    try:
        app = Application(backend='uia').connect(handle=hwnd)
        win = app.window(handle=hwnd)
        print(f"  UIA name: {win.window_text()!r}")
        print(f"  UIA type: {win.element_info.control_type}")
        print(f"  UIA rect: {win.rectangle()}")
        _dump(win, 0, max_depth)
    except Exception as e:
        print(f"  UIA error: {e}")


def _dump(ctrl, depth, max_depth):
    if depth > max_depth:
        return
    indent = "  " * (depth + 2)
    try:
        children = ctrl.children()
    except Exception:
        return
    for ch in children:
        try:
            name = ch.window_text()[:60]
            ct = ch.element_info.control_type
            aid = ch.automation_id()
            r = ch.rectangle()
            print(f"{indent}[{ct}] {name!r} id={aid!r} ({r.left},{r.top},{r.right},{r.bottom})")

            for pname, attr in [("TOGGLE", "iface_toggle"), ("INVOKE", "iface_invoke"),
                                ("EXPAND_COLLAPSE", "iface_expand_collapse")]:
                try:
                    iface = getattr(ch, attr)
                    if iface:
                        extra = ""
                        if pname == "TOGGLE":
                            extra = f" state={iface.CurrentToggleState}"
                        elif pname == "EXPAND_COLLAPSE":
                            s = {0: "Collapsed", 1: "Expanded", 2: "Partial", 3: "Leaf"}
                            extra = f" state={s.get(iface.CurrentExpandCollapseState, '?')}"
                        print(f"{indent}  ^ {pname}{extra}")
                except Exception:
                    pass

            _dump(ch, depth + 1, max_depth)
        except Exception:
            pass


if __name__ == "__main__":
    print("=== Finding Mastercam ===\n")
    mc_windows = find_mastercam()
    if not mc_windows:
        print("Mastercam not found!")
        sys.exit(1)

    for hwnd, title in mc_windows:
        print(f"  {hwnd}: {title!r}")

    mc_hwnd = mc_windows[0][0]

    print(f"\n=== Enumerating child windows of Mastercam (hwnd={mc_hwnd}) ===\n")
    children = list_all_children(mc_hwnd)
    print(f"Found {len(children)} child windows.\n")

    # Show interesting ones (non-empty title or known class patterns)
    interesting = []
    for hwnd, cls, title, vis in children:
        if title or 'Hwnd' in cls or 'Wpf' in cls.lower() or 'Avalon' in cls:
            print(f"  hwnd={hwnd}  vis={vis}  class={cls[:60]!r}  title={title[:40]!r}")
            interesting.append((hwnd, cls, title, vis))

    # Also show class distribution
    class_counts = {}
    for _, cls, _, _ in children:
        class_counts[cls] = class_counts.get(cls, 0) + 1
    print(f"\nClass distribution ({len(class_counts)} unique classes):")
    for cls, count in sorted(class_counts.items(), key=lambda x: -x[1])[:20]:
        print(f"  {count:3d}x  {cls}")

    # Try to find anything with "level" in it
    print("\n=== Searching for 'level' in titles/classes ===\n")
    for hwnd, cls, title, vis in children:
        if 'level' in title.lower() or 'level' in cls.lower():
            print(f"  MATCH: hwnd={hwnd} class={cls!r} title={title!r} vis={vis}")

    # Try UIA on the most likely candidate (HwndWrapper or titled windows)
    print("\n=== UIA inspection of candidates ===\n")
    inspected = 0
    for hwnd, cls, title, vis in interesting:
        if not vis:
            continue
        if inspected >= 5:
            print("  (stopping after 5 inspections)")
            break
        print(f"\n--- hwnd={hwnd} class={cls[:40]!r} title={title!r} ---")
        dump_uia(hwnd, max_depth=2)
        inspected += 1

    print("\n=== Done ===")
