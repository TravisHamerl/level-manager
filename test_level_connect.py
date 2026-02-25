"""Quick test: can we find and connect to the Levels panel fast?"""
import sys
import ctypes
import ctypes.wintypes
from pywinauto import Application

EnumWindows = ctypes.windll.user32.EnumWindows
EnumChildWindows = ctypes.windll.user32.EnumChildWindows
GetClassName = ctypes.windll.user32.GetClassNameW
GetWindowText = ctypes.windll.user32.GetWindowTextW
IsWindowVisible = ctypes.windll.user32.IsWindowVisible
WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)

def gc(hwnd):
    b = ctypes.create_unicode_buffer(256)
    GetClassName(hwnd, b, 256)
    return b.value

def gt(hwnd):
    b = ctypes.create_unicode_buffer(256)
    GetWindowText(hwnd, b, 256)
    return b.value

# Step 1: Find Mastercam
print("Step 1: Finding Mastercam...")
mc = None
def f(hwnd, _):
    global mc
    t = gt(hwnd)
    if 'Mastercam' in t and IsWindowVisible(hwnd) and '.mcam' in t:
        mc = hwnd
    return True
EnumWindows(WNDENUMPROC(f), 0)
print(f"  Mastercam hwnd: {mc}")

if not mc:
    print("Not found!")
    sys.exit(1)

# Step 2: Find HwndWrapper for Levels
print("Step 2: Finding Levels HwndWrapper...")
levels = None
def f2(hwnd, _):
    global levels
    c = gc(hwnd)
    if '732c3493' in c:
        levels = hwnd
    return True
EnumChildWindows(mc, WNDENUMPROC(f2), 0)
print(f"  Levels hwnd: {levels}")

if not levels:
    print("Not found!")
    sys.exit(1)

# Step 3: Connect UIA
print("Step 3: Connecting UIA to Levels panel...")
sys.stdout.flush()
app = Application(backend='uia').connect(handle=levels)
win = app.window(handle=levels)
print(f"  Connected: {win.element_info.control_type}")

# Step 4: Find tree
print("Step 4: Finding LevelTreeListBox...")
sys.stdout.flush()
tree = win.child_window(auto_id="LevelTreeListBox", control_type="Tree")
print(f"  Tree exists: {tree.exists(timeout=5)}")

# Step 5: Find level 202
print("Step 5: Finding level 202...")
sys.stdout.flush()
items = tree.children(control_type="TreeItem")
print(f"  {len(items)} tree items")
for item in items:
    if "LevelTreeItem" not in item.window_text():
        continue
    edits = item.children(control_type="Edit")
    for edit in edits:
        try:
            v = edit.iface_value
            if v and v.CurrentValue == "202":
                print(f"  FOUND level 202!")
                # Find visibility button via children
                for child in item.children():
                    try:
                        aid = child.automation_id()
                        ct = child.element_info.control_type
                        print(f"    child: [{ct}] id={aid!r}")
                        if aid == "IsLevelVisibleButton":
                            print(f"  Found visibility button!")
                            child.iface_invoke.Invoke()
                            print(f"  TOGGLED!")
                            sys.exit(0)
                    except Exception as ex:
                        print(f"    child error: {ex}")
                print("  Visibility button not found in children!")
                sys.exit(1)
        except Exception:
            pass

print("Level 202 not found!")
