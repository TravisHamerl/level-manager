"""Deep inspect the Levels tree items to find visibility toggles.

Drills into each TreeItem in the LevelTreeListBox to find:
- The display text (level name/number)
- Visibility toggle buttons
- Any other interactive controls
"""

import sys
from pywinauto import Application, findwindows


def connect_levels_panel():
    """Connect to the Levels panel - try multiple strategies."""
    # Strategy 1: By title "Levels"
    for cls in ["#32770", None]:
        try:
            kwargs = {"title": "Levels"}
            if cls:
                kwargs["class_name"] = cls
            elements = findwindows.find_elements(**kwargs)
            if elements:
                el = elements[0]
                print(f"Found Levels panel: hwnd={el.handle} class={el.class_name}")
                app = Application(backend='uia').connect(handle=el.handle)
                return app.window(handle=el.handle)
        except Exception as e:
            print(f"  Strategy failed: {e}")

    # Strategy 2: Find via Mastercam main window
    print("Trying via Mastercam main window...")
    mc_elements = findwindows.find_elements(title_re='.*Mastercam.*Router.*')
    if not mc_elements:
        mc_elements = findwindows.find_elements(title_re='.*Mastercam.*')
    if mc_elements:
        app = Application(backend='uia').connect(handle=mc_elements[0].handle)
        win = app.window(handle=mc_elements[0].handle)
        try:
            # Search for LevelTreeListBox directly
            tree = win.child_window(auto_id="LevelTreeListBox", control_type="Tree")
            if tree.exists(timeout=3):
                print(f"Found LevelTreeListBox directly")
                return tree.parent().parent()  # return the panel
        except Exception as e:
            print(f"  Direct tree search failed: {e}")

        # Try the HwndWrapper
        try:
            wrapper = win.child_window(class_name_re="HwndWrapper.*732c3493.*")
            if wrapper.exists(timeout=3):
                print(f"Found via HwndWrapper class")
                return wrapper
        except Exception as e:
            print(f"  HwndWrapper search failed: {e}")

    print("ERROR: Could not find Levels panel. Is it open in Mastercam?")
    sys.exit(1)


def dump_deep(ctrl, depth=0, max_depth=6):
    """Recursively dump everything about a control."""
    indent = "  " * depth
    try:
        name = ctrl.window_text()[:80]
        ct = ctrl.element_info.control_type
        aid = ctrl.automation_id()
        r = ctrl.rectangle()
        cls = ctrl.element_info.class_name or ""

        print(f"{indent}[{ct}] name={name!r} id={aid!r} class={cls[:30]!r} ({r.left},{r.top},{r.right},{r.bottom})")

        # Check all known UIA patterns
        for pname, attr in [
            ("Toggle", "iface_toggle"),
            ("Invoke", "iface_invoke"),
            ("SelectionItem", "iface_selection_item"),
            ("ExpandCollapse", "iface_expand_collapse"),
            ("Value", "iface_value"),
        ]:
            try:
                iface = getattr(ctrl, attr)
                if iface:
                    extra = ""
                    if pname == "Toggle":
                        extra = f" state={iface.CurrentToggleState}"
                    elif pname == "SelectionItem":
                        extra = f" selected={iface.CurrentIsSelected}"
                    elif pname == "ExpandCollapse":
                        s = {0: "Collapsed", 1: "Expanded", 2: "Partial", 3: "Leaf"}
                        extra = f" state={s.get(iface.CurrentExpandCollapseState, '?')}"
                    elif pname == "Value":
                        try:
                            extra = f" value={iface.CurrentValue!r}"
                        except Exception:
                            pass
                    print(f"{indent}  ^ {pname}{extra}")
            except Exception:
                pass

    except Exception as e:
        print(f"{indent}[error: {e}]")
        return

    if depth >= max_depth:
        print(f"{indent}  (max depth reached)")
        return

    try:
        for child in ctrl.children():
            dump_deep(child, depth + 1, max_depth)
    except Exception:
        pass


if __name__ == "__main__":
    # Connect to Mastercam and find the tree directly
    mc_elements = findwindows.find_elements(title_re='.*Mastercam.*Router.*')
    if not mc_elements:
        mc_elements = findwindows.find_elements(title_re='.*Mastercam.*')
    if not mc_elements:
        print("ERROR: Mastercam not found!")
        sys.exit(1)

    app = Application(backend='uia').connect(handle=mc_elements[0].handle)
    win = app.window(handle=mc_elements[0].handle)
    print(f"Connected to: {mc_elements[0].name}")

    tree = win.child_window(auto_id="LevelTreeListBox", control_type="Tree")
    print(f"Found tree: {tree.element_info.control_type}\n")

    items = tree.children(control_type="TreeItem")
    print(f"Found {len(items)} tree items\n")

    for i, item in enumerate(items):
        name = item.window_text()[:60]
        # Only deep-inspect LevelTreeItemViewModels and the first group
        is_level = "LevelTreeItem" in name
        is_group = "GroupTreeItem" in name

        print(f"\n{'='*60}")
        print(f"ITEM {i}: {name!r}")
        print(f"{'='*60}")

        # Deep dump — go deeper for levels, shallow for groups
        depth = 6 if is_level else 3
        dump_deep(item, depth=0, max_depth=depth)

        # Stop after we've seen a couple levels in detail
        if is_level and i > 10:
            print("\n(stopping after enough level items)")
            break

    print("\n=== Done ===")
