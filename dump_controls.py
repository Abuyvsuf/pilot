# Writes every control of the open Nexus Pro windows to nexus_controls.txt (copy-safe: no asterisks)
# Use: open Nexus on the screen you want to map (e.g. Patients), then run: python dump_controls.py
from pywinauto import Desktop

NEXUS_TITLE = "Nexus Pro"


def line(c, depth):
    try:
        info = c.element_info
        r = c.rectangle()
        return "%s%s | name='%s' | class='%s' | auto_id='%s' | rect=(%d,%d,%d,%d)" % (
            "  " if depth else "", info.control_type, c.window_text(), info.class_name,
            info.automation_id, r.left, r.top, r.right, r.bottom)
    except Exception as e:
        return "  ? (%s)" % e


def main():
    pids = set()
    for w in Desktop(backend="uia").windows():
        try:
            if NEXUS_TITLE.lower() in (w.window_text() or "").lower():
                pids.add(w.process_id())
        except Exception:
            pass
    if not pids:
        print("No window title contains '%s'" % NEXUS_TITLE)
        return
    with open("nexus_controls.txt", "w", encoding="utf-8") as f:
        for w in Desktop(backend="uia").windows():
            try:
                if w.process_id() not in pids or not w.is_visible():
                    continue
            except Exception:
                continue
            f.write("=== WINDOW " + line(w, 0) + "\n")
            for c in w.descendants():
                try:
                    if c.is_visible():
                        f.write(line(c, 1) + "\n")
                except Exception:
                    pass
    print("Saved nexus_controls.txt")


if __name__ == "__main__":
    main()
