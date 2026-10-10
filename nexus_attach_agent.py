# Nexus Pro request-form auto-attach agent (copy-safe version: no asterisks)
# Install once: pip install watchdog pywinauto pyautogui requests pyzbar pymupdf pillow
#
# For every scan that lands in Z:\ (example: 224601091026CSF_0007.jpg):
#   1. The first 12 digits give the case 224601-09-10 and the date 09-Oct-2026.
#   2. Nexus Pro: Patients > Date = 09-Oct-2026 > Center = 0 > Case # = 224601-09-10 > Search.
#   3. Select the record, right-click it > EMR prescription > Upload file.
#   4. Windows file dialog: the full path of the scan > Open.
#   5. Back in Nexus: Select file.
#   6. The result goes to n8n (Telegram alert); the scan moves to Z:\done, Z:\failed or Z:\check.
# Every step is checked. When something does not match, the agent stops and reports
# instead of guessing, so a form is never attached to the wrong record or twice.
#   status "success"   = attached
#   status "failed"    = stopped BEFORE "Select file" was clicked, nothing attached (safe to retry)
#   status "uncertain" = stopped AFTER "Select file" was clicked: check the record in Nexus by hand
#   status "skipped"   = this scan was already attached (or attempted) earlier
import ctypes
import datetime
import logging
import operator
import queue
import re
import shutil
import threading
import time
from pathlib import Path

import requests
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

# ------------------------- CONFIG -------------------------
WATCH_DIR = Path("Z:/")
N8N_WEBHOOK = "https://abvyusuf.app.n8n.cloud/webhook/nexus-attach-result"
ALLOWED_EXT = {".pdf", ".png", ".jpg", ".jpeg", ".tif", ".tiff"}
PROCESS_EXISTING_ON_START = True  # also handle scans already waiting in Z:\
RESCAN_SECONDS = 60  # re-check Z:\ this often, in case the folder watcher missed something
USE_BARCODE = True  # if the form has a readable barcode, it must match the filename

NEXUS_TITLE = "Nexus Pro"
CENTER_VALUE = "0"
# Programs that are never Nexus, even if their window title mentions "Nexus Pro"
NOT_NEXUS_EXE = {"chrome.exe", "msedge.exe", "firefox.exe", "opera.exe", "brave.exe",
                 "iexplore.exe", "telegram.exe", "whatsapp.exe", "explorer.exe", "notepad.exe",
                 "notepad++.exe", "code.exe", "python.exe", "pythonw.exe", "cmd.exe",
                 "windowsterminal.exe", "winword.exe", "excel.exe", "outlook.exe", "teams.exe"}

# Texts on the Nexus screens (regular expressions, not case sensitive)
PATIENTS_TITLE = r"Patients"
DATE_LABEL = r"^\s{0,}Date\s{0,}:?\s{0,}$"
CENTER_LABEL = r"^\s{0,}Cent(er|re)\s{0,}:?\s{0,}$"
CASE_LABEL = r"^\s{0,}Case\s{0,}#\s{0,}:?\s{0,}$"
CASE_LABEL_LOOSE = r"Case\s{0,}#"
SEARCH_TITLE = r"^\s{0,}&?Search\s{0,}$"
EMR_TITLE = r"^\s{0,}&?EMR\s{0,}prescription\s{0,}(\.\.\.)?\s{0,}(\t.{0,})?$"
UPLOAD_PRESCRIPTION_TITLE = r"^\s{0,}&?Upload\s{0,}prescription\s{0,}(\.\.\.)?\s{0,}$"
UPLOAD_FILE_TITLE = r"^\s{0,}&?Upload\s{0,}file\s{0,}(\.\.\.)?\s{0,}$"
SELECT_FILE_TITLE = r"^\s{0,}&?Select\s{0,}file\s{0,}(\.\.\.)?\s{0,}$"
ERROR_WORDS = r"error|fail|invalid|unable|cannot|can't|not found|denied|exception|already"
QUESTION_BUTTONS = r"^\s{0,}&?(Cancel|No)\s{0,}$"
# ----------------------------------------------------------

DONE_DIR = WATCH_DIR / "done"
FAILED_DIR = WATCH_DIR / "failed"
CHECK_DIR = WATCH_DIR / "check"
# Kept on the local disk next to this script, so a network drop cannot lose it.
LEDGER = Path(__file__).resolve().with_name("attached_files.txt")

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Control types that are never a field label / input box / search result.
LABEL_SKIP = ["HeaderItem", "Header", "DataItem", "ListItem", "TreeItem", "MenuItem",
              "TabItem", "Button", "Hyperlink", "TitleBar", "Window"]
INPUT_SKIP = ["Text", "Group", "Window", "TitleBar", "MenuBar", "Menu", "MenuItem",
              "ToolBar", "StatusBar", "Header", "HeaderItem", "Image", "ScrollBar",
              "Thumb", "Separator", "Hyperlink", "Tab", "TabItem", "Table", "DataGrid",
              "List", "Tree", "TreeItem", "ListItem", "DataItem", "Button"]
RECORD_SKIP = ["Window", "TitleBar", "StatusBar", "MenuBar", "Menu", "MenuItem",
               "Button", "TabItem", "ComboBox", "Spinner", "ToolBar", "Header", "HeaderItem"]
ROW_TYPES = ["DataItem", "ListItem", "TreeItem", "Custom"]
RECORD_RANK = {"DataItem": 0, "ListItem": 0, "TreeItem": 0, "Custom": 1, "Text": 2, "Edit": 3}
CLICKABLE = ["Button", "SplitButton", "MenuItem", "ListItem", "Hyperlink"]
BUTTON_SKIP = ["TitleBar", "Window", "Edit", "Text", "HeaderItem"]


class UncertainAttach(RuntimeError):
    # Raised for any problem after "Select file" was clicked: the form may already be attached.
    pass


# ------------------------- case number -------------------------

def nexus_date_text(d):
    # 2026-10-09 becomes 09-Oct-2026 (same as Nexus shows it, independent of Windows language)
    return "%02d-%s-%04d" % (d.day, MONTHS[d.month - 1], d.year)


def parse_case(raw):
    # Only the first 12 digits count; anything after them is ignored.
    # 224601091026CSF_0007 becomes ("224601-09-10", date 2026-10-09)
    m = re.match(r"\D{0,}(\d{6})(\d{2})(\d{2})(\d{2})", raw or "")
    if not m:
        raise ValueError("No 12-digit case number in '%s'" % raw)
    number, day, month, year = m.groups()
    try:
        case_date = datetime.date(2000 + int(year), int(month), int(day))
    except ValueError:
        raise ValueError("Case number in '%s' has an impossible date (day %s, month %s)"
                         % (raw, day, month)) from None
    if case_date > datetime.date.today() + datetime.timedelta(days=1):
        raise ValueError("Case number in '%s' has a future date (%s)"
                         % (raw, nexus_date_text(case_date)))
    return "%s-%s-%s" % (number, day, month), case_date


def format_case_number(raw):
    return parse_case(raw)[0]


def parse_date_text(text):
    # Understands 09-Oct-2026, 9 Oct 2026, 09/10/2026 (day first). Returns None if no date.
    text = text or ""
    try:
        m = re.search(r"(\d{1,2})[-/ .]{1,2}([A-Za-z]{3,9})[-/ .,]{1,2}(\d{4})", text)
        if m:
            mon = m.group(2)[:3].title()
            if mon in MONTHS:
                return datetime.date(int(m.group(3)), MONTHS.index(mon) + 1, int(m.group(1)))
        m = re.search(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})", text)
        if m:
            return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None
    return None


def read_barcode(path):
    try:
        from PIL import Image
        from pyzbar.pyzbar import decode
        if path.suffix.lower() == ".pdf":
            import fitz
            page = fitz.open(path)[0]
            pix = page.get_pixmap(dpi=200)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        else:
            img = Image.open(path)
        res = decode(img)
        if res:
            return res[0].data.decode()
        return None
    except Exception as e:
        logging.warning("Barcode read failed for %s: %s", path.name, e)
        return None


def case_from_file(path):
    file_case = file_date = file_err = None
    try:
        file_case, file_date = parse_case(path.stem)
    except ValueError as e:
        file_err = str(e)
    bc_case = bc_date = None
    if USE_BARCODE:
        code = read_barcode(path)
        if code:
            try:
                bc_case, bc_date = parse_case(code)
            except ValueError:
                logging.info("Barcode '%s' is not a case number; using the filename", code)
    if file_case and bc_case and (file_case, file_date) != (bc_case, bc_date):
        raise RuntimeError("Barcode says %s (%s) but filename says %s (%s) - check this form by hand"
                           % (bc_case, nexus_date_text(bc_date), file_case, nexus_date_text(file_date)))
    if file_case:
        return file_case, file_date
    if bc_case:
        return bc_case, bc_date
    raise RuntimeError(file_err or "No case number found")


# ------------------------- UI helpers -------------------------

class Run:
    # The Nexus window being driven for one scan.
    def __init__(self, win):
        self.win = win
        self.pid = win.process_id()
        self.main = win.handle


def safe(fn, default=None):
    try:
        return fn()
    except Exception:
        return default


def is_visible(w):
    try:
        return w.is_visible()
    except Exception:
        return False


def area(w):
    try:
        r = w.rectangle()
        return operator.mul(max(r.width(), 0), max(r.height(), 0))
    except Exception:
        return 0


def ctype(c):
    return safe(lambda: c.element_info.control_type, "") or ""


def name_of(c):
    return (safe(c.window_text, "") or "").strip()


def rect(c):
    return safe(lambda: c.rectangle())


def rect_key(c):
    r = rect(c)
    return (r.left, r.top, r.right, r.bottom) if r is not None else None


def top_left(c):
    r = rect(c)
    return (r.top, r.left) if r is not None else (99999, 99999)


def overlaps(a, b):
    return not (a.right < b.left or a.left > b.right or a.bottom < b.top or a.top > b.bottom)


def wait_for(fn, timeout, interval=0.5):
    end = time.time() + timeout
    while True:
        res = safe(fn)
        if res:
            return res
        if time.time() >= end:
            return res
        time.sleep(interval)


def exe_of(pid):
    from pywinauto.application import process_module
    return Path(safe(lambda: process_module(pid), "") or "").name.lower()


def visible_titles():
    from pywinauto import Desktop
    titles = []
    for w in Desktop(backend="uia").windows():
        try:
            if is_visible(w) and w.window_text():
                titles.append(w.window_text())
        except Exception:
            pass
    return titles


def find_main_window(timeout=10):
    # Several windows can be titled Nexus Pro. Take the biggest visible one of a non-browser program.
    from pywinauto import Desktop
    end = time.time() + timeout
    while time.time() < end:
        cands = []
        for w in Desktop(backend="uia").windows():
            try:
                title = w.window_text() or ""
                cls = (w.class_name() or "").lower()
            except Exception:
                continue
            if NEXUS_TITLE.lower() not in title.lower():
                continue
            if "console" in cls or "cascadia" in cls or "terminal" in cls:
                continue
            if exe_of(safe(w.process_id)) in NOT_NEXUS_EXE:
                continue
            if is_visible(w):
                cands.append(w)
        if cands:
            if len({safe(c.process_id) for c in cands}) > 1:
                logging.error("More than one program has a '%s' window: %s",
                              NEXUS_TITLE, [c.window_text() for c in cands])
                return None
            cands.sort(key=area, reverse=True)
            return cands[0]
        time.sleep(0.5)
    logging.error("No window title contains '%s'. Visible windows: %s",
                  NEXUS_TITLE, visible_titles())
    return None


def nexus_windows(pid):
    # Real top-level windows of Nexus, including OWNED ones (EMR form, file dialog, message boxes).
    # The UIA desktop lists only un-owned windows, so enumerate with win32 and wrap them for UIA.
    from pywinauto import findwindows
    from pywinauto.controls.uiawrapper import UIAWrapper
    from pywinauto.uia_element_info import UIAElementInfo
    found = []
    handles = safe(lambda: findwindows.find_windows(backend="win32", process=pid, visible_only=True), []) or []
    for h in handles:
        w = safe(lambda: UIAWrapper(UIAElementInfo(h)))
        if w is not None:
            found.append(w)
    return found


def window_class(w):
    return safe(lambda: w.element_info.class_name, "") or ""


def nexus_dialogs(pid):
    return [w for w in nexus_windows(pid) if window_class(w) == "#32770"]


def new_windows(run, before):
    # Top-level Nexus windows that appeared after the "before" snapshot (menus and tooltips excluded).
    return [w for w in nexus_windows(run.pid)
            if safe(lambda: w.handle) not in before and safe(lambda: w.handle) != run.main
            and window_class(w) not in ("#32768", "tooltips_class32")]


def handles_of(windows):
    return {safe(lambda: w.handle) for w in windows}


def all_roots(run):
    # The main window plus every other visible top-level window of the Nexus process.
    # Other programs are never searched.
    return [run.win] + [w for w in nexus_windows(run.pid) if safe(lambda: w.handle) != run.main]


def dialog_text(d):
    parts = []
    for c in safe(d.descendants, []):
        if ctype(c) == "Text":
            t = name_of(c)
            if t and t not in parts:
                parts.append(t)
    return " ".join(parts)[:300] or name_of(d)


def button_names(d):
    return [name_of(b) for b in safe(lambda: d.descendants(control_type="Button"), [])]


def close_window(w):
    safe(w.close)
    time.sleep(0.5)


def foreground():
    user32 = ctypes.windll.user32
    user32.GetForegroundWindow.restype = ctypes.c_void_p
    return user32.GetForegroundWindow() or 0


def ga_root(h):
    # The top-level window that contains window h.
    user32 = ctypes.windll.user32
    user32.GetAncestor.restype = ctypes.c_void_p
    user32.GetAncestor.argtypes = [ctypes.c_void_p, ctypes.c_uint]
    return user32.GetAncestor(h, 2) or 0


def root_at(x, y):
    # The top-level window that is really on screen at (x, y).
    import ctypes.wintypes
    user32 = ctypes.windll.user32
    user32.WindowFromPoint.restype = ctypes.c_void_p
    user32.WindowFromPoint.argtypes = [ctypes.wintypes.POINT]
    h = user32.WindowFromPoint(ctypes.wintypes.POINT(x, y))
    return ga_root(h) if h else 0


def root_of(c):
    # The top-level window that holds control c (0 if unknown).
    p = c
    for _ in range(50):
        h = safe(lambda: p.handle)
        if h:
            return ga_root(h)
        p = safe(p.parent)
        if p is None:
            return 0
    return 0


def window_pid(h):
    from pywinauto import handleprops
    return safe(lambda: handleprops.processid(h))


def assert_focus(run, expected=0):
    # Keys are only sent while the window we clicked into (or at least Nexus) is the active window.
    h = foreground()
    ok = bool(h) and (h == expected if expected else window_pid(h) == run.pid)
    if not ok:
        from pywinauto import handleprops
        raise RuntimeError("The active window is '%s', not the expected Nexus window; "
                           "stopped so nothing is typed into it" % safe(lambda: handleprops.text(h), "?"))


def click_at(run, x, y, expect_root=0, button="left"):
    # Clicks only when the spot shows the expected Nexus window (nothing else on top of it).
    # Returns the top-level window that was clicked.
    import pyautogui
    root = root_at(x, y)
    if not root or window_pid(root) != run.pid or (expect_root and root != expect_root):
        raise RuntimeError("Another window covers the spot to click (%d,%d); stopped" % (x, y))
    pyautogui.click(x, y, button=button)
    return root


def click(run, c, button="left"):
    r = rect(c)
    if r is None:
        raise RuntimeError("'%s' disappeared before it could be clicked" % name_of(c))
    return click_at(run, (r.left + r.right) // 2, (r.top + r.bottom) // 2, root_of(c), button)


def find_all(parents, pattern, skip=(), only=None):
    # Visible controls whose name matches; the same control reached twice is listed once.
    pat = re.compile(pattern, re.IGNORECASE)
    found = []
    keys = set()
    for parent in parents:
        for c in safe(parent.descendants, []):
            try:
                name = c.window_text() or ""
                if not name or not pat.search(name):
                    continue
                t = c.element_info.control_type
                if t in skip or (only and t not in only):
                    continue
                if not is_visible(c):
                    continue
                k = (t, name, rect_key(c))
                if k in keys:
                    continue
                keys.add(k)
                found.append(c)
            except Exception:
                continue
    return found


def find_one(parents, pattern, prefer=(), skip=(), timeout=5, sort_key=None):
    # Waits for a visible match. Only the preferred control types are accepted while waiting;
    # any other type is accepted only at the very end, as a last resort.
    # parents can be a list or a function returning a list (re-read on every try).
    sort_key = sort_key or top_left
    end = time.time() + timeout
    while True:
        last = time.time() >= end
        roots = parents() if callable(parents) else parents
        only = None if (last or not prefer) else list(prefer)
        cands = find_all(roots, pattern, skip, only)
        if cands:
            cands.sort(key=lambda c: (0 if ctype(c) in prefer else 1,) + tuple(sort_key(c)))
            return cands[0]
        if last:
            return None
        time.sleep(0.5)


def nearest_to(cands, ref, above=False):
    rr = rect(ref)
    best = None
    for c in cands:
        r = rect(c)
        if r is None or rr is None:
            continue
        if above and r.bottom > rr.top + 3:
            continue
        dist = abs(r.top - rr.top) + abs(r.left - rr.left)
        if best is None or dist < best[0]:
            best = (dist, c)
    return best[1] if best else None


def looks_like_input(c):
    if ctype(c) in ("Edit", "ComboBox", "Spinner"):
        return True
    cls = window_class(c).lower()
    return any(k in cls for k in ("date", "edit", "combo", "mask"))


def input_right_of(win, label, max_gap=300):
    # The input box on the same row as the label, nearest to its right.
    lr = rect(label)
    if lr is None:
        return None
    mid = (lr.top + lr.bottom) // 2
    best = None
    for c in safe(win.descendants, []):
        try:
            if c.element_info.control_type in INPUT_SKIP or not is_visible(c):
                continue
            r = c.rectangle()
            if r.top > mid or r.bottom < mid:
                continue
            if r.height() > lr.height() + 40 or r.width() > 500:
                continue
            gap = r.left - lr.right
            if gap < -5 or gap > max_gap:
                continue
            k = (0 if looks_like_input(c) else 1, gap, area(c))
            if best is None or k < best[0]:
                best = (k, c)
        except Exception:
            continue
    return best[1] if best else None


def field_for(run, label):
    if ctype(label) in ("Edit", "ComboBox"):
        return label
    box = input_right_of(run.win, label)
    if box is None:
        from pywinauto import Desktop
        lr = rect(label)
        el = safe(lambda: Desktop(backend="uia").from_point(lr.right + 60, (lr.top + lr.bottom) // 2))
        er = rect(el) if el is not None else None
        if (er is not None and ctype(el) not in INPUT_SKIP and er.height() <= lr.height() + 40
                and safe(el.process_id) == run.pid):
            box = el
    return box


def click_field(run, label, box):
    # Same spot that already worked for Case #: 60 px right of the label,
    # unless the box was found somewhere else, then its centre.
    lr = rect(label)
    x, y = lr.right + 60, (lr.top + lr.bottom) // 2
    br = rect(box) if box is not None else None
    if br is not None and not (br.left <= x <= br.right and br.top <= y <= br.bottom):
        x, y = (br.left + br.right) // 2, (br.top + br.bottom) // 2
    return click_at(run, x, y, root_of(box) if box is not None else root_of(label))


def read_value(c):
    # The text inside a box (not its label). Empty when the box does not expose it.
    if c is None:
        return ""
    for get in (lambda: c.iface_value.CurrentValue,
                lambda: c.legacy_properties().get("Value")):
        v = safe(get)
        if v:
            return str(v)
    return ""


def set_value(c, text):
    return safe(lambda: c.iface_value.SetValue(text) or True, False)


def fill_text(run, label, box, text):
    import pyautogui
    root = click_field(run, label, box)
    time.sleep(0.3)
    assert_focus(run, root)
    if box is not None and ctype(box) == "ComboBox":
        pyautogui.hotkey("ctrl", "a")
    else:
        pyautogui.press("home")
        pyautogui.hotkey("shift", "end")
        pyautogui.press("delete")
    pyautogui.write(text, interval=0.02)
    pyautogui.press("tab")
    time.sleep(0.4)
    return read_value(box)


# ------------------------- Nexus steps -------------------------

def date_seen(box):
    return parse_date_text(read_value(box)) or parse_date_text(safe(box.window_text, ""))


def native_date_picker(box):
    handle = safe(lambda: box.handle)
    if not handle or "datetimepick" not in window_class(box).lower():
        return None
    from pywinauto.controls.common_controls import DateTimePickerWrapper
    return safe(lambda: DateTimePickerWrapper(handle))


def picker_date(dtp):
    t = dtp.get_time()
    return datetime.date(t.wYear, t.wMonth, t.wDay)


def focus_box(run, box):
    # Keys for the date picker only go out when the picker itself has the keyboard focus.
    safe(box.set_focus)
    time.sleep(0.3)
    assert_focus(run, root_of(box))
    if not safe(box.has_keyboard_focus, False):
        click(run, box)
        time.sleep(0.3)
        assert_focus(run, root_of(box))
        if not safe(box.has_keyboard_focus, False):
            raise RuntimeError("Date box did not take the keyboard focus; the date was not changed")


def set_date(run, label, box, d):
    import pyautogui
    want = nexus_date_text(d)
    if box is None:
        raise RuntimeError("Box next to 'Date:' not found")
    dtp = native_date_picker(box)
    if dtp is not None:
        # Windows date picker: set it directly, then Up + Down so Nexus registers the change.
        try:
            dtp.set_time(year=d.year, month=d.month, day=d.day)
        except Exception as e:
            logging.warning("Date picker could not be set directly (%s); typing instead", e)
        focus_box(run, box)
        pyautogui.press("up")
        pyautogui.press("down")
        time.sleep(0.3)
        if safe(lambda: picker_date(dtp)) == d:
            return "verified"
        for use_right in (True, False):
            focus_box(run, box)
            for i, part in enumerate(("%02d" % d.day, "%02d" % d.month, "%04d" % d.year)):
                if i and use_right:
                    pyautogui.press("right")
                pyautogui.write(part, interval=0.05)
            time.sleep(0.4)
            if safe(lambda: picker_date(dtp)) == d:
                return "verified"
        raise RuntimeError("Date field shows %s instead of %s"
                           % (safe(lambda: picker_date(dtp)), want))
    # Text-style date box
    fill_text(run, label, box, want)
    seen = date_seen(box)
    if seen == d:
        return "verified"
    if seen is not None and set_value(box, want):
        time.sleep(0.3)
        seen = date_seen(box)
        if seen == d:
            return "verified"
    if seen is not None:
        raise RuntimeError("Date field shows %s instead of %s" % (nexus_date_text(seen), want))
    return "typed, not readable"


def date_still_ok(box, d):
    dtp = native_date_picker(box)
    if dtp is not None:
        return safe(lambda: picker_date(dtp)) == d
    seen = date_seen(box)
    return seen is None or seen == d


def fill_search_form(run, case_number, case_date):
    win = run.win
    case_label = (find_one([win], CASE_LABEL, prefer=["Text"], skip=LABEL_SKIP, timeout=10)
                  or find_one([win], CASE_LABEL_LOOSE, prefer=["Text", "Edit"], skip=LABEL_SKIP, timeout=1))
    if case_label is None:
        raise RuntimeError("'Case #' field not found on the Patients screen")

    date_labels = find_all([win], DATE_LABEL, LABEL_SKIP)
    date_label = nearest_to(date_labels, case_label, above=True) or nearest_to(date_labels, case_label)
    if date_label is None:
        raise RuntimeError("'Date' field not found on the Patients screen")
    center_label = nearest_to(find_all([win], CENTER_LABEL, LABEL_SKIP), case_label)
    if center_label is None:
        raise RuntimeError("'Center' field not found on the Patients screen")

    # Date (the case date)
    date_box = field_for(run, date_label)
    date_note = set_date(run, date_label, date_box, case_date)
    logging.info("Date set to %s (%s)", nexus_date_text(case_date), date_note)

    # Center
    center_box = field_for(run, center_label)
    center_seen = fill_text(run, center_label, center_box, CENTER_VALUE).strip()
    m = re.match(r"^\s{0,}(\d+)", center_seen)
    if m and m.group(1) != CENTER_VALUE:
        raise RuntimeError("Center shows '%s' instead of %s" % (center_seen, CENTER_VALUE))
    center_note = ("reads '%s'" % center_seen) if center_seen else "typed, not readable"
    logging.info("Center set to %s (%s)", CENTER_VALUE, center_note)

    # Case #
    case_box = field_for(run, case_label)
    seen = fill_text(run, case_label, case_box, case_number)
    if seen and seen.strip() != case_number:
        set_value(case_box, case_number)
        time.sleep(0.3)
        seen = read_value(case_box)
        if seen and seen.strip() != case_number:
            raise RuntimeError("Case # field shows '%s' instead of %s" % (seen, case_number))
    logging.info("Case # set to %s (box reads '%s')", case_number, seen)

    # Changing Center can reset other fields: check the date once more.
    if not date_still_ok(date_box, case_date):
        date_note = set_date(run, date_label, date_box, case_date)
    avoid = [rect(b) for b in (date_box, center_box, case_box) if b is not None]
    notes = "Date %s %s, Center %s %s" % (nexus_date_text(case_date), date_note, CENTER_VALUE, center_note)
    return case_label, case_box, avoid, notes


def record_pattern(case_number):
    # Exactly this case number, not a longer number or a suffixed sub-case containing it.
    numbers = [re.escape(case_number), re.escape(case_number.replace("-", ""))]
    return re.compile(r"(?<![\w/-])(%s)(?![\w/-])" % "|".join(numbers))


def record_text(c):
    t = safe(c.window_text, "") or ""
    if ctype(c) in ROW_TYPES + ["Text", "Edit", ""]:
        t = t + " " + read_value(c)
    return t


def matching_records(win, pat, avoid, rows_only):
    cands = []
    for c in safe(win.descendants, []):
        t = ctype(c)
        if t in RECORD_SKIP or (rows_only and t not in ROW_TYPES):
            continue
        if not is_visible(c) or not pat.search(record_text(c)):
            continue
        r = rect(c)
        if r is None or any(overlaps(r, a) for a in avoid if a is not None):
            continue
        cands.append(c)
    return cands


def find_record(win, case_number, avoid, timeout):
    # Returns the one result row for this case, None if there is none,
    # and raises if more than one row shows it.
    pat = record_pattern(case_number)
    end = time.time() + timeout
    cands = []
    while True:
        final = time.time() >= end
        cands = matching_records(win, pat, avoid, rows_only=not final)
        if cands or final:
            break
        time.sleep(0.5)
    if not cands:
        return None
    time.sleep(1)  # let the list finish loading, then look again
    cands = matching_records(win, pat, avoid, rows_only=False) or cands
    best = min(RECORD_RANK.get(ctype(c), 5) for c in cands)
    cands = sorted([c for c in cands if RECORD_RANK.get(ctype(c), 5) == best], key=top_left)
    rows = []
    for c in cands:
        r = rect(c)
        if not any(r.top <= o.bottom and r.bottom >= o.top for o in rows):
            rows.append(r)
    if len(rows) > 1:
        raise RuntimeError("%d records show case %s; attach this one by hand" % (len(rows), case_number))
    return cands[0]


def search_record(run, case_number, case_label, case_box, avoid):
    import pyautogui
    win = run.win
    search = find_one(
        [win], SEARCH_TITLE, prefer=CLICKABLE,
        skip=["TabItem", "Edit", "HeaderItem", "TitleBar", "Window", "Text"], timeout=5,
        sort_key=lambda c: (top_left(c)[0], -(rect(c).right if rect(c) is not None else 0)))
    if search is not None:
        click(run, search)
        record = find_record(win, case_number, avoid, 12)
        if record:
            return record
        logging.warning("No result after clicking Search; trying Enter in Case #")
    else:
        logging.warning("'Search' button not found; trying Enter in Case #")
    root = click_field(run, case_label, case_box)
    time.sleep(0.3)
    assert_focus(run, root)
    pyautogui.press("enter")
    return find_record(win, case_number, avoid, 10)


def popup_menu_items(run, before):
    items = []
    for w in nexus_windows(run.pid):
        if safe(lambda: w.handle) in before:
            continue
        items.extend(i for i in safe(lambda: w.descendants(control_type="MenuItem"), []) if is_visible(i))
    return sorted(items, key=top_left)


def escape_menu(run):
    import pyautogui
    if safe(lambda: assert_focus(run) or True, False):
        pyautogui.press("escape")
        time.sleep(0.3)


def click_emr_prescription(run, case_number, record, avoid):
    before = handles_of(nexus_windows(run.pid))
    r_before = rect_key(record)
    click(run, record)  # select the row first so the menu acts on this record
    time.sleep(0.5)
    # Look again: the same spot must still show this case number.
    same = [c for c in matching_records(run.win, record_pattern(case_number), avoid, rows_only=False)
            if rect_key(c) == r_before]
    if not same:
        raise RuntimeError("The result list changed before the right-click; stopped")
    click(run, same[0], "right")

    items = wait_for(lambda: popup_menu_items(run, before), 4) or []
    if not items:
        raise RuntimeError("Right-click menu did not open on the record")
    names = [name_of(i) for i in items]
    hits = [i for i in items if re.search(EMR_TITLE, name_of(i), re.IGNORECASE)]
    if len(hits) == 1:
        if len(items) < 2 or top_left(hits[0]) != top_left(items[-2]):
            logging.warning("EMR prescription is not the second last option here (options: %s)", names)
        click(run, hits[0])
        return
    if len(items) >= 2 and not any(names):
        # Menu without readable names: EMR prescription is the second last option.
        click(run, items[-2])
        return
    escape_menu(run)
    raise RuntimeError("'EMR prescription' not found exactly once in the right-click menu (options: %s)"
                       % ", ".join(n for n in names if n))


def window_of(c):
    # The window (EMR form, MDI child or main window) that holds a control.
    p = c
    for _ in range(40):
        p = safe(p.parent)
        if p is None:
            return None
        if ctype(p) == "Window":
            return p
    return None


def upload_buttons(run, pattern):
    return find_all(all_roots(run), pattern, BUTTON_SKIP, ["Button", "SplitButton"])


def close_leftover_upload_windows(run):
    # An EMR upload window still open from an earlier scan must never receive this scan.
    for pattern in (UPLOAD_FILE_TITLE, SELECT_FILE_TITLE):
        for b in upload_buttons(run, pattern):
            host = window_of(b)
            if host is not None and safe(lambda: host.handle) != run.main:
                logging.info("Closing a leftover Nexus window: %s", name_of(host))
                close_window(host)
    time.sleep(1)
    for w in nexus_dialogs(run.pid):
        raise RuntimeError("Closing a leftover Nexus window brought up a message: %s; answer it by hand"
                           % dialog_text(w))
    for pattern in (UPLOAD_FILE_TITLE, SELECT_FILE_TITLE):
        for b in upload_buttons(run, pattern):
            host = window_of(b)
            if host is not None and safe(lambda: host.handle) != run.main:
                raise RuntimeError("An EMR upload window from an earlier scan is still open; close it by hand")


def the_only(run, pattern, label, roots=None, timeout=15):
    # Waits for exactly one visible button with this name and returns it.
    def look():
        found = find_all(roots() if roots else all_roots(run), pattern, BUTTON_SKIP, ["Button", "SplitButton"])
        return found or None
    found = wait_for(look, timeout) or []
    if not found:
        return None
    if len(found) > 1:
        raise RuntimeError("%d '%s' buttons are open in Nexus; stopped so the file cannot go to "
                           "another record" % (len(found), label))
    return found[0]


def file_name_box(dlg):
    edits = [e for e in safe(lambda: dlg.descendants(control_type="Edit"), []) if is_visible(e)]
    for e in edits:
        if safe(lambda: e.element_info.automation_id) in ("1148", "1152"):
            return e
    for e in edits:
        if re.match(r"^\s{0,}File\s{0,}name", name_of(e), re.IGNORECASE):
            return e
    return None


def same_path(a, b):
    def norm(p):
        return (p or "").strip().strip('"').replace("/", "\\").lower()
    return norm(a) == norm(b)


def choose_file(run, before, file_path):
    import pyautogui

    def dialog():
        return ([d for d in nexus_dialogs(run.pid) if safe(lambda: d.handle) not in before] or [None])[0]

    dlg = wait_for(dialog, 15)
    if dlg is None:
        late = wait_for(dialog, 30)  # never leave a late dialog open for the next scan
        if late is not None:
            close_window(late)
        raise RuntimeError("File dialog did not open within 15 s after 'Upload file'")
    box = wait_for(lambda: file_name_box(dlg), 5)
    if box is None:
        text = dialog_text(dlg)
        close_window(dlg)
        raise RuntimeError("Nexus showed a message instead of the file dialog: %s" % text)
    if foreground() != dlg.handle:
        safe(dlg.set_focus)
        time.sleep(0.5)

    full = str(file_path).replace("/", "\\")
    seen = ""
    for attempt in range(2):
        click(run, box)
        time.sleep(0.3)
        assert_focus(run, dlg.handle)
        pyautogui.hotkey("ctrl", "a")
        if attempt == 0:
            pyautogui.write(full, interval=0.01)
        else:
            set_value(box, full)
        time.sleep(0.6)
        seen = read_value(box)
        if same_path(seen, full) or (not seen and attempt == 0):
            break
    else:
        close_window(dlg)
        raise RuntimeError("File dialog shows '%s' instead of %s" % (seen, full))

    click(run, box)
    time.sleep(0.2)
    assert_focus(run, dlg.handle)
    pyautogui.press("end")
    pyautogui.press("enter")
    if wait_for(lambda: not is_visible(dlg), 10):
        return
    msgs = [m for m in nexus_dialogs(run.pid)
            if safe(lambda: m.handle) not in before and safe(lambda: m.handle) != dlg.handle]
    text = " | ".join(dialog_text(m) for m in msgs) or "no message"
    for m in msgs:
        close_window(m)
    close_window(dlg)
    raise RuntimeError("File dialog did not accept %s (%s)" % (full, text))


def mark_pending(key, case_number):
    # Written BEFORE "Select file" is clicked, so a crash right after the click can never
    # lead to the same scan being attached twice. If it cannot be written, nothing is clicked.
    with LEDGER.open("a", encoding="utf-8") as f:
        f.write("%s\t%s\tpending\t%s\n" % (key, case_number, time.strftime("%Y-%m-%d %H:%M:%S")))


def upload_file(run, case_number, file_path, key, state):
    upload = the_only(run, UPLOAD_FILE_TITLE, "Upload file", timeout=10)
    if upload is None:
        # Older layout: EMR prescription > Upload prescription > Upload file
        up = find_one(lambda: all_roots(run), UPLOAD_PRESCRIPTION_TITLE,
                      prefer=["MenuItem", "Button"], skip=BUTTON_SKIP, timeout=2)
        if up is not None:
            click(run, up)
            upload = the_only(run, UPLOAD_FILE_TITLE, "Upload file", timeout=10)
    if upload is None:
        raise RuntimeError("'Upload file' button not found after EMR prescription")

    before = handles_of(nexus_dialogs(run.pid))
    click(run, upload)
    choose_file(run, before, file_path)
    time.sleep(1)

    sel = the_only(run, SELECT_FILE_TITLE, "Select file", timeout=15)
    if sel is None:
        raise RuntimeError("'Select file' button not found after choosing the file")
    pending = nexus_dialogs(run.pid)
    if pending:
        raise RuntimeError("Nexus showed a message before Select file: %s"
                           % " | ".join(dialog_text(m) for m in pending))
    if not wait_for(lambda: sel.is_enabled(), 10):
        raise RuntimeError("'Select file' stayed disabled; nothing was attached")

    before = handles_of(nexus_windows(run.pid))
    mark_pending(key, case_number)
    state["clicked"] = True
    try:
        click(run, sel)
        return after_select_file(run, before)
    except Exception as e:
        raise UncertainAttach(str(e)) from e


def after_select_file(run, before):
    # Watch for a message from Nexus for up to 10 s. No message = done (as on the real screen).
    msgs = []
    end = time.time() + 10
    while time.time() < end:
        if new_windows(run, before):
            time.sleep(2)  # a short "uploading" window may close by itself
            msgs = [w for w in new_windows(run, before) if is_visible(w)]
            if msgs:
                break
        time.sleep(0.5)
    if not msgs:
        return ""
    texts = []
    for m in msgs:
        text = dialog_text(m)
        texts.append(text)
        buttons = button_names(m)
        if any(re.match(QUESTION_BUTTONS, b, re.IGNORECASE) for b in buttons):
            raise RuntimeError("Nexus is asking after Select file, answer it by hand: %s" % text)
        if re.search(ERROR_WORDS, text, re.IGNORECASE):
            if window_class(m) == "#32770":
                close_window(m)
            raise RuntimeError("Nexus said after Select file: %s" % text)
        if window_class(m) == "#32770" or any(re.match(r"^\s{0,}&?OK\s{0,}$", b, re.IGNORECASE) for b in buttons):
            close_window(m)
    return "Nexus said: %s" % " | ".join(t for t in texts if t)


def cleanup(run, before):
    # Cancel dialogs this attempt opened and close an open right-click menu.
    for d in nexus_dialogs(run.pid):
        if safe(lambda: d.handle) not in before:
            close_window(d)
    if any(window_class(w) == "#32768" for w in nexus_windows(run.pid)):
        escape_menu(run)


def save_screenshot(path):
    # Keeps a picture of the Nexus screen next to the failed file, for mapping.
    try:
        import pyautogui
        FAILED_DIR.mkdir(exist_ok=True)
        shot = FAILED_DIR / ("%s_%s_screen.png" % (path.stem, time.strftime("%Y%m%d-%H%M%S")))
        pyautogui.screenshot(str(shot))
    except Exception as e:
        logging.warning("Could not save screenshot: %s", e)


def attach_in_nexus(case_number, case_date, file_path, key, state):
    win = find_main_window()
    if win is None:
        raise RuntimeError("Nexus Pro window not found (is it open and logged in?) Open windows: %s"
                           % ", ".join(visible_titles())[:300])
    run = Run(win)
    logging.info("Nexus window: '%s' (%s)", name_of(win), exe_of(run.pid))
    dialogs = nexus_dialogs(run.pid)
    if dialogs:
        raise RuntimeError("A Nexus dialog is open (%s); close it first" % dialog_text(dialogs[0]))
    close_leftover_upload_windows(run)
    if not safe(win.is_enabled, True):
        raise RuntimeError("The Nexus main window is blocked by another Nexus window; close it first")
    start = handles_of(nexus_dialogs(run.pid))
    try:
        safe(win.set_focus)
        time.sleep(0.5)
        if window_pid(foreground()) != run.pid:
            raise RuntimeError("Nexus could not be brought to the front")

        # Patients menu
        patients = find_one([win], PATIENTS_TITLE, prefer=["MenuItem", "Button", "TabItem", "Text"])
        if patients is None:
            raise RuntimeError("'Patients' menu not found in Nexus Pro")
        click(run, patients)
        time.sleep(1.5)

        # Date, Center, Case #
        case_label, case_box, avoid, notes = fill_search_form(run, case_number, case_date)

        # Search, then the record row
        record = search_record(run, case_number, case_label, case_box, avoid)
        if not record:
            raise RuntimeError("Record %s not found (%s)" % (case_number, notes))
        logging.info("Found record %s", case_number)

        # Right-click > EMR prescription > Upload file > file dialog > Select file
        click_emr_prescription(run, case_number, record, avoid)
        time.sleep(1)
        return upload_file(run, case_number, file_path, key, state)
    except Exception:
        save_screenshot(file_path)
        safe(lambda: cleanup(run, start))
        raise


# ------------------------- files and reporting -------------------------

def report(case_number, filename, status, reason=""):
    try:
        requests.post(N8N_WEBHOOK, json={
            "case_number": case_number or "",
            "filename": filename,
            "status": status,
            "reason": (reason or "")[:500],
        }, timeout=20)
    except Exception as e:
        logging.error("Could not reach n8n: %s", e)


def file_complete(path):
    # True when the scan is a whole file (the scanner has finished writing it).
    try:
        ext = path.suffix.lower()
        with open(path, "rb") as f:
            f.seek(0, 2)
            n = f.tell()
            f.seek(max(0, n - 4096))
            tail = f.read()
        if ext in (".jpg", ".jpeg"):
            return b"\xff\xd9" in tail
        if ext == ".pdf":
            return b"%%EOF" in tail
        if ext == ".png":
            return b"IEND" in tail
        from PIL import Image
        with Image.open(path) as im:
            for i in range(getattr(im, "n_frames", 1)):
                im.seek(i)
                im.load()
        return True
    except Exception:
        return False


def wait_until_stable(path, tries=120, need=3):
    last, same = -1, 0
    for _ in range(tries):
        try:
            size = path.stat().st_size
        except FileNotFoundError:
            return False
        same = same + 1 if (size == last and size > 0) else 0
        if same >= need and file_complete(path):
            return True
        last = size
        time.sleep(1)
    return False


def ledger_key(path):
    return "%s\t%s" % (path.name.lower(), path.stat().st_size)


def ledger_state(key):
    # "attached", "pending" (Select file was clicked, result unknown) or None.
    try:
        lines = LEDGER.read_text(encoding="utf-8").splitlines()
    except FileNotFoundError:
        return None
    states = [line.split("\t")[3] for line in lines
              if line.startswith(key + "\t") and len(line.split("\t")) > 3]
    if "attached" in states:
        return "attached"
    return "pending" if states else None


def remember_attached(key, case_number):
    try:
        with LEDGER.open("a", encoding="utf-8") as f:
            f.write("%s\t%s\tattached\t%s\n" % (key, case_number, time.strftime("%Y-%m-%d %H:%M:%S")))
    except OSError as e:
        logging.error("Could not write %s: %s", LEDGER, e)


def move_file(path, folder):
    for _ in range(10):
        try:
            folder.mkdir(exist_ok=True)
            dest = folder / path.name
            if dest.exists():
                dest = folder / ("%s_%s%s" % (path.stem, time.strftime("%Y%m%d-%H%M%S"), path.suffix))
            shutil.move(str(path), str(dest))
            return True
        except FileNotFoundError:
            return False
        except Exception as e:
            logging.warning("Could not move %s yet: %s", path.name, e)
            time.sleep(1)
    return False


stuck = set()  # scans that could not be moved away; not picked up again by the rescan


def finish(path, folder):
    if not move_file(path, folder):
        stuck.add(str(path))
        return False
    return True


def process(path):
    case_number = None
    key = None
    state = {"clicked": False}
    try:
        if not wait_until_stable(path):
            if not path.exists():
                logging.info("%s disappeared before it could be processed", path.name)
                return
            raise RuntimeError("Scan is incomplete or still being written")
        key = ledger_key(path)
        known = ledger_state(key)
        if known:
            logging.warning("%s was %s earlier; not attaching it again", path.name, known)
            finish(path, DONE_DIR if known == "attached" else CHECK_DIR)
            reason = ("Already attached earlier; not attached again" if known == "attached" else
                      "An earlier attempt clicked Select file; check Nexus, then delete this scan's "
                      "lines in attached_files.txt to retry")
            report(safe(lambda: format_case_number(path.stem)), path.name, "skipped", reason)
            return
        case_number, case_date = case_from_file(path)
        logging.info("%s -> case %s, date %s", path.name, case_number, nexus_date_text(case_date))
        note = attach_in_nexus(case_number, case_date, path, key, state)
    except Exception as e:
        logging.exception("Failed: %s", path.name)
        if case_number is None:
            case_number = safe(lambda: format_case_number(path.stem))
        if state["clicked"]:
            finish(path, CHECK_DIR)
            report(case_number, path.name, "uncertain",
                   "Select file was clicked, then: %s. Check the record in Nexus; this scan will not "
                   "be attached again automatically." % e)
        else:
            finish(path, FAILED_DIR)
            report(case_number, path.name, "failed", str(e))
        return
    remember_attached(key, case_number)
    if not finish(path, DONE_DIR):
        note = (note + " " if note else "") + "(attached, but the scan could not be moved to done)"
    report(case_number, path.name, "success", note)
    logging.info("Attached %s -> %s", path.name, case_number)


work_queue = queue.Queue()
seen = set()
seen_lock = threading.Lock()


def enqueue(p):
    if p.suffix.lower() not in ALLOWED_EXT or p.parent != WATCH_DIR or str(p) in stuck:
        return
    with seen_lock:
        if str(p) in seen:
            return
        seen.add(str(p))
    work_queue.put(p)


class Handler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory:
            enqueue(Path(event.src_path))

    def on_moved(self, event):
        if not event.is_directory:
            enqueue(Path(event.dest_path))


def start_observer():
    obs = Observer()
    obs.schedule(Handler(), str(WATCH_DIR), recursive=False)
    obs.start()
    return obs


def rescan():
    for p in sorted(WATCH_DIR.iterdir(), key=lambda q: safe(lambda: q.stat().st_mtime, 0)):
        if safe(p.is_file, False):
            enqueue(p)


def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s",
                        filename="nexus_attach_agent.log")
    logging.getLogger().addHandler(logging.StreamHandler())
    for folder in (DONE_DIR, FAILED_DIR, CHECK_DIR):
        folder.mkdir(exist_ok=True)
    obs = start_observer()
    logging.info("Watching %s (ledger: %s)", WATCH_DIR, LEDGER)
    next_check = 0 if PROCESS_EXISTING_ON_START else time.time() + RESCAN_SECONDS
    watcher_alerted = False
    try:
        while True:
            if time.time() >= next_check:
                next_check = time.time() + RESCAN_SECONDS
                try:
                    # The watcher stops for good after a network error on Z:, so restart it.
                    if not obs.emitters or not all(e.is_alive() for e in obs.emitters):
                        logging.error("Folder watcher stopped; restarting it")
                        if not watcher_alerted:
                            report("", "", "failed", "Folder watcher on Z: stopped (network?); restarting it")
                            watcher_alerted = True
                        safe(obs.stop)
                        obs = start_observer()
                    else:
                        watcher_alerted = False
                    rescan()
                except Exception as e:
                    logging.error("Cannot check %s: %s", WATCH_DIR, e)
            try:
                p = work_queue.get(timeout=1)
            except queue.Empty:
                continue
            try:
                if p.exists():
                    process(p)
            except Exception as e:
                logging.exception("Agent error on %s", p.name)
                report(None, p.name, "failed", "Agent error, check Nexus before retrying: %s" % e)
            finally:
                with seen_lock:
                    seen.discard(str(p))
    finally:
        safe(obs.stop)


if __name__ == "__main__":
    main()
