# Nexus Pro request-form auto-attach agent (copy-safe version: no asterisks)
# Install once: pip install watchdog pywinauto pyautogui requests pyzbar pymupdf pillow
#
# For every scan that lands in Z:\ (example: 224601091026CSF_0007.jpg):
#   1. The first 12 digits give the case 224601-09-10 and the date 09-Oct-2026.
#   2. Nexus Pro: Patients > Date = 09-Oct-2026 > Center = 0 > Case # = 224601-09-10 > Search.
#   3. Select the record, right-click it > EMR prescription > Upload file.
#   4. Windows file dialog: the full path of the scan > Open.
#   5. Back in Nexus: Select file.
#   6. The result goes to n8n (Telegram alert); the scan moves to Z:\done or Z:\failed.
# Every step is checked. When something does not match, the agent stops and reports
# instead of guessing, so a form is never attached to the wrong record.
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
PROCESS_EXISTING_ON_START = True  # also handle scans already waiting in Z:\ at startup
USE_BARCODE = True  # if the form has a readable barcode, it must match the filename

NEXUS_TITLE = "Nexus Pro"
CENTER_VALUE = "0"

# Texts on the Nexus screens (regular expressions, not case sensitive)
PATIENTS_TITLE = r"Patients"
DATE_LABEL = r"^\s{0,}Date\s{0,}:?\s{0,}$"
CENTER_LABEL = r"^\s{0,}Cent(er|re)\s{0,}:?\s{0,}$"
CASE_LABEL = r"^\s{0,}Case\s{0,}#\s{0,}:?\s{0,}$"
CASE_LABEL_LOOSE = r"Case\s{0,}#"
SEARCH_TITLE = r"^\s{0,}&?Search\s{0,}$"
EMR_TITLE = r"EMR\s{0,}prescription"
UPLOAD_PRESCRIPTION_TITLE = r"Upload\s{0,}prescription"
UPLOAD_FILE_TITLE = r"Upload\s{0,}file"
SELECT_FILE_TITLE = r"Select\s{0,}file"
ERROR_WORDS = r"error|fail|invalid|unable|cannot|can't|not found|denied|exception|already"
# ----------------------------------------------------------

DONE_DIR = WATCH_DIR / "done"
FAILED_DIR = WATCH_DIR / "failed"
LEDGER = DONE_DIR / "attached_files.txt"

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Control types that are never a field label / input box / search result.
LABEL_SKIP = ["HeaderItem", "Header", "DataItem", "ListItem", "TreeItem", "MenuItem",
              "TabItem", "Button", "Hyperlink", "TitleBar", "Window"]
INPUT_SKIP = ["Text", "Group", "Window", "TitleBar", "MenuBar", "Menu", "MenuItem",
              "ToolBar", "StatusBar", "Header", "HeaderItem", "Image", "ScrollBar",
              "Thumb", "Separator", "Hyperlink", "Tab", "TabItem", "Table", "DataGrid",
              "List", "Tree", "TreeItem", "ListItem", "DataItem", "Button"]
RECORD_SKIP = ["Window", "TitleBar", "StatusBar", "MenuBar", "Menu", "MenuItem",
               "Button", "TabItem", "ComboBox", "Spinner", "ToolBar"]
RECORD_RANK = {"DataItem": 0, "ListItem": 0, "TreeItem": 0, "Custom": 1, "Text": 2, "Edit": 3}
CLICKABLE = ["Button", "SplitButton", "MenuItem", "ListItem", "Hyperlink"]


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
    if file_case and bc_case and file_case != bc_case:
        raise RuntimeError("Barcode says %s but filename says %s - check this form by hand"
                           % (bc_case, file_case))
    if file_case:
        return file_case, file_date
    if bc_case:
        return bc_case, bc_date
    raise RuntimeError(file_err or "No case number found")


# ------------------------- UI helpers -------------------------

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


def rect(c):
    return safe(lambda: c.rectangle())


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
    # Several windows can be titled Nexus Pro. Take the biggest visible one.
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
            if is_visible(w):
                cands.append(w)
        if cands:
            cands.sort(key=area, reverse=True)
            return cands[0]
        time.sleep(0.5)
    logging.error("No window title contains '%s'. Visible windows: %s",
                  NEXUS_TITLE, visible_titles())
    return None


def all_roots(win, pid):
    # The main window plus every other visible top-level window of the Nexus process
    # (pop-up menus, the EMR window, dialogs). Other programs are never searched.
    from pywinauto import Desktop
    roots = [win]
    for w in Desktop(backend="uia").windows():
        try:
            if w.handle != win.handle and w.process_id() == pid and is_visible(w):
                roots.append(w)
        except Exception:
            pass
    return roots


def top_handles(pid):
    from pywinauto import Desktop
    handles = set()
    for w in Desktop(backend="uia").windows():
        try:
            if w.process_id() == pid and is_visible(w):
                handles.add(w.handle)
        except Exception:
            pass
    return handles


def nexus_dialogs(pid):
    from pywinauto import Desktop
    found = []
    for d in Desktop(backend="uia").windows(class_name="#32770"):
        try:
            if d.process_id() == pid and is_visible(d):
                found.append(d)
        except Exception:
            pass
    return found


def new_dialogs(pid, before):
    return [d for d in nexus_dialogs(pid) if safe(lambda: d.handle) not in before]


def dialog_text(d):
    parts = []
    for c in safe(d.descendants, []):
        if ctype(c) == "Text":
            t = (safe(c.window_text, "") or "").strip()
            if t and t not in parts:
                parts.append(t)
    return " ".join(parts)[:300] or (safe(d.window_text, "") or "")


def close_window(w):
    safe(w.close)
    time.sleep(0.5)


def assert_focus(pid):
    # Keys are only sent while Nexus (or its own dialog) is the active window.
    from pywinauto import handleprops, win32functions
    h = win32functions.GetForegroundWindow()
    if h and handleprops.processid(h) != pid:
        raise RuntimeError("Another window took the focus ('%s'); stopped so nothing is typed into it"
                           % safe(lambda: handleprops.text(h), "?"))


def find_all(parents, pattern, skip=(), only=None):
    pat = re.compile(pattern, re.IGNORECASE)
    found = []
    for parent in parents:
        items = safe(parent.descendants, [])
        for c in items:
            try:
                name = c.window_text() or ""
                if not name or not pat.search(name):
                    continue
                t = c.element_info.control_type
                if t in skip or (only and t not in only):
                    continue
                if is_visible(c):
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
    cls = (safe(lambda: c.element_info.class_name, "") or "").lower()
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


def field_for(win, label):
    if ctype(label) in ("Edit", "ComboBox"):
        return label
    box = input_right_of(win, label)
    if box is None:
        from pywinauto import Desktop
        lr = rect(label)
        el = safe(lambda: Desktop(backend="uia").from_point(lr.right + 60, (lr.top + lr.bottom) // 2))
        er = rect(el) if el is not None else None
        if er is not None and ctype(el) not in INPUT_SKIP and er.height() <= lr.height() + 40:
            box = el
    return box


def click_field(label, box):
    # Same spot that already worked for Case #: 60 px right of the label,
    # unless the box was found somewhere else, then its centre.
    import pyautogui
    lr = rect(label)
    x, y = lr.right + 60, (lr.top + lr.bottom) // 2
    br = rect(box) if box is not None else None
    if br is not None and not (br.left <= x <= br.right and br.top <= y <= br.bottom):
        x, y = (br.left + br.right) // 2, (br.top + br.bottom) // 2
    pyautogui.click(x, y)


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


def fill_text(pid, label, box, text):
    import pyautogui
    click_field(label, box)
    time.sleep(0.3)
    assert_focus(pid)
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
    cls = (safe(lambda: box.element_info.class_name, "") or "").lower()
    if not handle or "datetimepick" not in cls:
        return None
    from pywinauto.controls.common_controls import DateTimePickerWrapper
    return safe(lambda: DateTimePickerWrapper(handle))


def picker_date(dtp):
    t = dtp.get_time()
    return datetime.date(t.wYear, t.wMonth, t.wDay)


def set_date(pid, label, box, d):
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
        box.set_focus()
        time.sleep(0.3)
        assert_focus(pid)
        pyautogui.press("up")
        pyautogui.press("down")
        time.sleep(0.3)
        if safe(lambda: picker_date(dtp)) == d:
            return "verified"
        for use_right in (True, False):
            box.set_focus()
            time.sleep(0.3)
            assert_focus(pid)
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
    fill_text(pid, label, box, want)
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


def fill_search_form(win, pid, case_number, case_date):
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
    date_box = field_for(win, date_label)
    date_note = set_date(pid, date_label, date_box, case_date)
    logging.info("Date set to %s (%s)", nexus_date_text(case_date), date_note)

    # Center
    center_box = field_for(win, center_label)
    center_seen = fill_text(pid, center_label, center_box, CENTER_VALUE)
    logging.info("Center set to %s (box reads '%s')", CENTER_VALUE, center_seen)

    # Case #
    case_box = field_for(win, case_label)
    seen = fill_text(pid, case_label, case_box, case_number)
    if seen and seen.strip() != case_number:
        set_value(case_box, case_number)
        time.sleep(0.3)
        seen = read_value(case_box)
        if seen and seen.strip() != case_number:
            raise RuntimeError("Case # field shows '%s' instead of %s" % (seen, case_number))
    logging.info("Case # set to %s (box reads '%s')", case_number, seen)

    # Changing Center can reset other fields: check the date once more.
    if not date_still_ok(date_box, case_date):
        date_note = set_date(pid, date_label, date_box, case_date)
    return case_label, case_box, [rect(b) for b in (date_box, center_box, case_box) if b is not None], date_note


def find_record(win, case_number, avoid, timeout):
    # A search result showing exactly this case number (not a longer number containing it).
    numbers = [re.escape(case_number), re.escape(case_number.replace("-", ""))]
    pattern = r"(?<![\d-])(%s)(?![\d-])" % "|".join(numbers)

    def look():
        cands = []
        for c in find_all([win], pattern, RECORD_SKIP):
            r = rect(c)
            if r is None or any(overlaps(r, a) for a in avoid if a is not None):
                continue
            cands.append(c)
        if not cands:
            return None
        cands.sort(key=lambda c: (RECORD_RANK.get(ctype(c), 5),) + top_left(c))
        return cands[0]
    return wait_for(look, timeout)


def search_record(win, pid, case_number, case_label, case_box, avoid):
    import pyautogui
    search = find_one(
        [win], SEARCH_TITLE, prefer=CLICKABLE,
        skip=["TabItem", "Edit", "HeaderItem", "TitleBar", "Window", "Text"], timeout=5,
        sort_key=lambda c: (top_left(c)[0], -(rect(c).right if rect(c) is not None else 0)))
    if search is not None:
        search.click_input()
        record = find_record(win, case_number, avoid, 12)
        if record:
            return record
        logging.warning("No result after clicking Search; trying Enter in Case #")
    else:
        logging.warning("'Search' button not found; trying Enter in Case #")
    click_field(case_label, case_box)
    time.sleep(0.3)
    assert_focus(pid)
    pyautogui.press("enter")
    return find_record(win, case_number, avoid, 10)


def click_emr_prescription(win, pid, record):
    import pyautogui
    before = top_handles(pid)
    record.click_input()  # select the row first so the menu acts on this record
    time.sleep(0.5)
    record.right_click_input()
    time.sleep(0.8)

    def popups():
        return [w for w in all_roots(win, pid)[1:] if safe(lambda: w.handle) not in before]

    item = find_one(popups, EMR_TITLE, prefer=["MenuItem"], timeout=4)
    if item is not None:
        item.click_input()
        return
    items = [i for w in popups() for i in safe(lambda: w.descendants(control_type="MenuItem"), [])
             if is_visible(i)]
    names = [(safe(i.window_text, "") or "").strip() for i in items]
    if len(items) >= 2 and not any(names):
        # Menu without readable names: EMR prescription is the second last option.
        items[-2].click_input()
        return
    if popups():
        assert_focus(pid)
        pyautogui.press("escape")
    if not items:
        raise RuntimeError("Right-click menu did not open on the record")
    raise RuntimeError("'EMR prescription' not in the right-click menu (options: %s)"
                       % ", ".join(n for n in names if n))


def file_name_box(dlg):
    edits = [e for e in safe(lambda: dlg.descendants(control_type="Edit"), []) if is_visible(e)]
    for e in edits:
        if safe(lambda: e.element_info.automation_id) in ("1148", "1152"):
            return e
    for e in edits:
        if re.match(r"^\s{0,}File\s{0,}name", safe(e.window_text, "") or "", re.IGNORECASE):
            return e
    return None


def same_path(a, b):
    def norm(p):
        return (p or "").strip().strip('"').replace("/", "\\").lower()
    return norm(a) == norm(b)


def choose_file(pid, before, file_path):
    import pyautogui
    dlg = wait_for(lambda: (new_dialogs(pid, before) or [None])[0], 15)
    if dlg is None:
        raise RuntimeError("File dialog did not open after 'Upload file'")
    box = wait_for(lambda: file_name_box(dlg), 5)
    if box is None:
        text = dialog_text(dlg)
        close_window(dlg)
        raise RuntimeError("Nexus showed a message instead of the file dialog: %s" % text)

    full = str(file_path).replace("/", "\\")
    seen = ""
    for attempt in range(2):
        box.click_input()
        time.sleep(0.3)
        assert_focus(pid)
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

    box.click_input()
    time.sleep(0.2)
    assert_focus(pid)
    pyautogui.press("end")
    pyautogui.press("enter")
    if wait_for(lambda: not is_visible(dlg), 10):
        return
    msgs = [m for m in new_dialogs(pid, before) if safe(lambda: m.handle) != safe(lambda: dlg.handle)]
    text = " | ".join(dialog_text(m) for m in msgs) or "no message"
    for m in msgs:
        close_window(m)
    close_window(dlg)
    raise RuntimeError("File dialog did not accept %s (%s)" % (full, text))


def upload_file(win, pid, file_path):
    def roots():
        return all_roots(win, pid)

    skip = ["TitleBar", "Window", "Edit"]
    upload = find_one(roots, UPLOAD_FILE_TITLE, prefer=["Button"], skip=skip, timeout=10)
    if upload is None:
        # Older layout: EMR prescription > Upload prescription > Upload file
        up = find_one(roots, UPLOAD_PRESCRIPTION_TITLE, prefer=["MenuItem", "Button"], skip=skip, timeout=2)
        if up is not None:
            up.click_input()
            upload = find_one(roots, UPLOAD_FILE_TITLE, prefer=["Button"], skip=skip, timeout=10)
    if upload is None:
        raise RuntimeError("'Upload file' button not found after EMR prescription")

    before = {safe(lambda: d.handle) for d in nexus_dialogs(pid)}
    upload.click_input()
    choose_file(pid, before, file_path)
    time.sleep(1)

    sel = find_one(roots, SELECT_FILE_TITLE, prefer=["Button"], skip=skip, timeout=15)
    if sel is None:
        raise RuntimeError("'Select file' button not found after choosing the file")
    before = {safe(lambda: d.handle) for d in nexus_dialogs(pid)}
    sel.click_input()
    time.sleep(2.5)

    msgs = new_dialogs(pid, before)
    if not msgs:
        return ""
    text = " | ".join(dialog_text(m) for m in msgs)
    if re.search(ERROR_WORDS, text, re.IGNORECASE):
        for m in msgs:
            close_window(m)
        raise RuntimeError("Nexus said after Select file: %s" % text)
    for m in msgs:
        close_window(m)
    if any(is_visible(m) for m in msgs):
        raise RuntimeError("Nexus is asking a question after Select file, answer it by hand: %s" % text)
    return "Nexus said: %s" % text


def cleanup(pid, before_dialogs):
    import pyautogui
    from pywinauto import Desktop
    for d in new_dialogs(pid, before_dialogs):
        close_window(d)
    for w in Desktop(backend="uia").windows():
        try:
            if w.process_id() == pid and is_visible(w) and ctype(w) == "Menu":
                assert_focus(pid)
                pyautogui.press("escape")
                break
        except Exception:
            pass


def save_screenshot(path):
    # Keeps a picture of the Nexus screen next to the failed file, for mapping.
    try:
        import pyautogui
        FAILED_DIR.mkdir(exist_ok=True)
        shot = FAILED_DIR / ("%s_%s_screen.png" % (path.stem, time.strftime("%Y%m%d-%H%M%S")))
        pyautogui.screenshot(str(shot))
    except Exception as e:
        logging.warning("Could not save screenshot: %s", e)


def attach_in_nexus(case_number, case_date, file_path):
    win = find_main_window()
    if win is None:
        raise RuntimeError("Nexus Pro window not found (is it open and logged in?) Open windows: %s"
                           % ", ".join(visible_titles())[:300])
    pid = win.process_id()
    start_dialogs = {safe(lambda: d.handle) for d in nexus_dialogs(pid)}
    if start_dialogs:
        raise RuntimeError("A Nexus dialog is already open (%s); close it first"
                           % dialog_text(nexus_dialogs(pid)[0]))
    try:
        safe(win.set_focus)
        time.sleep(0.5)

        # Patients menu
        patients = find_one([win], PATIENTS_TITLE, prefer=["MenuItem", "Button", "TabItem", "Text"])
        if patients is None:
            raise RuntimeError("'Patients' menu not found in Nexus Pro")
        patients.click_input()
        time.sleep(1.5)

        # Date, Center, Case #
        case_label, case_box, avoid, date_note = fill_search_form(win, pid, case_number, case_date)

        # Search, then the record row
        record = search_record(win, pid, case_number, case_label, case_box, avoid)
        if not record:
            raise RuntimeError("Record %s not found (Date %s %s, Center %s)"
                               % (case_number, nexus_date_text(case_date), date_note, CENTER_VALUE))
        logging.info("Found record %s", case_number)

        # Right-click > EMR prescription > Upload file > file dialog > Select file
        click_emr_prescription(win, pid, record)
        time.sleep(1)
        return upload_file(win, pid, file_path)
    except Exception:
        save_screenshot(file_path)
        safe(lambda: cleanup(pid, start_dialogs))
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


def wait_until_stable(path, tries=20):
    last = -1
    for _ in range(tries):
        try:
            size = path.stat().st_size
        except FileNotFoundError:
            return False
        if size == last and size > 0:
            try:
                with open(path, "rb"):
                    return True
            except OSError:
                pass
        last = size
        time.sleep(1)
    return False


def ledger_key(path):
    try:
        size = path.stat().st_size
    except OSError:
        size = -1
    return "%s\t%s" % (path.name.lower(), size)


def already_attached(key):
    try:
        lines = LEDGER.read_text(encoding="utf-8").splitlines()
    except OSError:
        return False
    return any(line.startswith(key + "\t") for line in lines)


def remember_attached(key, case_number):
    try:
        with LEDGER.open("a", encoding="utf-8") as f:
            f.write("%s\t%s\t%s\n" % (key, case_number, time.strftime("%Y-%m-%d %H:%M:%S")))
    except OSError as e:
        logging.error("Could not write %s: %s", LEDGER, e)


def move_file(path, folder):
    dest = folder / path.name
    if dest.exists():
        dest = folder / ("%s_%s%s" % (path.stem, time.strftime("%Y%m%d-%H%M%S"), path.suffix))
    for _ in range(10):
        try:
            shutil.move(str(path), str(dest))
            return True
        except FileNotFoundError:
            return False
        except Exception as e:
            logging.warning("Could not move %s yet: %s", path.name, e)
            time.sleep(1)
    return False


def process(path):
    case_number = None
    try:
        if not wait_until_stable(path):
            if not path.exists():
                logging.info("%s disappeared before it could be processed", path.name)
                return
            raise RuntimeError("File not readable / still being written")
        key = ledger_key(path)
        if already_attached(key):
            logging.warning("%s was already attached earlier; not attaching it twice", path.name)
            move_file(path, DONE_DIR)
            report(safe(lambda: format_case_number(path.stem)), path.name, "skipped",
                   "Already attached earlier; not attached again")
            return
        case_number, case_date = case_from_file(path)
        logging.info("%s -> case %s, date %s", path.name, case_number, nexus_date_text(case_date))
        note = attach_in_nexus(case_number, case_date, path)
    except Exception as e:
        logging.exception("Failed: %s", path.name)
        if case_number is None:
            case_number = safe(lambda: format_case_number(path.stem))
        move_file(path, FAILED_DIR)
        report(case_number, path.name, "failed", str(e))
        return
    remember_attached(key, case_number)
    if not move_file(path, DONE_DIR):
        note = (note + " " if note else "") + "(attached, but the scan could not be moved to done)"
    report(case_number, path.name, "success", note)
    logging.info("Attached %s -> %s", path.name, case_number)


work_queue = queue.Queue()
seen = set()
seen_lock = threading.Lock()


def enqueue(p):
    if p.suffix.lower() not in ALLOWED_EXT or p.parent != WATCH_DIR:
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


def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s",
                        filename="nexus_attach_agent.log")
    logging.getLogger().addHandler(logging.StreamHandler())
    DONE_DIR.mkdir(exist_ok=True)
    FAILED_DIR.mkdir(exist_ok=True)
    obs = Observer()
    obs.schedule(Handler(), str(WATCH_DIR), recursive=False)
    obs.start()
    logging.info("Watching %s", WATCH_DIR)
    if PROCESS_EXISTING_ON_START:
        for p in sorted(WATCH_DIR.iterdir(), key=lambda q: safe(lambda: q.stat().st_mtime, 0)):
            if p.is_file():
                enqueue(p)
    try:
        while True:
            try:
                p = work_queue.get(timeout=1)
            except queue.Empty:
                continue
            if p.exists():
                process(p)
            with seen_lock:
                seen.discard(str(p))
    finally:
        obs.stop()
        obs.join()


if __name__ == "__main__":
    main()
