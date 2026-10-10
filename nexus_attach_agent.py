# Nexus Pro request-form auto-attach agent (copy-safe version: no asterisks)
# Install once: pip install watchdog pywinauto pyautogui requests pyzbar pymupdf pillow
import re
import shutil
import time
import logging
import queue
import threading
import operator
from pathlib import Path

import requests
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

# ------------------------- CONFIG -------------------------
WATCH_DIR = Path("Z:/")
N8N_WEBHOOK = "https://abvyusuf.app.n8n.cloud/webhook/nexus-attach-result"
ALLOWED_EXT = {".pdf", ".png", ".jpg", ".jpeg", ".tif", ".tiff"}

NEXUS_TITLE = "Nexus Pro"
PATIENTS_TITLE = "Patients"
CASE_FIELD_TITLE = r"Case #"
# "Center:" sits next to "Case #:" (above "Doctor:"). It filters the search,
# so it must be set before the case number is searched or the record is not found.
CENTER_FIELD_TITLE = r"^\s{0,}Cent(er|re)\s{0,}:?\s{0,}$"
CENTER_VALUE = "0"  # set to None to leave the Center field untouched
# ----------------------------------------------------------

DONE_DIR = WATCH_DIR / "done"
FAILED_DIR = WATCH_DIR / "failed"


def format_case_number(raw):
    # Only the first 12 digits count; anything after them is ignored.
    # 224601091026CSF_0007 becomes 224601-09-10
    # 146800091026SR becomes 146800-09-10
    m = re.match(r"\D{0,}(\d{6})(\d{2})(\d{2})(\d{2})", raw or "")
    if not m:
        raise ValueError("Cannot format case number from '%s'" % raw)
    number, day, month, _year = m.groups()
    if not (1 <= int(day) <= 31 and 1 <= int(month) <= 12):
        raise ValueError("Bad day/month in case number '%s'" % raw)
    return "%s-%s-%s" % (number, day, month)


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


def find_control(parents, pattern, types=None, exclude_types=None, timeout=5):
    # Never raises an ambiguity error: returns the first visible match, preferring the given types.
    exclude_types = exclude_types or []
    pat = re.compile(pattern, re.IGNORECASE)
    end = time.time() + timeout
    while True:
        matches = []
        for parent in parents:
            try:
                items = parent.descendants()
            except Exception:
                continue
            for c in items:
                try:
                    name = c.window_text() or ""
                    if not name or not pat.search(name):
                        continue
                    ctype = c.element_info.control_type
                    if ctype in exclude_types or not is_visible(c):
                        continue
                    rank = 1
                    if types and ctype in types:
                        rank = 0
                    matches.append((rank, c))
                except Exception:
                    continue
        if matches:
            matches.sort(key=lambda t: t[0])
            return matches[0][1]
        if time.time() >= end:
            return None
        time.sleep(0.5)


def input_right_of(win, label, max_gap=300):
    # The Edit/ComboBox on the same row as the label, nearest to its right.
    lr = label.rectangle()
    mid = (lr.top + lr.bottom) // 2
    best = None
    for c in win.descendants():
        try:
            if c.element_info.control_type not in ("Edit", "ComboBox"):
                continue
            if not is_visible(c):
                continue
            r = c.rectangle()
            if r.top > mid or r.bottom < mid:
                continue
            gap = r.left - lr.right
            if gap < -5 or gap > max_gap:
                continue
            if best is None or gap < best[0]:
                best = (gap, c)
        except Exception:
            continue
    return best[1] if best else None


def set_center(win, value):
    import pyautogui

    label = find_control([win], CENTER_FIELD_TITLE, types=["Text"])
    if label is None:
        raise RuntimeError("'Center' field not found")
    if label.element_info.control_type in ("Edit", "ComboBox"):
        box = label
    else:
        box = input_right_of(win, label)
    if box is not None:
        box.click_input()
    else:
        r = label.rectangle()
        pyautogui.click(r.right + 60, (r.top + r.bottom) // 2)
    time.sleep(0.3)
    pyautogui.hotkey("ctrl", "a")
    pyautogui.write(value, interval=0.02)
    pyautogui.press("tab")
    time.sleep(0.5)
    if box is not None:
        try:
            logging.info("Center field now reads '%s'", box.window_text())
        except Exception:
            pass


def all_roots(win):
    from pywinauto import Desktop
    roots = [win]
    for w in Desktop(backend="uia").windows():
        try:
            if w.handle != win.handle and is_visible(w):
                roots.append(w)
        except Exception:
            pass
    return roots


def attach_in_nexus(case_number, file_path):
    from pywinauto import Desktop
    import pyautogui

    win = find_main_window()
    if win is None:
        raise RuntimeError("Nexus Pro window not found (is it open and logged in?)")
    try:
        win.set_focus()
    except Exception:
        pass
    time.sleep(0.5)

    # Patients menu
    patients = find_control([win], PATIENTS_TITLE, types=["MenuItem", "Button", "TabItem", "Text"])
    if patients is None:
        raise RuntimeError("'Patients' menu not found in Nexus Pro")
    patients.click_input()
    time.sleep(1.5)

    # Center (must be set before searching, otherwise the search is filtered out)
    if CENTER_VALUE is not None:
        set_center(win, CENTER_VALUE)

    # Case # field
    label = find_control([win], CASE_FIELD_TITLE, types=["Edit", "Text"])
    if label is None:
        raise RuntimeError("'Case #' field not found")
    if label.element_info.control_type != "Edit":
        r = label.rectangle()
        pyautogui.click(r.right + 60, (r.top + r.bottom) // 2)
    else:
        label.click_input()
    time.sleep(0.3)
    pyautogui.hotkey("ctrl", "a")
    pyautogui.write(case_number, interval=0.02)
    pyautogui.press("enter")
    time.sleep(2)

    # Record row
    row = find_control([win], re.escape(case_number), exclude_types=["Edit"], timeout=8)
    if row is None:
        raise RuntimeError("Record %s not found in Nexus Pro (Center = %s)"
                           % (case_number, CENTER_VALUE))
    row.right_click_input()
    time.sleep(0.7)

    # EMR prescription
    item = find_control(all_roots(win), "EMR prescription", types=["MenuItem"])
    if item is None:
        raise RuntimeError("'EMR prescription' menu item not found")
    item.click_input()
    time.sleep(0.7)

    # Upload prescription
    up = find_control(all_roots(win), "Upload prescription", types=["MenuItem", "Button"])
    if up is None:
        raise RuntimeError("'Upload prescription' not found")
    up.click_input()
    time.sleep(1.5)

    # Upload file
    win2 = find_main_window() or win
    upload_btn = find_control(all_roots(win2), "Upload file", types=["Button"])
    if upload_btn is None:
        raise RuntimeError("'Upload file' button not found")
    upload_btn.click_input()
    time.sleep(1.5)

    # Select file
    sel = find_control(all_roots(win2), "Select file", types=["Button"], timeout=6)
    if sel is None:
        raise RuntimeError("'Select file' button not found after clicking Upload file")
    sel.click_input()
    time.sleep(1.5)

    # File dialog
    dlg = None
    end = time.time() + 10
    while time.time() < end and dlg is None:
        for d in Desktop(backend="uia").windows(class_name="#32770"):
            if is_visible(d):
                dlg = d
                break
        time.sleep(0.5)
    if dlg is None:
        raise RuntimeError("File selection dialog did not open")
    try:
        dlg.set_focus()
    except Exception:
        pass
    time.sleep(0.5)
    pyautogui.write(str(file_path).replace("/", "\\"), interval=0.01)
    pyautogui.press("enter")
    time.sleep(2.5)
    still = [d for d in Desktop(backend="uia").windows(class_name="#32770") if is_visible(d)]
    if still:
        raise RuntimeError("File dialog still open - file may not have been selected")


def report(case_number, filename, status, reason=""):
    try:
        requests.post(N8N_WEBHOOK, json={
            "case_number": case_number or "",
            "filename": filename,
            "status": status,
            "reason": reason,
        }, timeout=20)
    except Exception as e:
        logging.error("Could not reach n8n: %s", e)


def save_screenshot(path):
    # Keeps a picture of the Nexus screen next to the failed file, for mapping.
    try:
        import pyautogui
        pyautogui.screenshot(str(FAILED_DIR / (path.stem + "_screen.png")))
    except Exception as e:
        logging.warning("Could not save screenshot: %s", e)


def wait_until_stable(path, tries=20):
    last = -1
    for _ in range(tries):
        try:
            size = path.stat().st_size
        except FileNotFoundError:
            return False
        if size == last and size > 0:
            return True
        last = size
        time.sleep(1)
    return False


def process(path):
    case_number = None
    try:
        if not wait_until_stable(path):
            raise RuntimeError("File not readable / still being written")
        code = read_barcode(path)
        if code:
            try:
                case_number = format_case_number(code)
            except ValueError:
                case_number = None
        if case_number is None:
            case_number = format_case_number(path.stem)
        attach_in_nexus(case_number, path)
        shutil.move(str(path), str(DONE_DIR / path.name))
        report(case_number, path.name, "success")
        logging.info("Attached %s -> %s", path.name, case_number)
    except Exception as e:
        logging.exception("Failed: %s", path.name)
        save_screenshot(path)
        try:
            shutil.move(str(path), str(FAILED_DIR / path.name))
        except Exception:
            pass
        report(case_number, path.name, "failed", str(e))


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
