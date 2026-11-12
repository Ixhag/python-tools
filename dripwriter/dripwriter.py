import configparser
import copy
import ctypes
import difflib
import os
import random
import re
import shutil
import subprocess
import sys
from pathlib import Path
from time import monotonic, sleep

import keyboard

# ============================================================
# INPUT BACKEND GLOBALS
# ============================================================

INPUT_BACKEND = "windows"
FIREFOX_DRIVER = None
FIREFOX_PROFILE_COPY = None       # temp dir holding the launched profile copy
FIREFOX_MARIONETTE_PORT = 2828
GECKODRIVER_PATH = None
SELENIUM_AVAILABLE = False
try:
    from selenium import webdriver
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.firefox.options import Options as FirefoxOptions
    from selenium.webdriver.firefox.service import Service as FirefoxService
    SELENIUM_AVAILABLE = True
except ImportError:
    webdriver = ActionChains = Keys = FirefoxOptions = FirefoxService = None

try:
    from docx import Document
    from docx.shared import Pt
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import win32com.client
    WIN32_WORD_AVAILABLE = sys.platform == "win32"
except ImportError:
    WIN32_WORD_AVAILABLE = False

# ============================================================
# SETTINGS
# ============================================================

# ---------- Firefox (Background) ----------
# Only used when you pick "2. Firefox (Background)" at startup.

# The Firefox profile whose Google login gets borrowed.  The script copies
# this profile to a temp folder and launches a separate, invisible Firefox
# from the copy, so this profile itself is never opened and your normal
# Firefox keeps working.  The name must match about:profiles exactly.
FIREFOX_PROFILE_NAME = "Firefox 2"

# Which Google account signed into that profile should open the doc.
#   0 = the first account you signed in with (the default one)
#   1 = the second, 2 = the third, and so on
# You can paste a plain link and this gets inserted for you, so
#   https://docs.google.com/document/d/FILEID/edit?tab=t.0
# becomes
#   https://docs.google.com/document/u/1/d/FILEID/edit?tab=t.0
# Set to None to leave whatever link you paste exactly as it is.
FIREFOX_ACCOUNT_INDEX = 0

# Run with no window at all.  Nothing can cover a window that doesn't exist,
# so this is the only configuration that CANNOT be interrupted by another
# window.  Set to False if you want to watch it type.
FIREFOX_HEADLESS = True

# The Google Doc to type into.  Leave blank to be asked at startup.
# Required because we launch a fresh Firefox rather than attaching to one
# you set up by hand.
FIREFOX_DOC_URL = ""

# ---------- General ----------
PLAY_SOUND = True
EDITOR = "google_docs"              # "google_docs" or "word"
START_KEY = "backspace"
STOP_KEY = "escape"

# ---------- Run / pacing ----------
TARGET_TIME = ""
ASK_FOR_TARGET_TIME = True
MAX_SPEEDUP = 3.0
ANNOUNCE_LONG_BREAKS = True

# TEST_MODE skips pacing, delays and typos.
TEST_MODE = True
TEST_MODE_KEY_GAP = 0.02
TEST_MODE_SHORTCUT_GAP = 0.08
TEST_MODE_CATCH_UP_EVERY = 300
TEST_MODE_CATCH_UP_PAUSE = 1.5
TEST_MODE_FINISH_WAIT = 4.0

# ---------- Typing rhythm ----------
MIN_WPM = 40
MAX_WPM = 70
TYPO_CHANCE = 0.05
PAUSE_AFTER_SENTENCE = (1.0, 2.5)
PAUSE_AFTER_COMMA = (0.3, 1.0)
PAUSE_AFTER_NEWLINE = (1.5, 5.0)
PAUSE_NOTICE_TYPO = (0.2, 1.0)
THINK_CHANCE = 0.05
THINK_PAUSE = (3.0, 6.0)

# ---------- Long-run break distribution ----------
PAUSE_POINT_WEIGHTS = {
    "sentence": 1.0,
    "paragraph": 8.0,
    "section": 20.0,
    "references": 4.0,
}
PAUSE_BETWEEN_PARAGRAPHS = (1.5, 4.0)

# ---------- DOCX recreation ----------
DOCX_MODE_ENABLED = True
DOCX_ALLOW_TYPOS = True
DOCX_SHOW_PROGRESS = True
DOCX_START_DELAY = 2.0
DOCX_FORMAT_PAUSE = (0.08, 0.22)
SUPERSCRIPTS = True
REAL_LISTS = True
LIST_MARKER_GAP = "\u00a0"
COPY_INDENTS = True
KEEP_LIST_TOGETHER = True
PAGE_BREAKS = True

DEFAULT_REFERENCES_MODE = "last"
PAUSE_BEFORE_REFERENCES = (1.0, 3.0)
PAUSE_AFTER_REFERENCES = (0.8, 2.0)
REFERENCE_HEADINGS = [
    "references", "reference", "reference list", "bibliography", "works cited",
    "sources", "works consulted", "literature cited", "citations",
]

# ---------- Safety / focus ----------
KEEP_PC_AWAKE = True
FOCUS_GUARD = True
FOCUS_LOCK_TO_THIS_DOC = True
FOCUS_RESUME_AFTER = 2.0
FOCUS_TITLE_MUST_CONTAIN = {"google_docs": "Google Docs", "word": ""}
FOCUS_STATUS_PREVIEW = 90
REANCHOR_AFTER_EVERY_SENTENCE = True

# ---------- Google Docs preferences ----------
CHECK_DOCS_PREFERENCES = True

# ---------- Compare ----------
COMPARE_IGNORE_BLANK_LINES = True
COMPARE_IGNORE_SMART_QUOTES = True
COMPARE_FORMATTING = True
COMPARE_HEADINGS_AND_ALIGNMENT = True
COMPARE_MAX_DIFFERENCES = 20

# ---------- Microsoft Word only ----------
DOCX_TYPE_FONT_SIZES = True
WINDOWS_WORD_EXACT_FORMAT = True

# ============================================================

MAC = sys.platform == "darwin"
WINDOWS = sys.platform == "win32"
MOD = "command" if MAC else "ctrl"

NEIGHBOURS = {
    "q": "wa", "w": "qes", "e": "wrd", "r": "etf", "t": "ryg", "y": "tuh",
    "u": "yij", "i": "uok", "o": "ipl", "p": "ol", "a": "qsz", "s": "awdx",
    "d": "sefc", "f": "drgv", "g": "fthb", "h": "gyjn", "j": "hukm",
    "k": "jil", "l": "kop", "z": "asx", "x": "zsdc", "c": "xdfv",
    "v": "cfgb", "b": "vghn", "n": "bhjm", "m": "njk",
}


def editor_keys(editor=None, mac=None):
    """Shortcut table for the chosen editor and operating system."""
    editor = editor or EDITOR
    mac = MAC if mac is None else mac
    mod = "command" if mac else "ctrl"
    if editor == "google_docs":
        return {
            "doc_end": f"{mod}+down" if mac else "ctrl+end",
            "para_up": "alt+up" if mac else "ctrl+up",
            "heading": lambda n: f"{mod}+alt+{min(max(n, 1), 6)}",
            "normal": f"{mod}+alt+0",
            "clear_char": f"{mod}+\\",
            "align": {"left": f"{mod}+shift+l", "center": f"{mod}+shift+e",
                      "right": f"{mod}+shift+r", "justify": f"{mod}+shift+j"},
            "font_size_box": None,
            "list": {"number": f"{mod}+shift+7", "bullet": f"{mod}+shift+8"},
            "superscript": f"{mod}+period", "subscript": f"{mod}+comma",
            "indent_more": f"{mod}+]", "indent_less": f"{mod}+[",
        }
    return {
        "doc_end": f"{mod}+end",
        "para_up": f"{mod}+up",
        "heading": lambda n: f"{mod}+alt+{min(max(n, 1), 3)}",
        "normal": f"{mod}+shift+n",
        "clear_char": "ctrl+space",
        "align": {"left": f"{mod}+l", "center": f"{mod}+e",
                  "right": f"{mod}+r", "justify": f"{mod}+j"},
        "font_size_box": None if mac else "ctrl+shift+p",
        "list": None,
        "superscript": f"{mod}+shift+equal", "subscript": f"{mod}+equal",
        "indent_more": f"{mod}+m", "indent_less": f"{mod}+shift+m",
    }


class Stopped(Exception):
    """Raised when the user holds the stop key."""


# ============================================================
# PACING
# ============================================================


class Pacer:
    def __init__(self):
        self.dry = False
        self.active = False
        self.factor = 1.0

    def start_dry_run(self):
        self.dry, self.clock, self.kinds = True, 0.0, []

    def end_dry_run(self):
        self.dry = False
        return self.clock, self.kinds

    def begin(self, target_seconds, natural_total, kinds):
        self.active = True
        self.start = monotonic()
        self.deadline = self.start + target_seconds
        self.natural_total = natural_total
        self.natural_done = 0.0
        self.weights = [PAUSE_POINT_WEIGHTS[k] * random.uniform(0.3, 2.0) for k in kinds]
        self.next_point = 0
        self.weight_left = sum(self.weights)
        self.factor = 1.0
        self._update_factor()

    def stop(self):
        self.active = False
        self.factor = 1.0

    def _update_factor(self):
        time_left = self.deadline - monotonic()
        natural_left = max(self.natural_total - self.natural_done, 0.001)
        if time_left >= natural_left:
            self.factor = 1.0
        else:
            self.factor = max(1.0 / MAX_SPEEDUP, max(time_left, 0) / natural_left)

    def natural_delay(self, seconds):
        if self.dry:
            self.clock += seconds
            return 0.0
        if self.active:
            self.natural_done += seconds
            return seconds * self.factor
        return seconds

    def point(self, kind):
        if self.dry:
            self.kinds.append(kind)
            return 0.0
        if not self.active:
            return 0.0
        if self.next_point < len(self.weights):
            weight = self.weights[self.next_point]
        else:
            weight = PAUSE_POINT_WEIGHTS[kind]
        self.next_point += 1

        time_left = self.deadline - monotonic()
        natural_left = max(self.natural_total - self.natural_done, 0.0)
        spare = time_left - natural_left
        extra = spare * weight / self.weight_left if spare > 0 and self.weight_left > 0 else 0.0
        self.weight_left = max(self.weight_left - weight, 0.0)
        self._update_factor()
        return extra

    def status(self):
        if not self.active:
            return ""
        elapsed = monotonic() - self.start
        total = self.deadline - self.start
        return f"[{fmt_duration(elapsed)} of {fmt_duration(total)}]"


PACE = Pacer()


def parse_duration(text):
    text = (text or "").strip().lower().replace(" ", "")
    if not text:
        return None
    if re.fullmatch(r"\d+(\.\d+)?", text):
        return float(text) * 60
    parts = re.findall(r"(\d+(?:\.\d+)?)(hours?|hrs|hr|h|minutes?|mins|min|m|seconds?|secs|sec|s)", text)
    if not parts or "".join(n + u for n, u in parts) != text:
        return None
    seconds = 0.0
    for number, unit in parts:
        seconds += float(number) * {"h": 3600, "m": 60, "s": 1}[unit[0]]
    return seconds if seconds > 0 else None


def fmt_duration(seconds):
    seconds = int(max(seconds, 0))
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    if h:
        return f"{h}h {m:02d}m"
    if m:
        return f"{m}m {s:02d}s"
    return f"{s}s"


# ============================================================
# SAFETY: keep awake, foreground window
# ============================================================


def keep_awake(on):
    if not (WINDOWS and KEEP_PC_AWAKE):
        return
    try:
        ES_CONTINUOUS, ES_SYSTEM, ES_DISPLAY = 0x80000000, 0x00000001, 0x00000002
        flags = ES_CONTINUOUS | ES_SYSTEM | ES_DISPLAY if on else ES_CONTINUOUS
        ctypes.windll.kernel32.SetThreadExecutionState(flags)
    except Exception:
        pass


def foreground_window():
    if not WINDOWS:
        return None, ""
    try:
        user32 = ctypes.windll.user32
        hwnd = user32.GetForegroundWindow()
        length = user32.GetWindowTextLengthW(hwnd)
        buf = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buf, length + 1)
        return hwnd, buf.value
    except Exception:
        return None, ""


CURRENT_PROGRESS = {
    "paragraph": None,
    "total_paragraphs": None,
    "text": "",
    "position": 0,
    "context": "",
}


def set_progress(paragraph=None, total_paragraphs=None, text=None, position=None, context=None):
    if paragraph is not None:
        CURRENT_PROGRESS["paragraph"] = paragraph
    if total_paragraphs is not None:
        CURRENT_PROGRESS["total_paragraphs"] = total_paragraphs
    if text is not None:
        CURRENT_PROGRESS["text"] = text
    if position is not None:
        CURRENT_PROGRESS["position"] = position
    if context is not None:
        CURRENT_PROGRESS["context"] = context


def show_progress(label="Progress"):
    p = CURRENT_PROGRESS
    para = p["paragraph"]
    total = p["total_paragraphs"]
    where = f"paragraph {para}/{total}" if para and total else "current document position"
    context = p["context"] or "typing"
    text = p["text"] or ""
    pos = p["position"]
    preview = text[max(0, pos - FOCUS_STATUS_PREVIEW):pos].replace("\n", " ").strip()
    if len(preview) > FOCUS_STATUS_PREVIEW:
        preview = preview[-FOCUS_STATUS_PREVIEW:]
    print(f"{label}: {where}, {context}, about character {pos}.")
    if preview:
        print(f"Last text: ...{preview!r}")


# ============================================================
# WINDOWS INPUT (ctypes SendInput)
# ============================================================

VK_ESCAPE = 0x1B
MODIFIER_VKS = (0xA0, 0xA1, 0xA2, 0xA3, 0xA4, 0xA5, 0x5B, 0x5C)
EXTENDED_VKS = {0xA3, 0xA5, 0x5B, 0x5C}
KEYEVENTF_EXTENDEDKEY, KEYEVENTF_KEYUP, KEYEVENTF_UNICODE = 0x1, 0x2, 0x4

_WORD, _DWORD, _LONG, _ULONG_PTR = ctypes.c_uint16, ctypes.c_uint32, ctypes.c_int32, ctypes.c_size_t


class _KEYBDINPUT(ctypes.Structure):
    _fields_ = [("wVk", _WORD), ("wScan", _WORD), ("dwFlags", _DWORD),
                ("time", _DWORD), ("dwExtraInfo", _ULONG_PTR)]


class _MOUSEINPUT(ctypes.Structure):
    _fields_ = [("dx", _LONG), ("dy", _LONG), ("mouseData", _DWORD),
                ("dwFlags", _DWORD), ("time", _DWORD), ("dwExtraInfo", _ULONG_PTR)]


class _INPUTUNION(ctypes.Union):
    _fields_ = [("ki", _KEYBDINPUT), ("mi", _MOUSEINPUT)]


class _INPUT(ctypes.Structure):
    _fields_ = [("type", _DWORD), ("u", _INPUTUNION)]


def send_key_events(events):
    if not events:
        return
    if not WINDOWS:
        raise RuntimeError("send_key_events is Windows-only")
    arr = (_INPUT * len(events))()
    for slot, (vk, scan, flags) in zip(arr, events):
        slot.type = 1
        slot.u.ki = _KEYBDINPUT(vk, scan, flags, 0, 0)
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    sent = user32.SendInput(len(events), arr, ctypes.sizeof(_INPUT))
    if sent != len(events):
        raise RuntimeError(
            f"Windows blocked the keystrokes (sent {sent} of {len(events)}, error "
            f"{ctypes.get_last_error()}). If your browser is running as administrator, "
            f"run the console as administrator too (or neither).")


def key_is_down(vk):
    if not WINDOWS:
        return False
    try:
        return bool(ctypes.windll.user32.GetAsyncKeyState(vk) & 0x8000)
    except Exception:
        return False


def modifiers_down():
    return [vk for vk in MODIFIER_VKS if key_is_down(vk)]


def release_modifiers():
    if WINDOWS:
        held = modifiers_down()
        send_key_events([(vk, 0, KEYEVENTF_KEYUP | (KEYEVENTF_EXTENDEDKEY if vk in EXTENDED_VKS else 0))
                         for vk in held])
        return held
    for name in ("ctrl", "shift", "alt"):
        try:
            keyboard.release(name)
        except Exception:
            pass
    return []


def type_unicode(char):
    if WINDOWS:
        units = char.encode("utf-16-le")
        events = []
        for i in range(0, len(units), 2):
            code = int.from_bytes(units[i:i + 2], "little")
            events += [(0, code, KEYEVENTF_UNICODE), (0, code, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP)]
        send_key_events(events)
    else:
        keyboard.write(char, restore_state_after=False)


def stop_pressed():
    if keyboard.is_pressed(STOP_KEY):
        return True
    return STOP_KEY == "escape" and key_is_down(VK_ESCAPE)


# ============================================================
# FIREFOX / SELENIUM BACKEND
# ============================================================


def find_geckodriver():
    """Locate geckodriver on PATH or in common install locations."""
    found = shutil.which("geckodriver")
    if found:
        return found
    candidates = [
        r"C:\Program Files\geckodriver.exe",
        r"C:\Program Files (x86)\geckodriver.exe",
        r"C:\geckodriver.exe",
        str(Path.home() / "geckodriver.exe"),
        str(Path(__file__).parent / "geckodriver.exe"),
        str(Path(__file__).parent / "geckodriver"),
        "/usr/local/bin/geckodriver",
        "/usr/bin/geckodriver",
        "/opt/homebrew/bin/geckodriver",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def is_firefox_running():
    """Is any firefox process running?"""
    try:
        if WINDOWS:
            result = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq firefox.exe", "/NH"],
                capture_output=True, text=True, timeout=5,
            )
            return "firefox.exe" in result.stdout.lower()
        else:
            result = subprocess.run(["pgrep", "-x", "firefox"], capture_output=True, text=True)
            return result.returncode == 0
    except Exception:
        return False


def kill_firefox():
    """Force-close every Firefox process. Returns True if none are running
    after the attempt."""
    try:
        if WINDOWS:
            subprocess.run(["taskkill", "/IM", "firefox.exe", "/F"],
                           capture_output=True, check=False)
        else:
            subprocess.run(["pkill", "-x", "firefox"], check=False)
    except Exception:
        pass
    for _ in range(20):
        if not is_firefox_running():
            return True
        sleep(0.5)
    return not is_firefox_running()


# What is NOT copied out of the real profile into the temporary one.
#
# The copy only needs your Google cookies (cookies.sqlite) so Docs opens
# signed in.  Everything below is left behind on purpose:
#
#   Firefox Account / Sync sign-in  (signedInUser.json, weave/)
#       CRITICAL.  The temporary Firefox runs with Mozilla's automation
#       preferences (password saving off, blank new tab, Safe Browsing off).
#       If the copy were signed into your Firefox Account, Sync would upload
#       those values and your OTHER profiles would download them.  Without
#       the sign-in file the copy is not signed in, so it cannot sync.
#
#   Saved passwords  (logins.json, key4.db, ...)
#       Not needed -- Google sign-in lives in cookies -- and leaving them out
#       means a leftover temp folder can never contain your passwords.
#
#   Locks, caches, crash data
#       Locks would stop the copy opening; the rest is just size.
_PROFILE_SKIP_DIRS = {
    "cache2", "startupCache", "shader-cache", "thumbnails",
    "minidumps", "crashes", "safebrowsing", "settings",
    "weave",                               # Sync state
}
_PROFILE_SKIP_FILES = {
    "parent.lock", ".parentlock", "lock",
    "signedInUser.json",                   # Firefox Account / Sync sign-in
    "logins.json", "logins-backup.json",   # saved passwords
    "key4.db", "key3.db", "signons.sqlite",
    "user.js",                             # never inherit stale overrides
}

# Prefix every temporary profile uses.  launch_firefox_driver() refuses to
# launch anything that is not one of these, so a real profile can never be
# handed to geckodriver by mistake.
_PROFILE_COPY_PREFIX = "dripwriter-prof-"


# Google editor paths that take a /u/<n>/ account segment.
_GOOGLE_EDITORS = {"document", "spreadsheets", "presentation",
                   "drawings", "forms", "file"}


_ACCOUNT_INDEX_UNSET = object()


def apply_google_account_index(url, index=_ACCOUNT_INDEX_UNSET):
    """Force a Google Docs link to open as a specific signed-in account.

    Google numbers the accounts in a browser profile: /u/0/ is the one you
    signed in with first, /u/1/ the next, and so on.  A link without that
    segment opens as whichever account is the default, which may not be the
    one that has access to the document.

    Given any of these, with index 1:
        .../document/d/FILEID/edit?tab=t.0#heading=h.abc
        .../document/u/3/d/FILEID/edit          (wrong account -> corrected)
        .../u/3/document/d/FILEID/edit          (older link shape)
    this returns:
        .../document/u/1/d/FILEID/edit?tab=t.0#heading=h.abc

    The query string and the #heading fragment are preserved untouched.
    Non-Google URLs, and index=None, are returned unchanged.
    """
    if index is _ACCOUNT_INDEX_UNSET:
        index = FIREFOX_ACCOUNT_INDEX
    if index is None or str(index).strip() == "":
        return url
    try:
        index = int(index)
    except (TypeError, ValueError):
        return url

    try:
        from urllib.parse import urlsplit, urlunsplit
        parts = urlsplit(url)
    except Exception:
        return url

    if "google." not in (parts.netloc or "").lower():
        return url

    segments = [s for s in parts.path.split("/") if s != ""]

    # Drop any account segment the link already carries, wherever it sits.
    cleaned = []
    i = 0
    while i < len(segments):
        if (segments[i] == "u" and i + 1 < len(segments)
                and segments[i + 1].isdigit()):
            i += 2
            continue
        cleaned.append(segments[i])
        i += 1

    # Re-insert it directly after the editor segment (/document/u/1/d/...).
    placed = False
    for pos, segment in enumerate(cleaned):
        if segment in _GOOGLE_EDITORS:
            cleaned[pos + 1:pos + 1] = ["u", str(index)]
            placed = True
            break
    if not placed:
        cleaned[0:0] = ["u", str(index)]

    new_path = "/" + "/".join(cleaned)
    if parts.path.endswith("/") and not new_path.endswith("/"):
        new_path += "/"
    return urlunsplit((parts.scheme, parts.netloc, new_path,
                       parts.query, parts.fragment))


def is_dripwriter_profile_copy(path):
    """True only for a temporary profile this script created."""
    import tempfile
    try:
        path = Path(path).resolve()
        temp_root = Path(tempfile.gettempdir()).resolve()
    except Exception:
        return False
    return (path.name.startswith(_PROFILE_COPY_PREFIX)
            and path.parent == temp_root)


def cleanup_stale_profile_copies():
    """Delete temp profiles left behind by runs that crashed or were killed.

    They hold Google cookies, so they shouldn't sit around in Temp.
    """
    import tempfile
    try:
        for leftover in Path(tempfile.gettempdir()).glob(_PROFILE_COPY_PREFIX + "*"):
            if leftover.is_dir() and leftover != FIREFOX_PROFILE_COPY:
                shutil.rmtree(leftover, ignore_errors=True)
    except Exception:
        pass


def copy_firefox_profile(src):
    """Copy a profile to a temp dir so we can launch it while the original
    stays open in your normal Firefox.

    The source profile is only ever READ.  Only what Docs needs comes along
    (cookies, site storage); the Firefox Account sign-in, saved passwords,
    locks and caches are left behind -- see _PROFILE_SKIP_FILES for why the
    Firefox Account part matters.
    """
    import tempfile
    dest = Path(tempfile.mkdtemp(prefix=_PROFILE_COPY_PREFIX))

    def ignore(directory, names):
        skipped = set()
        for name in names:
            if name in _PROFILE_SKIP_FILES:
                skipped.add(name)
            elif name in _PROFILE_SKIP_DIRS and (Path(directory) / name).is_dir():
                skipped.add(name)
        return skipped

    shutil.copytree(src, dest, ignore=ignore, dirs_exist_ok=True)
    return dest


def launch_firefox_driver(profile_dir, headless=None):
    """Let geckodriver launch Firefox itself, pointed at *profile_dir*.

    This replaces the old "we launch Firefox, then attach with
    --connect-existing" approach, which had two fatal problems:

      1. Launching `firefox -P "Firefox 2"` while another Firefox was
         running did NOT open that profile.  Firefox's remoting handed the
         command line to the already-running instance, which just opened a
         new window in ITS OWN profile and ignored -P (bugzilla 1367743,
         "just opens a new window in the same instance, using the currently
         open profile").  So the profile we wrote prefs into was never the
         profile that ran.
      2. With --connect-existing, geckodriver never creates a profile, so it
         never applies its recommended prefs -- including the one that fixes
         background typing -- and every options.set_preference() call was
         silently discarded.

    Letting geckodriver launch Firefox fixes both: it passes an absolute
    -profile path (no remoting ambiguity) and applies focusmanager.testmode
    itself.  We set the prefs explicitly too, so it does not matter which
    geckodriver version is installed.
    """
    global FIREFOX_DRIVER, GECKODRIVER_PATH
    if not SELENIUM_AVAILABLE:
        print("Selenium is not installed. Run: python -m pip install selenium")
        return False
    if GECKODRIVER_PATH is None:
        GECKODRIVER_PATH = find_geckodriver()

    headless = FIREFOX_HEADLESS if headless is None else headless

    # HARD GUARD.  geckodriver writes its automation preferences into whatever
    # profile it is given.  Only ever hand it a temporary copy we made, never
    # a real profile, no matter how this function is called.
    if not is_dripwriter_profile_copy(profile_dir):
        print("REFUSING to launch: that is not a temporary profile copy.")
        print(f"  {profile_dir}")
        print("  Automation would write test settings into it. Nothing was done.")
        return False

    options = FirefoxOptions()
    options.add_argument("-profile")
    options.add_argument(str(profile_dir))
    if headless:
        options.add_argument("-headless")
    # These DO take effect now that geckodriver owns the launch.
    for key, value in FIREFOX_BACKGROUND_PREFS.items():
        options.set_preference(key, value)
    # Belt and braces: even if a Firefox Account sign-in somehow reached the
    # copy, make sure it can never sync anything to your other profiles.
    for key, value in FIREFOX_NO_SYNC_PREFS.items():
        options.set_preference(key, value)

    # If geckodriver isn't on PATH, leave executable_path unset so Selenium
    # Manager (Selenium 4.6+) downloads a matching geckodriver by itself.
    kwargs = {}
    if GECKODRIVER_PATH:
        kwargs["executable_path"] = GECKODRIVER_PATH
    else:
        print("  (geckodriver not on PATH — Selenium Manager will fetch it)")

    try:
        service = FirefoxService(**kwargs)
        FIREFOX_DRIVER = webdriver.Firefox(service=service, options=options)
        return True
    except Exception as exc:
        print(f"Could not launch Firefox: {exc}")
        print("If this mentions geckodriver, update Selenium:")
        print("  python -m pip install -U selenium")
        FIREFOX_DRIVER = None
        return False


# ============================================================
# BACKGROUND-TYPING PREFS  (the actual fix)
# ============================================================
#
# Why this exists
# ---------------
# Gecko has two separate notions of focus:
#   * the OS window focus, and
#   * nsFocusManager's internal "active window" + "focused element".
#
# When Windows tells Firefox its window was lowered (another window came on
# top), nsFocusManager::WindowLowered() clears the active window and fires a
# blur.  Google Docs captures keystrokes on a hidden iframe
# (iframe.docs-texteventtarget-iframe); once the focus manager drops the
# focused element, WebDriver keystrokes no longer route to that iframe, so
# typing stops.  It resumes the instant you click back, because the focus
# manager restores the stored focus.  No amount of JavaScript can fix this:
# overriding document.hasFocus() does not put the focused element back.
#
# Gecko ships a pref for precisely this.  geckodriver sets it by default,
# with this comment in testing/geckodriver/src/prefs.rs:
#
#     // Allow the application to have focus even it runs in the background
#     ("focusmanager.testmode", Pref::new(true)),
#
# With focusmanager.testmode = true, WindowLowered() skips lowering the
# window, so the active state and the focused element survive (Firefox 63+,
# bugzilla 1489967: "Keep active state when Firefox is moved to background
# and focusmanager test mode is enabled").
#
# These prefs are passed to geckodriver, which applies them to the temporary
# profile COPY it launches.  Nothing is ever written to one of your real
# Firefox profiles.
#
# (An earlier version did write a user.js into the real profile and started
# Firefox with --marionette.  Don't do that: Firefox applies a block of
# automation "recommended preferences" when it runs under Marionette, and
# they can persist in the profile afterwards -- disabling password autofill,
# blanking the homepage and turning off Safe Browsing.  fox_repair.py undoes
# that damage if it was ever done.)

FIREFOX_BACKGROUND_PREFS = {
    # THE fix: keep Gecko's focus state when the window is backgrounded.
    "focusmanager.testmode": True,

    # Don't pause/throttle the tab when the window is covered or occluded.
    "widget.windows.window_occlusion_tracking.enabled": False,
    "dom.min_background_timeout_value": 4,
    "dom.min_background_timeout_value_without_budget_throttling": 4,
    "dom.timeout.enable_budget_timer_throttling": False,
    "dom.suspend_inactive.enabled": False,
    "dom.animations.offscreen-throttling": False,
}

# Applied only to the temporary copy.  Firefox Sync carries exactly the
# settings automation changes (signon.rememberSignons, signon.autofillForms,
# browser.newtabpage.enabled, browser.safebrowsing.*, browser.startup.page)
# between every profile signed into the same Firefox Account -- see the
# services.sync.prefs.sync.* list in browser/app/profile/firefox.js.  The copy
# must never sync, or it would push automation values to your real profiles.
FIREFOX_NO_SYNC_PREFS = {
    "identity.fxaccounts.enabled": False,
    "services.sync.engine.prefs": False,
    "services.sync.engine.passwords": False,
    "services.sync.engine.bookmarks": False,
    "services.sync.engine.history": False,
    "services.sync.engine.tabs": False,
    "services.sync.engine.addons": False,
    "services.sync.engine.addresses": False,
    "services.sync.engine.creditcards": False,
    "services.sync.autoconnect": False,
}

def firefox_profiles_ini():
    """Return the Path to Firefox's profiles.ini, or None."""
    if WINDOWS:
        base = Path(os.environ.get("APPDATA", "")) / "Mozilla" / "Firefox"
    elif MAC:
        base = Path.home() / "Library" / "Application Support" / "Firefox"
    else:
        base = Path.home() / ".mozilla" / "firefox"
    ini = base / "profiles.ini"
    return ini if ini.is_file() else None


def find_firefox_profile_dir(profile_name):
    """Look up a named Firefox profile's directory via profiles.ini."""
    ini_path = firefox_profiles_ini()
    if ini_path is None:
        return None
    parser = configparser.ConfigParser()
    try:
        parser.read(ini_path, encoding="utf-8")
    except Exception:
        return None

    root = ini_path.parent
    for section in parser.sections():
        if not section.lower().startswith("profile"):
            continue
        if parser.get(section, "Name", fallback=None) != profile_name:
            continue
        rel_path = parser.get(section, "Path", fallback=None)
        if not rel_path:
            return None
        is_relative = parser.get(section, "IsRelative", fallback="1").strip() == "1"
        # profiles.ini always stores Path with forward slashes
        path = (root / rel_path.replace("/", os.sep)) if is_relative else Path(rel_path)
        return path if path.is_dir() else None
    return None


def firefox_runtime_report(driver):
    """Report what Firefox is ACTUALLY running: version, headless, the real
    profile directory, and whether the background-typing pref is live.

    The profile path matters because launching `firefox -P <name>` while
    another Firefox is open can silently open a DIFFERENT profile.  Reading
    it back is the only way to be sure.
    """
    info = {"version": None, "headless": None, "profile": None,
            "testmode": None, "error": None}

    try:
        driver.set_context("chrome")
    except Exception as exc:
        info["error"] = f"chrome context unavailable: {type(exc).__name__}"
        _runtime_report_fallback(driver, info)
        return info

    probes = {
        "version": "return Services.appinfo.version;",
        "headless": "return Services.appinfo.headless;",
        # Ci is not always defined in the chrome sandbox; use the string form.
        "profile": "return Services.dirsvc.get('ProfD', "
                   "Components.interfaces.nsIFile).path;",
        "testmode": "return Services.prefs.getBoolPref("
                    "'focusmanager.testmode', false);",
    }
    for key, script in probes.items():
        try:
            info[key] = driver.execute_script(script)
        except Exception as exc:
            if info["error"] is None:
                info["error"] = f"{key}: {type(exc).__name__}"

    try:
        driver.set_context("content")
    except Exception:
        pass

    if info["version"] is None:
        _runtime_report_fallback(driver, info)
    return info


def _runtime_report_fallback(driver, info):
    """If the chrome context is unavailable, get what we can from content."""
    try:
        ua = driver.execute_script("return navigator.userAgent;") or ""
        marker = "Firefox/"
        if marker in ua:
            info["version"] = ua.split(marker, 1)[1].split(" ")[0] + " (from UA)"
    except Exception:
        pass


def cleanup_firefox_profile_copy():
    """Delete the temporary profile copy, if we made one."""
    global FIREFOX_PROFILE_COPY
    if FIREFOX_PROFILE_COPY:
        shutil.rmtree(FIREFOX_PROFILE_COPY, ignore_errors=True)
        FIREFOX_PROFILE_COPY = None


def list_firefox_tabs(driver):
    """Return [(handle, title, url)] for every open tab."""
    tabs = []
    try:
        current = driver.current_window_handle
    except Exception:
        current = None
    for h in driver.window_handles:
        try:
            driver.switch_to.window(h)
            tabs.append((h, driver.title, driver.current_url))
        except Exception:
            tabs.append((h, "?", "?"))
    if current is not None:
        try:
            driver.switch_to.window(current)
        except Exception:
            pass
    return tabs


def focus_google_docs_tab(driver):
    """Switch to a tab that looks like Google Docs. Returns True if found."""
    for h, title, url in list_firefox_tabs(driver):
        if "docs.google.com" in url or "Google Docs" in title:
            driver.switch_to.window(h)
            print(f"Switched to Google Docs tab: {title[:70]!r}")
            return True
    return False


def firefox_click_body(driver):
    """Click inside the Google Docs editor so the cursor is active."""
    # Google Docs' editable area is inside a specific element; try several
    # selectors from most-specific to most-generic.
    selectors = [
        "div.kix-appview-editor",           # main Docs editor pane
        "div.kix-page",                     # a page inside the editor
        "canvas.kix-canvas-tile-content",   # canvas-based editor (newer Docs)
        "iframe.docs-texteventtarget",      # hidden event-capture iframe
        "body",                             # fallback
    ]
    for sel in selectors:
        try:
            elem = driver.find_element("css selector", sel)
            ActionChains(driver).move_to_element(elem).click().perform()
            return
        except Exception:
            continue
    print("  [Could not auto-focus the editor — click inside the Doc yourself.]")


def firefox_inject_focus_lock(driver):
    """Trick Google Docs into believing the window always has focus.

    Google Docs checks focus through several channels:
      1. document.hasFocus()  — polled by Docs' editor loop
      2. Page Visibility API  — document.hidden / visibilityState
      3. 'blur' events on window, document, and the hidden text-event iframe
      4. 'visibilitychange' events on the document
      5. 'focusout' events on the editor and the hidden iframe

    We override ALL of these on both the main document AND inside Google Docs'
    hidden text-event iframe (docs-texteventtarget-iframe), which is where Docs
    actually captures keyboard input.

    Safe to call repeatedly (idempotent overrides + one-time event hooks).
    """
    script = r"""
    (function() {
        // ---------- helper: lock one document/window pair ----------
        function lockContext(doc, win, label) {
            // hasFocus() → always true (on the prototype so it covers all docs)
            if (doc.constructor && doc.constructor.prototype) {
                doc.constructor.prototype.hasFocus = function() { return true; };
            }
            // Also override the instance directly in case Docs caches the ref
            doc.hasFocus = function() { return true; };

            // Page Visibility API
            try {
                Object.defineProperty(doc, 'hidden', {
                    get: function() { return false; }, configurable: true
                });
            } catch(e) {}
            try {
                Object.defineProperty(doc, 'visibilityState', {
                    get: function() { return 'visible'; }, configurable: true
                });
            } catch(e) {}

            // One-time event captures
            var flag = '__dripwriterLock';
            if (!win[flag]) {
                win[flag] = true;

                // Swallow blur on window
                win.addEventListener('blur', function(e) {
                    e.stopImmediatePropagation();
                    e.preventDefault();
                }, true);

                // Swallow focusout on document (Docs listens for this too)
                doc.addEventListener('focusout', function(e) {
                    e.stopImmediatePropagation();
                    e.preventDefault();
                }, true);

                // Swallow visibilitychange
                doc.addEventListener('visibilitychange', function(e) {
                    e.stopImmediatePropagation();
                    e.preventDefault();
                }, true);

                // Swallow blur on document
                doc.addEventListener('blur', function(e) {
                    e.stopImmediatePropagation();
                }, true);

                console.log('[DripWriter] Focus lock active on ' + label);
            }
        }

        // ---------- lock the MAIN document ----------
        lockContext(document, window, 'main');

        // ---------- lock the hidden text-event iframe ----------
        var iframes = document.querySelectorAll(
            'iframe.docs-texteventtarget-iframe, iframe.docs-texteventtarget'
        );
        for (var i = 0; i < iframes.length; i++) {
            try {
                var idoc = iframes[i].contentDocument;
                var iwin = iframes[i].contentWindow;
                if (idoc && iwin) {
                    lockContext(idoc, iwin, 'iframe#' + i);
                }
            } catch(e) {
                // cross-origin — can't access, but shouldn't happen for same-origin Docs
            }
        }

        // ---------- also lock every future iframe Docs might create ----------
        if (!window.__dripwriterObserver) {
            window.__dripwriterObserver = new MutationObserver(function(mutations) {
                for (var m = 0; m < mutations.length; m++) {
                    for (var n = 0; n < mutations[m].addedNodes.length; n++) {
                        var node = mutations[m].addedNodes[n];
                        if (node.tagName === 'IFRAME' &&
                            (node.className || '').indexOf('docs-texteventtarget') !== -1) {
                            try {
                                var d = node.contentDocument, w = node.contentWindow;
                                if (d && w) lockContext(d, w, 'new-iframe');
                            } catch(e) {}
                        }
                    }
                }
            });
            window.__dripwriterObserver.observe(document.documentElement, {
                childList: true, subtree: true
            });
        }
    })();
    """
    try:
        driver.execute_script(script)
        _FOCUS_LOCK_LAST[0] = monotonic()
        return True
    except Exception as exc:
        print(f"  [Warning: could not inject focus lock: {exc}]")
        return False


def choose_input_backend():
    """Ask at launch which input path to use; Firefox is optional."""
    global INPUT_BACKEND, FIREFOX_DRIVER

    print("\nChoose input method:")
    print("  1. Windows")
    print("  2. Firefox (Background)")
    while True:
        choice = input("[1/2, default 1]: ").strip().lower() or "1"
        if choice in {"1", "windows", "keystrokes", "w"}:
            INPUT_BACKEND = "windows"
            break
        if choice in {"2", "firefox", "f"}:
            if not SELENIUM_AVAILABLE:
                print("Selenium is not installed. Install it with: python -m pip install selenium")
                continue
            INPUT_BACKEND = "firefox"
            break
        print("Please enter 1 or 2.")

    if INPUT_BACKEND != "firefox":
        return

    # ----------------------------------------------------------------
    # Firefox (Background)
    #
    # We let geckodriver launch Firefox against a COPY of the profile,
    # rather than launching Firefox ourselves and attaching.  See
    # launch_firefox_driver() for why the old approach could never work.
    # ----------------------------------------------------------------
    global FIREFOX_PROFILE_COPY

    source_profile = find_firefox_profile_dir(FIREFOX_PROFILE_NAME)
    if source_profile is None:
        print(f"\nCould not find a Firefox profile named {FIREFOX_PROFILE_NAME!r}.")
        print("Open about:profiles in Firefox to check the exact name, then set")
        print("FIREFOX_PROFILE_NAME at the top of this script.")
        print("Falling back to Windows keystrokes.")
        INPUT_BACKEND = "windows"
        return

    doc_url = FIREFOX_DOC_URL.strip()
    if not doc_url:
        print("\nPaste the URL of the Google Doc to type into.")
        print("(It must already be signed in on the "
              f"'{FIREFOX_PROFILE_NAME}' profile.)")
        try:
            doc_url = input("Doc URL: ").strip()
        except EOFError:
            doc_url = ""
    if not doc_url:
        print("No URL given. Falling back to Windows keystrokes.")
        INPUT_BACKEND = "windows"
        return

    # Pin the link to a specific signed-in Google account (/u/<n>/).
    resolved_url = apply_google_account_index(doc_url)
    if resolved_url != doc_url:
        print(f"\nUsing Google account {FIREFOX_ACCOUNT_INDEX} "
              f"(FIREFOX_ACCOUNT_INDEX at the top of the script):")
        print(f"  {resolved_url}")
    doc_url = resolved_url

    cleanup_stale_profile_copies()
    print(f"\nCopying the '{FIREFOX_PROFILE_NAME}' profile "
          "(so your own Firefox stays usable)...")
    try:
        FIREFOX_PROFILE_COPY = copy_firefox_profile(source_profile)
    except Exception as exc:
        print(f"Could not copy the profile: {exc}")
        print("Falling back to Windows keystrokes.")
        INPUT_BACKEND = "windows"
        return

    print("Launching Firefox"
          f"{' (headless — no window)' if FIREFOX_HEADLESS else ''}...")
    if not launch_firefox_driver(FIREFOX_PROFILE_COPY):
        print("Falling back to Windows keystrokes.")
        cleanup_firefox_profile_copy()
        INPUT_BACKEND = "windows"
        return

    driver = FIREFOX_DRIVER

    # --- prove the configuration is what we think it is ---
    info = firefox_runtime_report(driver)
    print("\n--- Firefox runtime ---")
    print(f"  version               : {info.get('version')}")
    print(f"  profile in use        : {info.get('profile') or '(not readable)'}")
    print(f"  focusmanager.testmode : {info.get('testmode')}")
    if info.get("error"):
        print(f"  (pref readout unavailable: {info['error']} — harmless)")

    if FIREFOX_HEADLESS:
        print("\n[OK] Headless: there is no window, so no other window can cover")
        print("     it and no focus can be lost. Background typing is safe.")
    elif info.get("testmode") is True:
        print("\n[OK] focusmanager.testmode is ON — typing should survive other")
        print("     windows being placed on top.")
    elif info.get("testmode") is False:
        print("\n[!!] focusmanager.testmode is OFF. Typing may pause when covered.")
        print("     Set FIREFOX_HEADLESS = True at the top of the script.")
    else:
        print("\n[ ] Could not read the pref, but geckodriver sets it by default")
        print("    when it launches Firefox (which it did). If typing pauses when")
        print("    covered, set FIREFOX_HEADLESS = True at the top of the script.")

    # --- open the doc ---
    print(f"\nOpening the doc...")
    try:
        driver.get(doc_url)
        sleep(4)
    except Exception as exc:
        print(f"Could not open {doc_url}: {exc}")

    title = ""
    landed_url = ""
    body_text = ""
    try:
        title = driver.title or ""
        landed_url = driver.current_url or ""
        print(f"  page title: {title[:70]!r}")
    except Exception:
        pass
    try:
        body_text = (driver.find_element("css selector", "body").text or "")[:600]
    except Exception:
        pass

    low_title, low_url, low_body = title.lower(), landed_url.lower(), body_text.lower()

    # Not signed in at all.
    if "sign in" in low_title or "accounts.google.com" in low_url:
        print("\n[!!] That page is asking you to sign in, so the profile copy")
        print("     does not have a live Google session.")
        print(f"     Fix: open Firefox normally on the '{FIREFOX_PROFILE_NAME}'")
        print("     profile, sign into Google, close it, then rerun this script.")

    # Signed in, but as an account that cannot open this document.
    elif ("request access" in low_body or "you need access" in low_body
            or "request edit access" in low_body
            or "don't have access" in low_body
            or "denied" in low_title):
        print("\n[!!] This Google account cannot open that document.")
        print(f"     FIREFOX_ACCOUNT_INDEX is currently {FIREFOX_ACCOUNT_INDEX}, "
              "which selects")
        print(f"     the /u/{FIREFOX_ACCOUNT_INDEX}/ account on the "
              f"'{FIREFOX_PROFILE_NAME}' profile.")
        print("     Fix: open the doc in your normal Firefox as the account that")
        print("     DOES have access, look at the /u/<n>/ number in the address")
        print("     bar, and set FIREFOX_ACCOUNT_INDEX to that number.")
        print("     Typing now would do nothing, so stop here and fix it first.")

    # Looks like the editor really is open.
    elif "google docs" in low_title or "/edit" in low_url:
        print("  [OK] Document is open and editable.")

    try:
        firefox_click_body(driver)
    except Exception:
        pass
    firefox_inject_focus_lock(driver)
    print("\nReady. Press the start key to begin.")


def _selenium_key(name):
    """Translate keyboard-library key names into Selenium key constants."""
    names = {
        "enter": Keys.ENTER, "return": Keys.ENTER, "tab": Keys.TAB,
        "backspace": Keys.BACKSPACE, "delete": Keys.DELETE, "escape": Keys.ESCAPE,
        "esc": Keys.ESCAPE, "up": Keys.ARROW_UP, "down": Keys.ARROW_DOWN,
        "left": Keys.ARROW_LEFT, "right": Keys.ARROW_RIGHT, "home": Keys.HOME,
        "end": Keys.END, "pageup": Keys.PAGE_UP, "pagedown": Keys.PAGE_DOWN,
        "space": Keys.SPACE, "period": ".", "comma": ",",
        "equal": "=", "[": "[", "]": "]", "\\": "\\",
        "ctrl": Keys.CONTROL, "control": Keys.CONTROL, "alt": Keys.ALT,
        "shift": Keys.SHIFT, "command": Keys.COMMAND, "win": Keys.META,
    }
    return names.get(name.lower(), name)


def firefox_alive():
    """Return True if the WebDriver session is still usable."""
    if FIREFOX_DRIVER is None:
        return False
    try:
        _ = FIREFOX_DRIVER.title
        return True
    except Exception:
        return False


_FOCUS_LOCK_LAST = [0.0]          # monotonic timestamp of last re-injection
_FOCUS_LOCK_INTERVAL = 15         # seconds between re-injections

def _maybe_reinject_focus_lock():
    """Re-inject the focus lock periodically so Docs can't reclaim blur."""
    now = monotonic()
    if now - _FOCUS_LOCK_LAST[0] >= _FOCUS_LOCK_INTERVAL:
        if FIREFOX_DRIVER is not None:
            firefox_inject_focus_lock(FIREFOX_DRIVER)
        _FOCUS_LOCK_LAST[0] = now


def firefox_send(hotkey):
    """Send one keyboard shortcut through WebDriver ActionChains.

    ActionChains uses Firefox's nsITextInputProcessor, which produces
    isTrusted:true events.  The focus-lock injection prevents Google Docs
    from noticing that the OS window lost focus.
    """
    if FIREFOX_DRIVER is None:
        raise RuntimeError("Firefox backend is selected, but its WebDriver is unavailable.")
    parts = str(hotkey).replace(" ", "").lower().split("+")
    modifiers = []
    while len(parts) > 1:
        part = parts.pop(0)
        modifiers.append(_selenium_key(part))
    key = _selenium_key(parts[0])
    try:
        actions = ActionChains(FIREFOX_DRIVER)
        for modifier in modifiers:
            actions.key_down(modifier)
        actions.send_keys(key)
        for modifier in reversed(modifiers):
            actions.key_up(modifier)
        actions.perform()
    except Exception as exc:
        if not firefox_alive():
            raise RuntimeError(
                "Firefox was closed or crashed. Restart the script to reconnect.") from exc
        raise


def firefox_type_text(text):
    """Send a whole string through WebDriver ActionChains.

    Uses isTrusted:true events via Marionette.  The focus-lock injection
    keeps Google Docs from pausing when the window is in the background.
    """
    if FIREFOX_DRIVER is None:
        raise RuntimeError("Firefox WebDriver is not available.")
    _maybe_reinject_focus_lock()
    try:
        ActionChains(FIREFOX_DRIVER).send_keys(text).perform()
    except Exception as exc:
        if not firefox_alive():
            raise RuntimeError(
                "Firefox was closed or crashed. Restart the script to reconnect.") from exc
        raise


def firefox_type_char(char):
    firefox_type_text(char)


# ============================================================
# FOCUS GUARD (Windows global-keystroke mode only)
# ============================================================


class FocusGuard:
    def __init__(self):
        self.enabled = False
        self.hwnd = None
        self.lock = ""
        self.first_line = ""

    def _title_lock(self, title):
        need = FOCUS_TITLE_MUST_CONTAIN.get(EDITOR, "")
        if not need:
            return ""
        if FOCUS_LOCK_TO_THIS_DOC and need in title:
            return title[:title.index(need) + len(need)]
        return need

    def arm(self):
        # Only meaningful for global keystroke injection.
        self.enabled = (FOCUS_GUARD and WINDOWS and INPUT_BACKEND == "windows")
        if not self.enabled:
            return
        need = FOCUS_TITLE_MUST_CONTAIN.get(EDITOR, "")
        warned = False
        while True:
            hwnd, title = foreground_window()
            if hwnd and (not need or need in title):
                break
            if not warned:
                print(f"Waiting for a {need or 'document'} window to be in front "
                      f"(hold {STOP_KEY.upper()} to cancel)...")
                warned = True
            if stop_pressed():
                raise Stopped
            sleep(0.1)
        self.hwnd, self.lock = hwnd, self._title_lock(title)
        print(f"Locked to: {title[:80]!r}")

    def disarm(self):
        self.enabled = False

    def ok(self):
        hwnd, title = foreground_window()
        if not hwnd or hwnd != self.hwnd:
            return False
        if not self.lock.endswith(" - Google Docs"):
            return self.lock in title
        if title.startswith(self.lock):
            return True
        if " - Google Docs" in title:
            new_name = title.split(" - Google Docs")[0].strip().rstrip(".… ")
            first = " ".join(self.first_line.split())
            if new_name and first.startswith(new_name):
                self.lock = title.split(" - Google Docs")[0].strip() + " - Google Docs"
                print(f"[Docs renamed the document to {new_name!r}; still typing into it.]")
                return True
        return False

    def before_key(self):
        if not self.enabled:
            return
        if not self.ok():
            self._paused()
        if modifiers_down():
            end = monotonic() + 1.0
            while modifiers_down() and monotonic() < end:
                if stop_pressed():
                    raise Stopped
                sleep(0.02)
            if release_modifiers():
                print("[Released a held Ctrl/Alt/Shift/Win key before typing.]")
            if not self.ok():
                self._paused()

    def after_key(self):
        if self.enabled and not self.ok():
            release_modifiers()

    def _paused(self):
        release_modifiers()
        print("\n========== FOCUS LOST - TYPING PAUSED ==========")
        print("The document isn't in front, so nothing is being typed.")
        print("To continue: switch back to the same document (Alt+Tab, or click its tab).")
        print(f"It resumes {FOCUS_RESUME_AFTER:g}s after the document is back. "
              f"Hold {STOP_KEY.upper()} to stop instead.")
        print("Don't click inside the page: that moves the typing position.")
        print(f"Expected window: {self.lock!r}")
        print(f"Window in front: {foreground_window()[1][:80]!r}")
        show_progress("Stopped at")
        print("================================================\n")
        back_since = None
        while True:
            if stop_pressed():
                raise Stopped
            if self.ok() and not modifiers_down():
                back_since = back_since or monotonic()
                if monotonic() - back_since >= FOCUS_RESUME_AFTER:
                    print("[Document is back in front - resuming.]\n")
                    return
            else:
                back_since = None
            sleep(0.05)


FOCUS = FocusGuard()


# ============================================================
# LOW-LEVEL KEY HELPERS
# ============================================================


def check_stop():
    if not PACE.dry and stop_pressed():
        raise Stopped


def wait(seconds):
    end = monotonic() + seconds
    while True:
        check_stop()
        left = end - monotonic()
        if left <= 0:
            return
        sleep(min(left, 0.05))


def pause(seconds):
    if TEST_MODE and not PACE.dry:
        return
    real = PACE.natural_delay(seconds)
    if real > 0:
        wait(real)


def pause_point(kind):
    extra = PACE.point(kind)
    if extra <= 0:
        return
    if ANNOUNCE_LONG_BREAKS and extra >= 60:
        print(f"  [Break: {fmt_duration(extra)}] {PACE.status()}")
    wait(extra)


KEYS_SENT = [0]


def test_mode_gap(seconds):
    if not TEST_MODE:
        return
    KEYS_SENT[0] += 1
    sleep(seconds)
    if TEST_MODE_CATCH_UP_EVERY and KEYS_SENT[0] % TEST_MODE_CATCH_UP_EVERY == 0:
        wait(TEST_MODE_CATCH_UP_PAUSE)


BLOCKED_HOTKEYS = {"ctrl+f", "control+f", "command+f", "ctrl+w", "ctrl+n", "ctrl+t",
                   "alt+f4", "ctrl+shift+w", "ctrl+q", "ctrl+p", "ctrl+o", "ctrl+s"}


def send(hotkey):
    if PACE.dry:
        return
    normalized = str(hotkey).replace(" ", "").lower()
    if normalized in BLOCKED_HOTKEYS:
        raise RuntimeError(f"Safety lock: {hotkey} is never sent by DripWriter.")
    check_stop()
    FOCUS.before_key()
    if INPUT_BACKEND == "firefox":
        firefox_send(hotkey)
    else:
        keyboard.send(hotkey)
        if "+" in normalized:
            release_modifiers()
    FOCUS.after_key()
    test_mode_gap(TEST_MODE_SHORTCUT_GAP)


def shortcut(hotkey):
    if not hotkey:
        return
    pause(random.uniform(*DOCX_FORMAT_PAUSE))
    send(hotkey)
    pause(random.uniform(*DOCX_FORMAT_PAUSE))


def type_char(char):
    if PACE.dry or char == "\r":
        return
    if char in "\n\t\f":
        send({"\n": "enter", "\t": "tab", "\f": f"{MOD}+enter"}[char])
        return
    check_stop()
    FOCUS.before_key()
    if INPUT_BACKEND == "firefox":
        firefox_type_char(char)
    else:
        type_unicode(char)
    FOCUS.after_key()
    test_mode_gap(TEST_MODE_KEY_GAP)


def type_literal(text):
    """Type text exactly. Firefox mode batches in chunks for speed;
    Windows mode goes char-by-char to keep the focus guard tight."""
    if INPUT_BACKEND == "firefox" and text and not PACE.dry:
        chunk = 40
        for i in range(0, len(text), chunk):
            check_stop()
            firefox_type_text(text[i:i + chunk])
            test_mode_gap(TEST_MODE_KEY_GAP)
        return
    for char in text:
        type_char(char)


def typo_for(char):
    near = NEIGHBOURS.get(char.lower())
    if not near:
        return None
    wrong = random.choice(near)
    return wrong.upper() if char.isupper() else wrong


def say(*args, **kwargs):
    if not PACE.dry:
        print(*args, **kwargs)


def beep():
    if not PLAY_SOUND:
        return
    try:
        if WINDOWS:
            import winsound
            winsound.Beep(880, 200)
        elif MAC and os.path.exists("/System/Library/Sounds/Glass.aiff"):
            subprocess.run(["afplay", "/System/Library/Sounds/Glass.aiff"], check=False)
        else:
            print("\a", end="", flush=True)
    except Exception:
        pass


# ============================================================
# HUMAN-LIKE TYPING
# ============================================================


def type_like_a_human(text, allow_typos=True, wpm=None, after_sentence=None, progress_callback=None):
    if TEST_MODE and not PACE.dry:
        if INPUT_BACKEND == "firefox":
            type_literal(text)
            if progress_callback is not None:
                progress_callback(len(text))
            if after_sentence is not None and any(c in text for c in ".!?"):
                after_sentence()
            return
        for char in text:
            type_char(char)
            if progress_callback is not None:
                progress_callback(1)
            if char in ".!?" and after_sentence is not None:
                after_sentence()
        return

    wpm = wpm or random.uniform(MIN_WPM, MAX_WPM)
    base = 60.0 / (wpm * 5)
    word_speed = 1.0

    for i, char in enumerate(text):
        if (i == 0 or text[i - 1] in " \n\t") and not char.isspace():
            word_speed = random.choice((0.6, 0.8, 1.0, 1.0, 1.3))
            if random.random() < THINK_CHANCE:
                pause(random.uniform(*THINK_PAUSE))

        wrong = typo_for(char) if allow_typos and random.random() < TYPO_CHANCE else None
        if wrong:
            type_char(wrong)
            pause(base * random.uniform(0.5, 1.2))
            extra = 0
            for _ in range(random.randint(0, 2)):
                nxt = i + 1 + extra
                if nxt >= len(text) or not text[nxt].isalpha():
                    break
                type_char(text[nxt])
                pause(base * random.uniform(0.5, 1.2))
                extra += 1
            pause(random.uniform(*PAUSE_NOTICE_TYPO))
            for _ in range(extra + 1):
                send("backspace")
                pause(base * random.uniform(0.3, 0.6))

        type_char(char)
        if progress_callback is not None:
            progress_callback(1)

        if char in ".!?":
            pause(random.uniform(*PAUSE_AFTER_SENTENCE))
            pause_point("sentence")
            if after_sentence is not None:
                after_sentence()
        elif char in ",;:":
            pause(random.uniform(*PAUSE_AFTER_COMMA))
        elif char == "\n":
            pause(random.uniform(*PAUSE_AFTER_NEWLINE))
            pause_point("paragraph")
        elif char == " ":
            pause(base * random.uniform(1.0, 2.5))
        elif char == "\t":
            pause(base * random.uniform(1.0, 1.8))
        else:
            pause(base * word_speed * random.uniform(0.5, 1.5))


# ============================================================
# DOCX PARSING
# ============================================================

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def run_font_size(run):
    try:
        return round(run.font.size.pt, 1) if run.font.size else None
    except Exception:
        return None


def run_color(run):
    try:
        return str(run.font.color.rgb) if run.font.color and run.font.color.rgb else None
    except Exception:
        return None


def style_name(paragraph):
    try:
        return paragraph.style.name or ""
    except Exception:
        return ""


def paragraph_alignment(paragraph):
    try:
        return {
            WD_ALIGN_PARAGRAPH.LEFT: "left",
            WD_ALIGN_PARAGRAPH.CENTER: "center",
            WD_ALIGN_PARAGRAPH.RIGHT: "right",
            WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
        }.get(paragraph.alignment)
    except Exception:
        return None


def points(length):
    try:
        return round(length.pt, 2) if length is not None else None
    except Exception:
        return None


def outline_level(paragraph):
    p_pr = paragraph._p.pPr
    found = p_pr.find(W_NS + "outlineLvl") if p_pr is not None else None
    style = paragraph.style if found is None else None
    while found is None and style is not None:
        s_pr = style.element.find(W_NS + "pPr")
        found = s_pr.find(W_NS + "outlineLvl") if s_pr is not None else None
        style = style.base_style
    try:
        value = int(found.get(W_NS + "val")) if found is not None else None
    except (TypeError, ValueError):
        return None
    return value if value is not None and value < 9 else None


def detect_heading_level(paragraph):
    name = style_name(paragraph).strip().lower()
    if name == "title":
        return 0
    if name == "subtitle":
        return 2
    match = re.match(r"heading\s*([1-9])$", name)
    if match:
        return int(match.group(1))
    level = outline_level(paragraph)
    return level + 1 if level is not None else None


def is_reference_heading(text):
    return text.strip().rstrip(":").strip().lower() in REFERENCE_HEADINGS


def detect_list(paragraph):
    style = style_name(paragraph).lower()
    if "list bullet" in style:
        return "bullet"
    if "list number" in style:
        return "number"
    try:
        p_pr = paragraph._p.pPr
        if p_pr is None or p_pr.numPr is None or p_pr.numPr.numId is None:
            return None
        num_id = str(p_pr.numPr.numId.val)
        ilvl = str(p_pr.numPr.ilvl.val) if p_pr.numPr.ilvl is not None else "0"
        numbering = paragraph.part.numbering_part.element
        abstract_id = None
        for num in numbering.findall(W_NS + "num"):
            if num.get(W_NS + "numId") == num_id:
                abstract = num.find(W_NS + "abstractNumId")
                abstract_id = abstract.get(W_NS + "val") if abstract is not None else None
                break
        for abstract in numbering.findall(W_NS + "abstractNum"):
            if abstract.get(W_NS + "abstractNumId") != abstract_id:
                continue
            for lvl in abstract.findall(W_NS + "lvl"):
                if lvl.get(W_NS + "ilvl") == ilvl:
                    fmt = lvl.find(W_NS + "numFmt")
                    return "bullet" if fmt is not None and fmt.get(W_NS + "val") == "bullet" else "number"
        return "bullet"
    except Exception:
        return None


def run_flag(run, attr):
    value = getattr(run, attr)
    style = getattr(run, "style", None) if value is None else None
    while value is None and style is not None:
        value = getattr(style.font, attr)
        style = style.base_style
    return bool(value)


def list_level(paragraph):
    match = re.search(r"list (?:bullet|number) ([2-9])$", style_name(paragraph).lower())
    if match:
        return int(match.group(1)) - 1
    try:
        num_pr = paragraph._p.pPr.numPr
        return int(num_pr.ilvl.val) if num_pr is not None and num_pr.ilvl is not None else 0
    except Exception:
        return 0


def effective_format(paragraph, attr):
    value = getattr(paragraph.paragraph_format, attr, None)
    style = paragraph.style if value is None else None
    while value is None and style is not None:
        value = getattr(style.paragraph_format, attr, None)
        style = style.base_style
    return value


def doc_default_spacing(document):
    try:
        sp = document.styles.element.find(f"{W_NS}docDefaults/{W_NS}pPrDefault/{W_NS}pPr/{W_NS}spacing")
        if sp is None:
            return None, None, None
        line, rule, after = sp.get(W_NS + "line"), sp.get(W_NS + "lineRule"), sp.get(W_NS + "after")
        before = sp.get(W_NS + "before")
        if line is None:
            line_spacing = None
        elif rule in (None, "auto"):
            line_spacing = round(int(line) / 240, 2)
        else:
            line_spacing = Pt(int(line) / 20)
        return (line_spacing, round(int(after) / 20, 1) if after is not None else None,
                round(int(before) / 20, 1) if before is not None else None)
    except Exception:
        return None, None, None


PREFS_MARKER = Path(__file__).with_name(".dripwriter_docs_prefs_done")


def check_docs_preferences():
    if EDITOR != "google_docs" or not CHECK_DOCS_PREFERENCES or PREFS_MARKER.exists():
        return
    print("\nONE-TIME SETUP (saved to your Google account, so you only do this once):")
    print("  In any Google Doc: Tools > Preferences > General tab, UNTICK:")
    print("    - Automatically capitalize words")
    print("    - Automatically detect lists")
    print("    - Use smart quotes   (keeps quotes identical to the original)")
    print("  Then the Substitutions tab: UNTICK 'Automatic substitution'.")
    print("  Click OK.")
    answer = input("Type y when done (Enter = remind me next time): ").strip().lower()
    if answer.startswith("y"):
        try:
            PREFS_MARKER.write_text("done\n")
            print("Saved - you won't be asked again.\n")
        except OSError:
            pass


def uses_real_lists():
    return REAL_LISTS and editor_keys()["list"] is not None


def run_text(run):
    out = []
    for child in run._r:
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "t":
            out.append(child.text or "")
        elif tag == "tab":
            out.append("\t")
        elif tag == "br":
            kind = child.get(W_NS + "type")
            out.append("\f" if kind == "page" else ("" if kind == "column" else "\n"))
        elif tag == "cr":
            out.append("\n")
        elif tag == "noBreakHyphen":
            out.append("-")
    return "".join(out)


def list_id(paragraph):
    try:
        num_pr = paragraph._p.pPr.numPr if paragraph._p.pPr is not None else None
        if num_pr is not None and num_pr.numId is not None:
            return f"num{num_pr.numId.val}"
        style = paragraph.style
        while style is not None:
            s_pr = style.element.find(W_NS + "pPr")
            num = s_pr.find(f"{W_NS}numPr/{W_NS}numId") if s_pr is not None else None
            if num is not None:
                return f"num{num.get(W_NS + 'val')}"
            style = style.base_style
    except Exception:
        pass
    return "style:" + style_name(paragraph)


PASS_THROUGH = {"sdt", "sdtContent", "customXml", "smartTag", "ins", "moveTo",
                "tbl", "tr", "tc", "fldSimple", "hyperlink", "bdo", "dir"}
SKIP = {"del", "moveFrom", "txbxContent", "drawing", "pict", "object", "AlternateContent"}


def local_name(element):
    tag = element.tag
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def body_paragraphs(document):
    out = []

    def walk(element, in_table):
        for child in element.iterchildren():
            name = local_name(child)
            if name == "p":
                out.append((Paragraph(child, document._body), in_table))
            elif name in PASS_THROUGH:
                walk(child, in_table or name == "tbl")

    walk(document.element.body, False)
    return out


def run_elements(p_element):
    found = []

    def walk(element, in_link):
        for child in element.iterchildren():
            name = local_name(child)
            if name == "r":
                found.append((child, in_link))
            elif name in SKIP:
                continue
            elif name in PASS_THROUGH or name in ("sdt", "sdtContent"):
                walk(child, in_link or name == "hyperlink")

    walk(p_element, False)
    return found


def paragraph_runs(paragraph):
    runs = [(Run(r, paragraph), link) for r, link in run_elements(paragraph._p)]
    texts = [run_text(run) for run, _ in runs]
    return [{
        "text": text,
        "bold": run_flag(run, "bold"),
        "italic": run_flag(run, "italic"),
        "underline": run_flag(run, "underline"),
        "strike": bool(getattr(run.font, "strike", False)),
        "superscript": bool(getattr(run.font, "superscript", False)),
        "subscript": bool(getattr(run.font, "subscript", False)),
        "font_name": run.font.name,
        "font_size": run_font_size(run),
        "color": run_color(run),
        "link": link,
    } for (run, link), text in zip(runs, texts) if text]


def parse_docx(path):
    if not DOCX_AVAILABLE:
        raise RuntimeError("python-docx is not installed")
    document = Document(str(path))
    items, counters = [], {}
    real_lists = uses_real_lists()
    default_line, default_after, default_before = doc_default_spacing(document)
    for index, (paragraph, in_table) in enumerate(body_paragraphs(document), start=1):
        runs = paragraph_runs(paragraph)
        list_kind = detect_list(paragraph)
        level = list_level(paragraph) if list_kind else 0
        lid = list_id(paragraph) if list_kind else None
        if list_kind:
            levels = counters.setdefault(lid, {})
            for deeper in [l for l in levels if l > level]:
                del levels[deeper]
            if list_kind == "number":
                levels[level] = levels.get(level, 0) + 1
        counter = counters.get(lid, {}).get(level, 0) if list_kind else 0
        if real_lists or not list_kind:
            prefix = ""
        else:
            prefix = (f"{counter}." if list_kind == "number" else "•") + LIST_MARKER_GAP
        pf = paragraph.paragraph_format
        items.append({
            "index": index,
            "text": "".join(r["text"] for r in runs),
            "runs": runs,
            "style": style_name(paragraph),
            "heading_level": detect_heading_level(paragraph),
            "alignment": paragraph_alignment(paragraph),
            "list_kind": list_kind,
            "list_number": counter if list_kind == "number" else None,
            "list_level": level,
            "list_id": lid,
            "in_table": in_table,
            "break_before": bool(effective_format(paragraph, "page_break_before")),
            "line_spacing": effective_format(paragraph, "line_spacing") or default_line,
            "eff_space_after": (points(effective_format(paragraph, "space_after"))
                                if effective_format(paragraph, "space_after") is not None
                                else default_after),
            "eff_space_before": (points(effective_format(paragraph, "space_before"))
                                 if effective_format(paragraph, "space_before") is not None
                                 else default_before),
            "eff_left_indent": points(effective_format(paragraph, "left_indent")),
            "eff_first_line": points(effective_format(paragraph, "first_line_indent")),
            "left_indent": points(pf.left_indent),
            "right_indent": points(pf.right_indent),
            "first_line_indent": points(pf.first_line_indent),
            "space_before": points(pf.space_before),
            "space_after": points(pf.space_after),
            "prefix": prefix,
        })
    mark_list_blocks(items)
    sections = [{"left": points(s.left_margin), "right": points(s.right_margin),
                 "top": points(s.top_margin), "bottom": points(s.bottom_margin)}
                for s in document.sections]
    return document, split_document(items), sections


def mark_list_blocks(items):
    block_kind, block_id, seen_ids = None, None, set()
    i = 0
    while i < len(items):
        item = items[i]
        kind = item["list_kind"]
        if kind and item["text"].strip():
            level = item.get("list_level") or 0
            if block_kind is None or level == 0 and (kind != block_kind or item["list_id"] != block_id):
                if block_kind is None and item["list_id"] in seen_ids and kind == "number":
                    item["continues_numbering"] = True
                elif block_kind == kind == "number":
                    item["list_restart"] = True
                block_kind, block_id = kind, item["list_id"]
            item["eff_list_kind"] = block_kind
            seen_ids.add(item["list_id"])
            j = i + 1
            while j < len(items) and not items[j]["text"].strip() and items[j]["heading_level"] is None:
                j += 1
            nxt = items[j] if j < len(items) else None
            if (KEEP_LIST_TOGETHER and j > i + 1 and nxt is not None and nxt["list_kind"]
                    and nxt["text"].strip() and nxt["heading_level"] is None
                    and (nxt["list_id"] == item["list_id"] or (nxt.get("list_level") or 0) > 0)):
                item["blank_lines_after"] = j - i - 1
                for k in range(i + 1, j):
                    items[k]["soft_blank"] = True
                i = j
                continue
        elif item["text"].strip() or item["heading_level"] is not None:
            block_kind = block_id = None
        elif not kind:
            block_kind = block_id = None
        i += 1


def split_document(items):
    for pos, item in enumerate(items):
        if is_reference_heading(item["text"]):
            refs, end = [], pos + 1
            while end < len(items) and items[end]["heading_level"] is None:
                if items[end]["text"].strip():
                    refs.append(items[end])
                end += 1
            return {"body": items[:pos], "ref_heading": item,
                    "refs": refs, "trailing": items[end:]}
    return {"body": items, "ref_heading": None, "refs": [], "trailing": []}


# ============================================================
# CITATION MATCHING
# ============================================================

YEAR_RE = re.compile(r"\b(?:1[89]\d{2}|20\d{2})[a-z]?\b|\bn\.d\.")


def reference_key(ref):
    text = ref["text"].strip()
    m = re.match(r"^\[(\d+)\]|^(\d+)[.)]\s", text)
    if m:
        return ("number", int(m.group(1) or m.group(2)))
    if ref.get("list_number"):
        return ("number", ref["list_number"])
    author = re.split(r",|\(|\.\s", text, maxsplit=1)[0].strip()
    year = YEAR_RE.search(text)
    return ("author", re.sub(r"\s+", " ", author), year.group(0) if year else None)


SUPER_DIGITS = str.maketrans({
    "⁰": "0", "¹": "1", "²": "2", "³": "3", "⁴": "4",
    "⁵": "5", "⁶": "6", "⁷": "7", "⁸": "8", "⁹": "9",
})
SUPER_DIGIT_RE = re.compile(r"[⁰¹²³⁴⁵⁶⁷⁸⁹]+")

PLAIN_TEXT_TRANSLATION = str.maketrans({
    "⁰": "0", "¹": "1", "²": "2", "³": "3", "⁴": "4",
    "⁵": "5", "⁶": "6", "⁷": "7", "⁸": "8", "⁹": "9",
    "₀": "0", "₁": "1", "₂": "2", "₃": "3", "₄": "4",
    "₅": "5", "₆": "6", "₇": "7", "₈": "8", "₉": "9",
    "⁺": "+", "⁻": "-", "⁽": "(", "⁾": ")",
})

SUPERSCRIPT_CHARS = set("⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁽⁾")
SUBSCRIPT_CHARS = set("₀₁₂₃₄₅₆₇₈₉")


def script_pieces(text, run):
    pieces = []
    for char in text:
        sup = bool(run.get("superscript")) or char in SUPERSCRIPT_CHARS
        sub = not sup and (bool(run.get("subscript")) or char in SUBSCRIPT_CHARS)
        if not SUPERSCRIPTS:
            sup = sub = False
        char = plain_text(char)
        if pieces and pieces[-1][1:] == (sup, sub):
            pieces[-1] = (pieces[-1][0] + char, sup, sub)
        else:
            pieces.append((char, sup, sub))
    return pieces


def plain_text(text):
    return text.translate(PLAIN_TEXT_TRANSLATION)


def superscript_cited_numbers(item):
    found = set()
    for run in item.get("runs", []):
        text = run.get("text", "")
        is_sup = bool(run.get("superscript"))
        for match in SUPER_DIGIT_RE.finditer(text):
            if is_sup or any(ch in "⁰¹²³⁴⁵⁶⁷⁸⁹" for ch in match.group(0)):
                try:
                    found.add(int(match.group(0).translate(SUPER_DIGITS)))
                except ValueError:
                    pass
    return found


def citation_numbers_from_item(item):
    return cited_numbers(item.get("text", "")) | superscript_cited_numbers(item)


def cited_numbers(text):
    found = set()
    for group in re.findall(r"\[([\d,\s\-–]+)\]", text):
        for part in group.split(","):
            span = [s.strip() for s in re.split(r"[-–]", part.strip())]
            if len(span) == 2 and span[0].isdigit() and span[1].isdigit():
                a, b = int(span[0]), int(span[1])
                if a <= b and b - a < 100:
                    found.update(range(a, b + 1))
            elif len(span) == 1 and span[0].isdigit():
                found.add(int(span[0]))
    return found


def surname_of(key):
    return key[1].split(" ")[0] if key[0] == "author" and key[1] else ""


def cites(item, key):
    text = item.get("text", "") if isinstance(item, dict) else str(item)
    if key[0] == "number":
        if isinstance(item, dict):
            return key[1] in citation_numbers_from_item(item)
        return key[1] in cited_numbers(text)
    name, year = surname_of(key), key[2]
    if not name or not re.search(rf"\b{re.escape(name)}\b", text):
        return False
    return year is None or year in text


def cites_loosely(item, key):
    text = item.get("text", "") if isinstance(item, dict) else str(item)
    name = surname_of(key)
    if not name:
        return False
    n = re.escape(name)
    return bool(re.search(rf"\([^()]*\b{n}\b[^()]*\)|\b{n}\s*\(", text))


def plan_citations(parts):
    body = [(pos, item) for pos, item in enumerate(parts["body"]) if item["heading_level"] is None]
    first_cited = {}
    for ref_pos, ref in enumerate(parts["refs"]):
        key = reference_key(ref)
        for matcher in (cites, cites_loosely):
            hit = next((pos for pos, item in body if matcher(item, key)), None)
            if hit is not None:
                first_cited[ref_pos] = hit
                break
    per_paragraph = {}
    for ref_pos, body_pos in sorted(first_cited.items()):
        per_paragraph.setdefault(body_pos, []).append(ref_pos)
    uncited = [i for i in range(len(parts["refs"])) if i not in first_cited]
    return per_paragraph, uncited


# ============================================================
# OPTIONAL: MICROSOFT WORD EXACT FORMATTING
# ============================================================


def word_color(value):
    try:
        r, g, b = int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16)
        return r + (g << 8) + (b << 16)
    except (TypeError, ValueError):
        return None


def get_word_app():
    if EDITOR != "word" or not (WIN32_WORD_AVAILABLE and WINDOWS_WORD_EXACT_FORMAT):
        return None
    try:
        app = win32com.client.GetActiveObject("Word.Application")
        app.ActiveDocument
        return app
    except Exception:
        return None


def word_polish_just_typed(app, item):
    try:
        doc = app.ActiveDocument
        end = app.Selection.Start
        start = max(end - len(item["text"]), 0)
        rng = doc.Range(start, end).Paragraphs(1).Range
        if item["style"]:
            try:
                rng.Style = item["style"]
            except Exception:
                pass
        fmt = rng.ParagraphFormat
        align = {"left": 0, "center": 1, "right": 2, "justify": 3}.get(item["alignment"])
        if align is not None:
            fmt.Alignment = align
        for attr, key in (("LeftIndent", "left_indent"), ("RightIndent", "right_indent"),
                          ("FirstLineIndent", "first_line_indent"),
                          ("SpaceBefore", "space_before"), ("SpaceAfter", "space_after")):
            if item[key] is not None:
                setattr(fmt, attr, item[key])
        offset = start
        for run in item["runs"]:
            font = doc.Range(offset, offset + len(run["text"])).Font
            if run["font_name"]:
                font.Name = run["font_name"]
            if run["font_size"] is not None:
                font.Size = run["font_size"]
            font.Bold = -1 if run["bold"] else 0
            font.Italic = -1 if run["italic"] else 0
            font.Underline = 1 if run["underline"] else 0
            font.StrikeThrough = -1 if run["strike"] else 0
            font.Superscript = 0
            font.Subscript = 0
            color = word_color(run["color"])
            if color is not None:
                font.Color = color
            offset += len(run["text"])
    except Exception as exc:
        print(f"  [Word exact formatting skipped for this paragraph: {exc}]")


def word_page_setup(app, sections):
    if not sections:
        return
    try:
        setup = app.ActiveDocument.Sections(1).PageSetup
        for attr, key in (("LeftMargin", "left"), ("RightMargin", "right"),
                          ("TopMargin", "top"), ("BottomMargin", "bottom")):
            if sections[0][key] is not None:
                setattr(setup, attr, sections[0][key])
    except Exception as exc:
        print(f"[Page setup skipped: {exc}]")


# ============================================================
# LIVE DOCX WRITER
# ============================================================


NO_FORMAT = {"bold": False, "italic": False, "underline": False,
             "superscript": False, "subscript": False}


class LiveWriter:
    def __init__(self, parts, keys=None, word_app=None, wpm=None, references_mode=None):
        self.parts = parts
        self.keys = keys or editor_keys()
        self.word_app = word_app
        self.wpm = wpm or random.uniform(MIN_WPM, MAX_WPM)
        self.references_mode = references_mode or DEFAULT_REFERENCES_MODE
        self.page_breaks = PAGE_BREAKS and self.references_mode == "last"
        self.refs_started = False
        self.typed_refs = []
        self.fmt = dict(NO_FORMAT)
        self.size = None
        self.doc = []
        self.cur = -1

    @staticmethod
    def layout(item):
        has_text = bool(item["text"].strip())
        kind = (item.get("eff_list_kind") or item["list_kind"]) if (uses_real_lists() and has_text) else None
        level = min(item.get("list_level") or 0, 8) if kind else 0
        indent = tab = 0
        if COPY_INDENTS and not item["list_kind"] and item["heading_level"] is None:
            first = item.get("eff_first_line") or 0
            left = (item.get("eff_left_indent") or 0) if first >= 0 else 0
            indent = max(0, min(int(round(left / 36.0)), 8))
            tab = has_text and first >= 18
        return kind, level, indent, tab

    def here(self):
        return self.layout(self.doc[self.cur]) if self.cur >= 0 else (None, 0, 0, False)

    def go_end(self):
        shortcut(self.keys["doc_end"])
        self.cur = len(self.doc) - 1

    def apply_layout(self, item, inherited):
        kind, level, indent, _ = self.layout(item)
        ikind, ilevel, iindent, _ = inherited
        restart = bool(item.get("list_restart")) and kind == ikind
        if ikind and (ikind != kind or restart):
            for _ in range(ilevel):
                send("shift+tab")
            shortcut(self.keys["list"][ikind])
            ilevel = 0
        if kind and (kind != ikind or restart):
            shortcut(self.keys["list"][kind])
            ilevel = 0
        if kind:
            for _ in range(level - ilevel):
                send("tab")
            for _ in range(ilevel - level):
                send("shift+tab")
        if COPY_INDENTS:
            for _ in range(iindent):
                shortcut(self.keys["indent_less"])
            for _ in range(indent):
                shortcut(self.keys["indent_more"])

    def set_toggle(self, name, on):
        if self.fmt[name] != on:
            key = {"bold": f"{MOD}+b", "italic": f"{MOD}+i", "underline": f"{MOD}+u"}.get(name) \
                or self.keys.get(name)
            if not key:
                return
            if on and name in ("superscript", "subscript"):
                other = "subscript" if name == "superscript" else "superscript"
                self.set_toggle(other, False)
            shortcut(key)
            self.fmt[name] = on

    def all_toggles_off(self):
        for name in NO_FORMAT:
            self.set_toggle(name, False)

    def clear_formatting(self):
        shortcut(self.keys["clear_char"])
        self.fmt = dict(NO_FORMAT)
        self.size = None

    def set_size(self, size):
        box = self.keys["font_size_box"]
        if DOCX_TYPE_FONT_SIZES and box and size and size != self.size:
            shortcut(box)
            type_literal(f"{size:g}")
            shortcut("enter")
            self.size = size

    def start_paragraph(self, item, inherited=(None, 0, 0, False), after_heading=False):
        self.clear_formatting()
        level = item["heading_level"]
        if level is not None:
            shortcut(self.keys["heading"](level or 1))
        elif after_heading:
            shortcut(self.keys["normal"])
        shortcut(self.keys["align"][item["alignment"] or "left"])
        self.apply_layout(item, inherited)

    def type_item(self, item):
        full_text = item["prefix"] + "".join(run["text"] for run in item["runs"])
        typed_offset = 0
        set_progress(text=full_text, position=0, context="typing paragraph")

        def progress(delta):
            nonlocal typed_offset
            typed_offset += delta
            set_progress(position=min(typed_offset, len(full_text)))

        reanchor = None
        if self.references_mode == "last" and REANCHOR_AFTER_EVERY_SENTENCE:
            reanchor = lambda: shortcut(self.keys["doc_end"])

        if self.layout(item)[3]:
            send("tab")
        if item["prefix"]:
            type_literal(item["prefix"])
            progress(len(item["prefix"]))
        for run in item["runs"]:
            self.set_toggle("bold", run["bold"])
            self.set_toggle("italic", run["italic"])
            self.set_toggle("underline", run["underline"])
            self.set_size(run["font_size"])
            for n, piece in enumerate(run["text"].split("\n")):
                if n:
                    send("shift+enter")
                if not self.page_breaks:
                    piece = piece.replace("\f", "")
                for text, sup, sub in script_pieces(piece, run):
                    self.set_toggle("superscript", sup)
                    self.set_toggle("subscript", sub)
                    type_like_a_human(
                        text,
                        allow_typos=DOCX_ALLOW_TYPOS and not (sup or sub),
                        wpm=self.wpm * random.uniform(0.9, 1.1),
                        after_sentence=reanchor,
                        progress_callback=progress,
                    )
        self.all_toggles_off()
        for _ in range(item.get("blank_lines_after", 0)):
            send("shift+enter")
        if self.word_app is not None and not PACE.dry:
            word_polish_just_typed(self.word_app, item)

    def new_paragraph(self, item):
        inherited = self.here()
        after_heading = self.cur >= 0 and self.doc[self.cur]["heading_level"] is not None
        if item.get("break_before") and self.page_breaks and self.cur >= 0:
            send(f"{MOD}+enter")
        else:
            send("enter")
        self.doc.insert(self.cur + 1, item)
        self.cur += 1
        self.start_paragraph(item, inherited, after_heading)
        self.type_item(item)

    def add_references(self, ref_positions, come_back=True):
        if not ref_positions:
            return
        refs = self.parts["refs"]
        say(f"  -> adding {len(ref_positions)} reference(s) at the bottom")
        pause(random.uniform(*PAUSE_BEFORE_REFERENCES))

        for ref_pos in sorted(ref_positions):
            if not self.refs_started:
                self.new_paragraph(self.parts["ref_heading"])
                self.refs_started = True
                self.new_paragraph(refs[ref_pos])
                self.typed_refs.append(ref_pos)
                continue
            self.go_end()
            insert_at = sum(1 for p in self.typed_refs if p < ref_pos)
            after = len(self.typed_refs) - insert_at
            if after == 0:
                self.new_paragraph(refs[ref_pos])
            else:
                for _ in range(after):
                    shortcut(self.keys["para_up"])
                target = len(self.doc) - after
                self.clear_formatting()
                self.type_item(refs[ref_pos])
                send("enter")
                self.doc.insert(target, refs[ref_pos])
                self.cur = target + 1
            self.typed_refs.insert(insert_at, ref_pos)

        if come_back:
            shortcut(self.keys["doc_end"])
            for _ in range(len(self.typed_refs) + 1):
                shortcut(self.keys["para_up"])
            shortcut("left")
            heading = next(i for i, x in enumerate(self.doc) if x is self.parts["ref_heading"])
            self.cur = heading - 1
            pause(random.uniform(*PAUSE_AFTER_REFERENCES))
        pause_point("references")

    def run(self):
        body = self.parts["body"]
        has_refs = bool(self.parts["refs"]) and self.parts["ref_heading"] is not None
        per_paragraph, uncited = plan_citations(self.parts) if has_refs else ({}, [])
        if self.references_mode == "last":
            per_paragraph, uncited = {}, list(range(len(self.parts["refs"])))

        for pos, item in enumerate(body):
            if item.get("soft_blank") and pos > 0:
                continue
            set_progress(paragraph=pos + 1, total_paragraphs=len(body), text="", position=0,
                         context="starting paragraph")
            if DOCX_SHOW_PROGRESS:
                preview = item["text"].strip().replace("\n", " ")
                preview = preview[:60] + "..." if len(preview) > 63 else preview
                say(f"Paragraph {pos + 1}/{len(body)} {PACE.status()}: {preview or '(blank line)'}")
            if pos == 0:
                self.doc, self.cur = [item], 0
                self.start_paragraph(item)
                self.type_item(item)
            else:
                pause(random.uniform(*PAUSE_BETWEEN_PARAGRAPHS))
                if item.get("heading_level") is not None:
                    pause_point("section")
                self.new_paragraph(item)
            if has_refs and pos in per_paragraph:
                self.add_references(per_paragraph[pos], come_back=True)
            if pos < len(body) - 1:
                pause_point("paragraph")

        if has_refs and uncited:
            if self.references_mode == "as_you_go":
                say(f"{len(uncited)} reference(s) weren't matched to a paragraph; adding them now.")
            if self.refs_started:
                self.go_end()
            self.add_references(uncited, come_back=False)
        elif not has_refs and self.parts["ref_heading"] is not None:
            self.new_paragraph(self.parts["ref_heading"])

        if self.parts["trailing"]:
            self.go_end()
            for item in self.parts["trailing"]:
                if not item.get("soft_blank"):
                    self.new_paragraph(item)
        typed = [i for i in self.doc if i["text"].strip()]
        say(f"\nTyped {len(typed)} paragraphs, {sum(len(i['text'].split()) for i in typed)} words.")


# ============================================================
# RUNNING A JOB
# ============================================================


def ask_target_time():
    default = TARGET_TIME.strip()
    if not ASK_FOR_TARGET_TIME:
        return parse_duration(default)
    label = default or "natural speed"
    while True:
        answer = input(f"How long should it take? e.g. 5h, 90m, 1h 30m "
                       f"(Enter = {label}): ").strip() or default
        if not answer:
            return None
        seconds = parse_duration(answer)
        if seconds:
            return seconds
        print("  Didn't understand that. Try something like 5h, 45m or 2h 15m.")


def wait_for_start():
    keyboard.wait(START_KEY)
    while keyboard.is_pressed(START_KEY):
        sleep(0.02)
    sleep(0.3)


def finish_run():
    PACE.stop()
    FOCUS.disarm()
    if INPUT_BACKEND != "firefox":
        release_modifiers()
    keep_awake(False)


def run_job(job, target_seconds, test_mode=None):
    instant = TEST_MODE if test_mode is None else test_mode

    if instant:
        print("TEST MODE: no target-time pacing, delays or simulated typos.")
        if INPUT_BACKEND == "firefox":
            print("Firefox backend: key events go straight to the WebDriver tab.")
            print("You can use your other monitors freely while it types.")
            print(f"Press {START_KEY.upper()} anywhere to start (e.g. in the blank Doc).")
        else:
            print(f"Click into your document, then press {START_KEY.upper()} to start.")
        print(f"Hold {STOP_KEY.upper()} to stop.")
        wait_for_start()
        keep_awake(True)
        try:
            FOCUS.arm()
            if INPUT_BACKEND == "firefox" and FIREFOX_DRIVER is not None:
                firefox_inject_focus_lock(FIREFOX_DRIVER)
            KEYS_SENT[0] = 0
            job()
            print(f"Sent {KEYS_SENT[0]} keys; letting the editor finish typing...")
            wait(TEST_MODE_FINISH_WAIT)
            beep()
            print("\n[TEST COMPLETE]")
        except Stopped:
            print("\n[Stopped by user]")
        finally:
            finish_run()
        return

    seed = random.random()
    random.seed(seed)
    PACE.start_dry_run()
    try:
        job()
    finally:
        natural, kinds = PACE.end_dry_run()

    print(f"At natural speed this takes about {fmt_duration(natural)}.")
    if target_seconds:
        fastest = natural / MAX_SPEEDUP
        if target_seconds < fastest:
            print(f"WARNING: {fmt_duration(target_seconds)} is too fast; the quickest "
                  f"it can go is about {fmt_duration(fastest)} (see MAX_SPEEDUP).")
        elif target_seconds >= natural:
            print(f"Target {fmt_duration(target_seconds)}: the extra time becomes "
                  f"longer pauses and breaks, especially at paragraphs and section ends.")
        else:
            print(f"Target {fmt_duration(target_seconds)}: typing faster than natural.")

    if INPUT_BACKEND == "firefox":
        print(f"\nPress {START_KEY.upper()} anywhere to start (e.g. in the blank Doc). "
              f"Hold {STOP_KEY.upper()} to stop.")
        print("You can use your other monitors freely while it types.")
    else:
        print(f"\nClick into your document, then press {START_KEY.upper()} to start. "
              f"Hold {STOP_KEY.upper()} to stop.")
    wait_for_start()
    keep_awake(True)
    try:
        FOCUS.arm()
        if INPUT_BACKEND == "firefox" and FIREFOX_DRIVER is not None:
            firefox_inject_focus_lock(FIREFOX_DRIVER)
        wait(DOCX_START_DELAY)
        if target_seconds:
            PACE.begin(target_seconds, natural, kinds)
        random.seed(seed)
        start = monotonic()
        job()
        beep()
        print(f"\n[Done in {fmt_duration(monotonic() - start)}]")
    except Stopped:
        print("\n[Stopped by user]")
    finally:
        finish_run()


# ============================================================
# DOCX WORKFLOW
# ============================================================


def print_docx_summary(document, parts, path):
    body, refs = parts["body"], parts["refs"]
    print("\n========== DOCX ANALYSIS ==========")
    print(f"File: {path.name}   Editor: {EDITOR}")
    print(f"Body paragraphs: {len(body)}   "
          f"Headings: {sum(1 for i in body if i['heading_level'] is not None)}   "
          f"References: {len(refs)}")
    if refs and parts["ref_heading"] is not None:
        per_paragraph, uncited = plan_citations(parts)
        print("\nReference plan:")
        print("  (Source superscripts are flattened to plain-text digits when typed.)")
        for body_pos in sorted(per_paragraph):
            for r in per_paragraph[body_pos]:
                print(f"  after paragraph {body_pos + 1}: {refs[r]['text'][:60]}")
        for r in uncited:
            print(f"  at the end (no citation found): {refs[r]['text'][:60]}")
    table_paras = [i for i in body + refs + parts["trailing"] if i.get("in_table") and i["text"].strip()]
    if table_paras:
        print(f"\nTABLES: their text ({len(table_paras)} cell paragraph(s)) is typed in reading")
        print("  order as ordinary paragraphs; the table itself isn't recreated.")
    words = sum(len(i["text"].split()) for i in body + refs + parts["trailing"]
                + ([parts["ref_heading"]] if parts["ref_heading"] else []))
    print(f"\nTotal to type: {words} words. Check the finished doc has about the same")
    print("  (Tools > Word count), or use mode 5 to compare.")
    all_items = body + refs + parts["trailing"]
    if EDITOR == "google_docs":
        if not PREFS_MARKER.exists():
            print("\nGoogle Docs: Tools > Preferences, turn these OFF so Docs doesn't change")
            print("  what is typed: Automatically capitalize words, Automatically detect")
            print("  lists, and Substitutions (e.g. '--', '(c)', '1/2' are replaced).")
        text = " ".join(i["text"] for i in all_items + ([parts["ref_heading"]] if parts["ref_heading"] else []))
        risky = sorted({m for m in re.findall(r"--|\(c\)|\(r\)|->|<-|\b[13]/[24]\b|\b\d+(?:st|nd|rd|th)\b", text)})
        if risky:
            print(f"\n  This document contains text that Docs' Substitutions would change: "
                  f"{', '.join(risky[:8])}")
    if EDITOR == "google_docs":
        for i in all_items:
            if i.get("continues_numbering"):
                print(f"\nLIST: '{i['text'][:40]}' continues an earlier list's numbering")
                print(f"  (the original shows {i['list_number']}.). Docs will restart at 1: when it")
                print("  finishes, right-click that item > Continue previous numbering.")
    if any(i.get("break_before") or "\f" in i["text"] for i in all_items):
        print("\nPAGE BREAKS: typed in reference mode 'last'; skipped in 'as_you_go'.")
    if any((r.get("eff_first_line") or 0) < 0 for r in refs):
        print("\nThe references use a hanging indent (no shortcut exists). When it finishes:")
        print("  select the references > Format > Align & indent > Indentation options >")
        print("  Special indent: Hanging.")
    if EDITOR == "google_docs":
        print("\nNot copied (set these yourself): spacing, margins, font. Title and")
        print("  Subtitle become Heading 1 and Heading 2 (Docs has no shortcut for them).")
    print("===================================\n")


def choose_docx(prompt="Choose the DOCX to recreate"):
    if not DOCX_AVAILABLE:
        print("\nDOCX support needs python-docx:  py -m pip install python-docx")
        return None
    try:
        if WINDOWS:
            ps = ("Add-Type -AssemblyName System.Windows.Forms; "
                  "$d=New-Object System.Windows.Forms.OpenFileDialog; "
                  '$d.Filter="Word Documents (*.docx)|*.docx"; '
                  f"$d.Title='{prompt.replace(chr(39), chr(39) * 2)}'; "
                  "if($d.ShowDialog() -eq 'OK'){Write-Output $d.FileName}")
            result = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                                    capture_output=True, text=True, check=False)
        elif MAC:
            result = subprocess.run(
                ["osascript", "-e",
                 'set f to choose file of type {"org.openxmlformats.wordprocessingml.document"} '
                 f'with prompt "{prompt}"', "-e", "POSIX path of f"],
                capture_output=True, text=True, check=False)
        else:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            path = filedialog.askopenfilename(title=prompt,
                                              filetypes=[("Word Documents", "*.docx")])
            root.destroy()
            return Path(path) if path else None
        path = result.stdout.strip()
        return Path(path) if path else None
    except Exception as exc:
        print(f"Could not open file picker: {exc}")
        return None


def docx_workflow(analyze_only=False, references_mode=None, test_mode=None):
    path = choose_docx()
    if not path:
        print("No DOCX selected.\n")
        return
    print(f"\nReading: {path.name}")
    try:
        document, parts, sections = parse_docx(path)
    except Exception as exc:
        print(f"Could not read the DOCX: {exc}")
        return
    print_docx_summary(document, parts, path)
    if analyze_only:
        return

    instant = TEST_MODE if test_mode is None else test_mode
    check_docs_preferences()
    if INPUT_BACKEND == "firefox":
        print("Make sure your Google Doc is open in the attached Firefox tab.")
        if not firefox_alive():
            print("WARNING: Firefox WebDriver is not responding. Restart the script.")
            return
    else:
        print("Open a BLANK document to type into.")
    target = None if instant else ask_target_time()
    wpm = random.uniform(MIN_WPM, MAX_WPM)
    selected_reference_mode = references_mode or DEFAULT_REFERENCES_MODE
    print(f"Reference mode: {selected_reference_mode}")
    if instant:
        print("TEST MODE is ON: the recreation runs without pacing or simulated typos.")
    state = {"word_app": None}

    FOCUS.first_line = next((i["text"] for i in parts["body"] if i["text"].strip()), "")

    def job():
        if not PACE.dry:
            state["word_app"] = get_word_app()
            if state["word_app"] is not None:
                print("Microsoft Word detected: exact formatting applied as it types.")
                word_page_setup(state["word_app"], sections)
        LiveWriter(parts, word_app=state["word_app"], wpm=wpm,
                   references_mode=selected_reference_mode).run()

    try:
        run_job(job, target, test_mode=instant)
    except Exception as exc:
        print(f"\n[DOCX workflow error] {exc}")


# ============================================================
# COMPARE MODE
# ============================================================

SMART_QUOTES = {"\u2018": "'", "\u2019": "'", "\u201c": '"', "\u201d": '"'}


def heading_as_typed(level):
    if level is None:
        return None
    return int(editor_keys()["heading"](level or 1).rsplit("+", 1)[1])


def comparable(item):
    chars = [(c, None) for c in item["prefix"]]
    for run in item["runs"]:
        fmt = (run["bold"], run["italic"], run["underline"], run["link"])
        chars.extend((c, fmt) for c in plain_text(run["text"]))
    out = []
    for c, fmt in chars:
        if COMPARE_IGNORE_SMART_QUOTES:
            c = SMART_QUOTES.get(c, c)
        if c.isspace():
            if not out or out[-1][0] == " ":
                continue
            c = " "
        out.append((c, fmt))
    while out and out[-1][0] == " ":
        out.pop()
    return {"text": "".join(c for c, _ in out), "fmt": [f for _, f in out],
            "heading": heading_as_typed(item["heading_level"]),
            "align": item["alignment"] or "left", "number": item["index"],
            "list": item["list_kind"]}


def compare_paragraphs(path):
    _, parts, _ = parse_docx(path)
    items = parts["body"] + [parts["ref_heading"]] * (parts["ref_heading"] is not None)
    items += parts["refs"] + parts["trailing"]
    paras = [comparable(i) for i in items]
    if COMPARE_IGNORE_BLANK_LINES:
        paras = [p for p in paras if p["text"]]
    refs = [comparable(r)["text"] for r in parts["refs"]]
    return paras, refs


def short(text, width=50):
    return repr(text if len(text) <= width else text[:width - 3] + "...")


def first_difference(a, b):
    n = next((i for i, (x, y) in enumerate(zip(a, b)) if x != y), min(len(a), len(b)))
    start = max(n - 15, 0)
    return f"original {short(a[start:n + 25])} vs copy {short(b[start:n + 25])}"


def format_differences(a, b):
    found = []
    for attr, name in ((0, "bold"), (1, "italic"), (2, "underlined")):
        i = 0
        while i < len(a["text"]):
            def differs(k):
                fa, fb = a["fmt"][k] or (False,) * 4, b["fmt"][k] or (False,) * 4
                if a["text"][k] == " " or (attr == 2 and (fa[3] or fb[3])):
                    return False
                return fa[attr] != fb[attr]
            if not differs(i):
                i += 1
                continue
            j = i
            while j < len(a["text"]) and (differs(j) or a["text"][j] == " "):
                j += 1
            want = (a["fmt"][i] or (False,) * 4)[attr]
            found.append(f"{short(a['text'][i:j].strip(), 40)} should "
                         f"{'' if want else 'not '}be {name}")
            i = j
    return found


def compare_documents(original_path, copy_path):
    orig, orig_refs = compare_paragraphs(original_path)
    copy, copy_refs = compare_paragraphs(copy_path)
    diffs = []

    def where(p, q=None):
        return f"original paragraph {p['number']}" + (f" / copy paragraph {q['number']}" if q else "")

    def check_pair(a, b):
        if a["text"] != b["text"]:
            diffs.append(f"{where(a, b)}: text differs - {first_difference(a['text'], b['text'])}")
            return
        if COMPARE_HEADINGS_AND_ALIGNMENT:
            if a["heading"] != b["heading"]:
                label = lambda h: "normal text" if h is None else f"Heading {h}"
                diffs.append(f"{where(a, b)}: should be {label(a['heading'])}, "
                             f"copy has {label(b['heading'])} - {short(a['text'], 40)}")
            if a["align"] != b["align"]:
                diffs.append(f"{where(a, b)}: should be {a['align']}-aligned, "
                             f"copy is {b['align']} - {short(a['text'], 40)}")
        if a["list"] != b["list"]:
            label = lambda k: f"a {k} list item" if k else "not a list item"
            diffs.append(f"{where(a, b)}: should be {label(a['list'])}, copy is "
                         f"{label(b['list'])} - {short(a['text'], 40)}")
        if COMPARE_FORMATTING:
            diffs.extend(f"{where(a, b)}: {d}" for d in format_differences(a, b))

    matcher = difflib.SequenceMatcher(None, [p["text"] for p in orig],
                                      [p["text"] for p in copy], autojunk=False)
    for op, i1, i2, j1, j2 in matcher.get_opcodes():
        pairs = min(i2 - i1, j2 - j1) if op in ("equal", "replace") else 0
        for k in range(pairs):
            check_pair(orig[i1 + k], copy[j1 + k])
        for p in orig[i1 + pairs:i2]:
            diffs.append(f"{where(p)}: missing from the copy - {short(p['text'])}")
        for q in copy[j1 + pairs:j2]:
            diffs.append(f"copy paragraph {q['number']}: extra, not in the original - {short(q['text'])}")

    if orig_refs != copy_refs and sorted(orig_refs) == sorted(copy_refs):
        k = next(i for i, (x, y) in enumerate(zip(orig_refs, copy_refs)) if x != y)
        diffs.insert(0, f"references are all there but out of order: entry {k + 1} "
                        f"should be {short(orig_refs[k], 40)}, copy has {short(copy_refs[k], 40)}")
    return diffs


def compare_workflow():
    if not DOCX_AVAILABLE:
        print("\nComparing needs python-docx:  py -m pip install python-docx")
        return
    print("\nPick the ORIGINAL, then the COPY.")
    if EDITOR == "google_docs":
        print("Get the copy from Google Docs with File > Download > Microsoft Word (.docx).")
    original = choose_docx("Choose the ORIGINAL DOCX")
    copy = original and choose_docx("Choose the COPY to check")
    if not original or not copy:
        print("Two files are needed.\n")
        return
    try:
        diffs = compare_documents(original, copy)
    except Exception as exc:
        print(f"Could not compare: {exc}")
        return
    print("\n========== COMPARE ==========")
    print(f"Original: {original.name}\nCopy:     {copy.name}")
    if not diffs:
        print("\nPASS: same text, paragraph order and reference order"
              + (", headings, alignment" if COMPARE_HEADINGS_AND_ALIGNMENT else "")
              + (", bold/italic/underline" if COMPARE_FORMATTING else "") + ".")
    else:
        print(f"\nFAIL: {len(diffs)} difference(s).")
        for d in diffs[:COMPARE_MAX_DIFFERENCES]:
            print(f"  - {d}")
        if len(diffs) > COMPARE_MAX_DIFFERENCES:
            print(f"  ...and {len(diffs) - COMPARE_MAX_DIFFERENCES} more.")
    print("Not checked: tables, font sizes, fonts, colours, spacing.")
    print("=============================\n")


# ============================================================
# PASTE MODE
# ============================================================


def read_text():
    print("Paste your text below, then press Enter twice when done:\n")
    lines = []
    while True:
        line = input()
        if line.strip().lower() == "exit" and not lines:
            return None
        if line == "" and lines and lines[-1] == "":
            break
        lines.append(line)
    if lines and lines[-1] == "":
        lines.pop()
    return "\n".join(lines)


def paste_workflow():
    text = read_text()
    if text is None:
        print("Goodbye!")
        return False
    if not text.strip():
        print("No text entered. Try again.\n")
        return True
    target = ask_target_time()
    wpm = random.uniform(MIN_WPM, MAX_WPM)
    try:
        run_job(lambda: type_like_a_human(text, allow_typos=True, wpm=wpm), target)
    except Exception as exc:
        print(f"\n[Paste mode error] {exc}")
    return True


# ============================================================
# MAIN MENU
# ============================================================


def dripwriter():
    print("======== Drip Writer ========")
    if INPUT_BACKEND == "firefox":
        print(f"Input backend: FIREFOX (WebDriver)   Editor: {EDITOR}")
        print(f"  Start: {START_KEY.upper()}   Stop: hold {STOP_KEY.upper()}")
        print("  You can use your other monitors while it types.")
    else:
        print(f"Input backend: {INPUT_BACKEND}   Editor: {EDITOR}   "
              f"Start: {START_KEY.upper()}   Stop: hold {STOP_KEY.upper()}")
    print(f"DOCX support: {'available' if DOCX_AVAILABLE else 'missing (py -m pip install python-docx)'}")
    print(f"Default reference mode: {DEFAULT_REFERENCES_MODE}")
    print("\nModes:")
    print("  1 = Paste/type text")
    print("  2 = Recreate a DOCX (default reference behavior)")
    print("  3 = Recreate a DOCX (references LAST)")
    print("  4 = Analyze a DOCX only")
    print("  5 = Compare a finished copy with the original DOCX")
    print("  exit = Quit")

    while True:
        choice = input("\nChoose mode [1/2/3/4/5]: ").strip().lower()
        if choice == "exit":
            print("Goodbye!")
            return
        if choice == "1":
            if not paste_workflow():
                return
        elif choice == "2":
            if DOCX_MODE_ENABLED:
                docx_workflow(analyze_only=False, references_mode=DEFAULT_REFERENCES_MODE)
            else:
                print("DOCX mode is turned off in the settings.")
        elif choice == "3":
            if DOCX_MODE_ENABLED:
                docx_workflow(analyze_only=False, references_mode="last")
            else:
                print("DOCX mode is turned off in the settings.")
        elif choice == "4":
            docx_workflow(analyze_only=True)
        elif choice == "5":
            compare_workflow()
        else:
            print("Please choose 1, 2, 3, 4, 5, or exit.")


if __name__ == "__main__":
    try:
        choose_input_backend()
        dripwriter()
    except KeyboardInterrupt:
        print("\nGoodbye!")
    finally:
        if FIREFOX_DRIVER is not None:
            try:
                FIREFOX_DRIVER.quit()
            except Exception:
                pass
        cleanup_firefox_profile_copy()
        if INPUT_BACKEND != "firefox":
            release_modifiers()