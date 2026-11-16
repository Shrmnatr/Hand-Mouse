#!/usr/bin/env python3
"""
hand_mouse_pro.py - Hand Mouse with app profiles and presentation mode.
Multi-monitor, one or two hands, command gestures, steady click, per-app
gestures (PowerPoint, Google Slides, Teams, Zoom, YouTube, Spotify,
SolidWorks) and a presentation mode with laser pointer and drawing.
(MediaPipe Tasks API, Python 3.12-3.14, Windows)

Settings and per-app gestures live in hand_mouse_settings.json, which is
created next to this program the first time it runs.

MOUSE (mouse hand):
  Open hand ............ move        Thumb + index ....... left click / drag
  Thumb + middle ....... right click Thumb + ring ........ middle click / drag
  Two fingers together . scroll      Fist ................ pause

PRESENTATION MODE: turns on by itself during a PowerPoint slide show or a
full-screen Google Slides presentation (or press Ctrl+Alt+P anywhere):
  Open hand swipe left / right ..... next / previous slide
  Thumbs up / down (hold) .......... next / previous slide
  Point index finger ............... laser pointer
  Point, then pinch thumb + index .. draw on the slide
  Rock sign (hold) ................. clear drawings
  Three fingers (hold) ............. black screen on/off
  Both peace signs (hold 1 s) ...... end the slide show
  Both fists (hold 1 s) ............ Hand Mouse off / on

SOLIDWORKS: ring pinch + move = rotate, command-hand fist + ring pinch =
pan, command-hand fist + index pinch = add to selection, plus view gestures.

GLOBAL HOTKEYS (work in any app):
  Ctrl+Alt+P ... presentation mode on/off
  Ctrl+Alt+H ... Hand Mouse on/off
  Ctrl+Alt+C ... clear drawings

PREVIEW WINDOW KEYS (click it first):
  0-9 screens, H swap hands, G commands on/off, P presentation mode,
  SPACE on/off, Q / ESC quit

FAR MODE (--far): HD camera + digital zoom that follows your hands.
"""
import argparse
import json
import math
import os
import subprocess
import sys
import time
import urllib.request
from collections import deque

IS_WINDOWS = sys.platform.startswith("win")

if IS_WINDOWS:
    import ctypes
    from ctypes import wintypes

    # Use real pixel coordinates even when Windows display scaling isn't 100%.
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    user32.GetForegroundWindow.restype = wintypes.HWND
    user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
    user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND,
                                                ctypes.POINTER(wintypes.DWORD)]
    user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
    user32.GetParent.restype = wintypes.HWND
    user32.GetParent.argtypes = [wintypes.HWND]
    user32.GetWindowLongW.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.SetWindowLongW.argtypes = [wintypes.HWND, ctypes.c_int, wintypes.LONG]
    user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int,
                                    ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                    wintypes.UINT]
    user32.GetAsyncKeyState.restype = ctypes.c_short
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel32.QueryFullProcessImageNameW.argtypes = [
        wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR,
        ctypes.POINTER(wintypes.DWORD)]
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]

import cv2
import mediapipe as mp
import numpy as np
import pyautogui
from mediapipe.tasks import python as mp_python
from mediapipe.tasks.python import vision

pyautogui.FAILSAFE = False   # this version uses its own multi-monitor safety checks
pyautogui.PAUSE = 0

# Where the settings file and model live (next to the exe when built)
APP_DIR = (os.path.dirname(sys.executable) if getattr(sys, "frozen", False)
           else os.path.dirname(os.path.abspath(__file__)))
SETTINGS_PATH = os.path.join(APP_DIR, "hand_mouse_settings.json")
MODEL_PATH = os.path.join(APP_DIR, "hand_landmarker.task")
MODEL_URL = ("https://storage.googleapis.com/mediapipe-models/hand_landmarker/"
             "hand_landmarker/float16/latest/hand_landmarker.task")

# =====================================================================
#  DEFAULT SETTINGS - written to hand_mouse_settings.json on first run.
#  Edit that file (not this code) to customise gestures and apps.
# =====================================================================
NEXT_SLIDE = {"label": "Next slide", "keys": ["right"], "do": "clear_ink"}
PREV_SLIDE = {"label": "Previous slide", "keys": ["left"], "do": "clear_ink"}

DEFAULT_SETTINGS = {
    "_help": [
        "Hand Mouse settings. Save this file and restart Hand Mouse after editing.",
        "An action can use: label (text shown), keys (pressed together),",
        "type (text to type), open (program, folder or website),",
        "repeat (true = repeats while held), do (clear_ink, toggle_presentation,",
        "toggle_enabled). Set a gesture to null to turn it off.",
        "Gesture names: THUMBS_UP, THUMBS_DOWN, ROCK, SHAKA, THREE,",
        "BOTH_L, BOTH_POINT, BOTH_PEACE, SWIPE_LEFT, SWIPE_RIGHT.",
        "Profiles are checked top to bottom; the first match wins. A profile",
        "matches by program (e.g. powerpnt.exe) and/or title_contains.",
        "Profile extras: command_fist_modifier (keys held while the command",
        "hand makes a fist), zoom_modifier (keys for two-hand zoom, default",
        "ctrl), zoom_reverse (true/false), labels (left, left_mod, middle,",
        "middle_mod: text shown while those buttons are held).",
        "swipe_distance: lower = shorter swipe needed (1.0 easy, 1.8 strict).",
        "ring_pinch_middle: thumb + ring finger = middle mouse button.",
        "Delete this file to get the default settings back."
    ],
    "mouse_hand": "right",
    "far_mode": False,
    "far_zoom": 2.5,
    "swipe_distance": 1.4,
    "ring_pinch_middle": True,
    "laser_gain": 0.18,
    "laser_color": "#ff2020",
    "pen_color": "#ffd000",
    "show_toasts": True,
    "global": {
        "THUMBS_UP":   {"label": "Volume up",   "keys": ["volumeup"],   "repeat": True},
        "THUMBS_DOWN": {"label": "Volume down", "keys": ["volumedown"], "repeat": True},
        "ROCK":        {"label": "Mute",        "keys": ["volumemute"]},
        "SHAKA":       {"label": "Start menu",  "keys": ["win"]},
        "THREE":       {"label": "Play/Pause",  "keys": ["playpause"]},
        "BOTH_L":      {"label": "Screenshot",  "keys": ["win", "shift", "s"]},
        "BOTH_POINT":  {"label": "Task view",   "keys": ["win", "tab"]},
        "BOTH_PEACE":  {"label": "Show desktop", "keys": ["win", "d"]},
        "SWIPE_LEFT":  {"label": "Next tab",     "keys": ["ctrl", "tab"]},
        "SWIPE_RIGHT": {"label": "Previous tab", "keys": ["ctrl", "shift", "tab"]}
    },
    "presentation": {
        "SWIPE_LEFT":  NEXT_SLIDE,
        "SWIPE_RIGHT": PREV_SLIDE,
        "THUMBS_UP":   NEXT_SLIDE,
        "THUMBS_DOWN": PREV_SLIDE,
        "THREE":       {"label": "Black screen", "keys": ["b"]},
        "ROCK":        {"label": "Clear drawings", "do": "clear_ink"},
        "BOTH_PEACE":  {"label": "End slide show", "keys": ["esc"]}
    },
    "profiles": [
        {"name": "PowerPoint", "program": "powerpnt.exe",
         "gestures": {"BOTH_PEACE": {"label": "Start slide show", "keys": ["f5"]}}},
        {"name": "Google Slides", "title_contains": "google slides",
         "gestures": {"BOTH_PEACE": {"label": "Start presenting",
                                     "keys": ["ctrl", "f5"]}}},
        {"name": "Teams", "program": ["ms-teams.exe", "teams.exe"],
         "gestures": {
             "ROCK":  {"label": "Mic mute",   "keys": ["ctrl", "shift", "m"]},
             "SHAKA": {"label": "Camera",     "keys": ["ctrl", "shift", "o"]},
             "THREE": {"label": "Raise hand", "keys": ["ctrl", "shift", "k"]}}},
        {"name": "Zoom", "program": "zoom.exe",
         "gestures": {
             "ROCK":  {"label": "Mic mute",   "keys": ["alt", "a"]},
             "SHAKA": {"label": "Camera",     "keys": ["alt", "v"]},
             "THREE": {"label": "Raise hand", "keys": ["alt", "y"]}}},
        {"name": "YouTube", "title_contains": "youtube",
         "gestures": {
             "THREE":       {"label": "Play/Pause",  "keys": ["k"]},
             "SWIPE_LEFT":  {"label": "Forward 10s", "keys": ["l"]},
             "SWIPE_RIGHT": {"label": "Back 10s",    "keys": ["j"]},
             "BOTH_PEACE":  {"label": "Full screen", "keys": ["f"]}}},
        {"name": "Spotify", "program": "spotify.exe",
         "gestures": {
             "SWIPE_LEFT":  {"label": "Next song",     "keys": ["nexttrack"]},
             "SWIPE_RIGHT": {"label": "Previous song", "keys": ["prevtrack"]}}},
        {"name": "SolidWorks", "program": "sldworks.exe",
         "command_fist_modifier": ["ctrl"],
         "zoom_modifier": [],
         "zoom_reverse": False,
         "labels": {"middle": "ROTATE", "middle_mod": "PAN",
                    "left": "SELECT / DRAG", "left_mod": "ADD TO SELECTION"},
         "gestures": {
             "THUMBS_UP":   {"label": "Zoom to fit",     "keys": ["f"]},
             "THUMBS_DOWN": {"label": "Clear selection", "keys": ["esc"]},
             "ROCK":        {"label": "View selector",   "keys": ["space"]},
             "SHAKA":       {"label": "Normal to",       "keys": ["ctrl", "8"]},
             "THREE":       {"label": "Isometric",       "keys": ["ctrl", "7"]},
             "BOTH_POINT":  {"label": "Front view",      "keys": ["ctrl", "1"]},
             "BOTH_L":      {"label": "Top view",        "keys": ["ctrl", "5"]},
             "BOTH_PEACE":  {"label": "Hide component",  "keys": ["tab"]},
             "SWIPE_LEFT":  {"label": "Turn 90 left",    "keys": ["shift", "left"]},
             "SWIPE_RIGHT": {"label": "Turn 90 right",   "keys": ["shift", "right"]}}}
    ]
}

POSE_NAMES = {
    "THUMBS_UP": "Thumbs up",
    "THUMBS_DOWN": "Thumbs down",
    "ROCK": "Rock sign (index + pinky up)",
    "SHAKA": "Shaka (thumb + pinky out)",
    "THREE": "Three fingers (index, middle, ring)",
    "BOTH_L": "Both hands L-shape (hold 1 s)",
    "BOTH_POINT": "Both index fingers up (hold 1 s)",
    "BOTH_PEACE": "Both peace signs (hold 1 s)",
    "SWIPE_LEFT": "Both open hands swipe left",
    "SWIPE_RIGHT": "Both open hands swipe right",
}
PRES_NAMES = dict(POSE_NAMES, SWIPE_LEFT="Open hand swipe left",
                  SWIPE_RIGHT="Open hand swipe right")

# Advanced tuning (code only)
STEADY_CLICK = True       # command hand can click while the mouse hand aims
STEADY_DELAY = 0.25       # wait before a command-hand pinch holds the button
ZOOM_ENABLED = True       # two-hand pinch zoom
ZOOM_STEP = 0.04          # hand movement (part of camera width) per zoom step
SWIPE_HANDS = 1.4         # default swipe distance, measured in hand-lengths
SWIPE_WINDOW = 0.6        # ...within this many seconds
SWIPE_COOLDOWN = 0.5      # wait this long before another swipe
SWIPE_RETURN_BLOCK = 1.2  # ignore the opposite direction this long after a swipe
SWIPE_GAP = 0.25          # tracking can drop out this long without spoiling a swipe
POINT_DELAY = 0.15        # finger must point this long before the laser appears
HOLD_SECONDS = 0.6        # hold time for one-hand command gestures
TWO_HAND_HOLD = 1.0       # hold time for two-hand hold gestures
REPEAT_SECONDS = 0.2      # repeat speed for "repeat": true gestures
GRACE_SECONDS = 0.25      # ignore tracking flickers shorter than this
LASER_GRACE = 0.4         # keep the laser visible this long after pointing stops
THUMB_MARGIN = 0.25       # thumbs up/down: how far the thumb must stick out
THUMB_OUT = 0.6           # thumb at least this far out = shaka / L-shape
THUMB_IN = 0.45           # thumb closer than this = tucked in (pointing)
PEACE_SPREAD = 0.45       # index/middle tips this far apart = peace sign
REGION_X = (0.20, 0.80)   # camera area that maps to the screen(s)
REGION_Y = (0.15, 0.65)
SCROLL_UNIT = 120 if IS_WINDOWS else 1
TAKEOVER_PX = 25          # real mouse moved this far = user took over
TAKEOVER_SECONDS = 1.5    # how long hand control pauses afterwards

# MediaPipe hand landmark indices
WRIST = 0
THUMB_TIP = 4
INDEX_MCP, INDEX_PIP, INDEX_TIP = 5, 6, 8
MIDDLE_MCP, MIDDLE_PIP, MIDDLE_TIP = 9, 10, 12
RING_PIP, RING_TIP = 14, 16
PINKY_PIP, PINKY_TIP = 18, 20
NON_THUMB = [0] + list(range(5, 21))

HAND_CONNECTIONS = [
    (0, 1), (1, 2), (2, 3), (3, 4),
    (0, 5), (5, 6), (6, 7), (7, 8),
    (5, 9), (9, 10), (10, 11), (11, 12),
    (9, 13), (13, 14), (14, 15), (15, 16),
    (13, 17), (0, 17), (17, 18), (18, 19), (19, 20),
]

# Windows constants for the see-through overlay window
GWL_EXSTYLE = -20
WS_EX_LAYERED, WS_EX_TRANSPARENT = 0x80000, 0x20
WS_EX_TOOLWINDOW, WS_EX_NOACTIVATE = 0x80, 0x08000000
SWP_NOZORDER, SWP_NOACTIVATE, SWP_SHOWWINDOW = 0x4, 0x10, 0x40


# ---------------------------------------------------------------------
#  Settings
# ---------------------------------------------------------------------
def load_settings():
    settings = json.loads(json.dumps(DEFAULT_SETTINGS))   # deep copy
    default_names = [p["name"] for p in DEFAULT_SETTINGS["profiles"]]
    if not os.path.exists(SETTINGS_PATH):
        try:
            data = json.loads(json.dumps(DEFAULT_SETTINGS))
            data["_known_profiles"] = default_names
            with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            print(f"Created settings file: {SETTINGS_PATH}")
        except OSError as e:
            print(f"Could not create the settings file ({e}). Using defaults.")
        return settings
    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            user = json.load(f)
    except Exception as e:
        print(f"\nCould not read {SETTINGS_PATH}:\n  {e}")
        print("Using the built-in defaults. Fix the file (or delete it) and restart.\n")
        return settings
    for key, value in user.items():
        if key in ("global", "presentation") and isinstance(value, dict):
            settings[key].update(value)
        else:
            settings[key] = value

    # Add any new built-in app profiles (e.g. SolidWorks) to an older file,
    # but never re-add a profile you deleted yourself.
    profiles = settings.get("profiles")
    if isinstance(profiles, list):
        have = {p.get("name") for p in profiles if isinstance(p, dict)}
        known = set(user.get("_known_profiles") or have)
        new = [p for p in DEFAULT_SETTINGS["profiles"]
               if p["name"] not in have and p["name"] not in known]
        if new or "_known_profiles" not in user:
            settings["profiles"] = profiles + json.loads(json.dumps(new))
            user["profiles"] = settings["profiles"]
            user["_known_profiles"] = sorted(known | set(default_names))
            try:
                with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
                    json.dump(user, f, indent=2)
                if new:
                    print("Added new app profile(s) to your settings file:",
                          ", ".join(p["name"] for p in new))
            except OSError as e:
                print(f"Could not update the settings file ({e}).")
    print(f"Loaded settings from {SETTINGS_PATH}")
    return settings


def actions_for(settings, profile):
    acts = dict(settings.get("global") or {})
    if profile:
        acts.update(profile.get("gestures") or {})
    return acts


def as_key_list(value):
    if not value:
        return []
    return [value] if isinstance(value, str) else list(value)


# ---------------------------------------------------------------------
#  Which app is in front?
# ---------------------------------------------------------------------
def foreground_window():
    """Program name, title, class and position of the window in front."""
    if not IS_WINDOWS:
        return None
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return None
    n = user32.GetWindowTextLengthW(hwnd)
    title = ctypes.create_unicode_buffer(n + 1)
    user32.GetWindowTextW(hwnd, title, n + 1)
    cls = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, cls, 256)
    pid = wintypes.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    program = ""
    hproc = kernel32.OpenProcess(0x1000, False, pid.value)
    if hproc:
        size = wintypes.DWORD(1024)
        path = ctypes.create_unicode_buffer(1024)
        if kernel32.QueryFullProcessImageNameW(hproc, 0, path, ctypes.byref(size)):
            program = os.path.basename(path.value).lower()
        kernel32.CloseHandle(hproc)
    rect = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    return {"program": program, "title": title.value, "class": cls.value,
            "rect": (rect.left, rect.top, rect.right, rect.bottom)}


def match_profile(win, profiles):
    if not win:
        return None
    for p in profiles or []:
        if not isinstance(p, dict):
            continue
        progs, titles = p.get("program"), p.get("title_contains")
        if not progs and not titles:
            continue
        if progs:
            progs = [progs] if isinstance(progs, str) else progs
            if win["program"] not in [x.lower() for x in progs]:
                continue
        if titles:
            titles = [titles] if isinstance(titles, str) else titles
            if not any(t.lower() in win["title"].lower() for t in titles):
                continue
        return p
    return None


def covers_monitor(rect, monitors):
    l, t, r, b = rect
    return any(l <= m[0] + 2 and t <= m[1] + 2 and r >= m[2] - 2 and b >= m[3] - 2
               for m in monitors)


def detect_presenting(win, monitors):
    """Returns 'PowerPoint' / 'Google Slides' while a presentation is running."""
    if not win:
        return None
    title = win["title"].lower()
    if win["program"] == "powerpnt.exe" and (
            win["class"] == "screenClass" or "slide show" in title
            or "presenter view" in title):
        return "PowerPoint"
    if "google slides" in title and covers_monitor(win["rect"], monitors):
        return "Google Slides"
    return None


def monitor_of(rect, monitors, fallback):
    cx, cy = (rect[0] + rect[2]) / 2, (rect[1] + rect[3]) / 2
    for m in monitors:
        if m[0] <= cx < m[2] and m[1] <= cy < m[3]:
            return m
    return fallback


# ---------------------------------------------------------------------
#  Mouse control
# ---------------------------------------------------------------------
class WindowsMouse:
    """Talks to Windows directly, so it can reach every monitor."""
    LEFTDOWN, LEFTUP = 0x0002, 0x0004
    RIGHTDOWN, RIGHTUP = 0x0008, 0x0010
    MIDDLEDOWN, MIDDLEUP = 0x0020, 0x0040
    WHEEL = 0x0800
    MOVE, ABSOLUTE, VIRTUALDESK = 0x0001, 0x8000, 0x4000

    def __init__(self):
        self.user32 = ctypes.windll.user32
        self.buttons_down = 0

    def monitors(self):
        found = []
        proc_type = ctypes.WINFUNCTYPE(
            ctypes.c_int, wintypes.HANDLE, wintypes.HDC,
            ctypes.POINTER(wintypes.RECT), wintypes.LPARAM)

        def callback(hmon, hdc, rect, data):
            r = rect.contents
            found.append((r.left, r.top, r.right, r.bottom))
            return 1

        self.user32.EnumDisplayMonitors(None, None, proc_type(callback), 0)
        return found

    def position(self):
        pt = wintypes.POINT()
        self.user32.GetCursorPos(ctypes.byref(pt))
        return pt.x, pt.y

    def move(self, x, y):
        if self.buttons_down:
            # While a button is held, move like a real mouse so drag-based
            # tools (CAD rotate/pan, window dragging) see the movement.
            vx = self.user32.GetSystemMetrics(76)
            vy = self.user32.GetSystemMetrics(77)
            vw = self.user32.GetSystemMetrics(78)
            vh = self.user32.GetSystemMetrics(79)
            nx = int(round((x - vx) * 65535 / max(vw - 1, 1)))
            ny = int(round((y - vy) * 65535 / max(vh - 1, 1)))
            self._event(self.MOVE | self.ABSOLUTE | self.VIRTUALDESK, dx=nx, dy=ny)
        else:
            self.user32.SetCursorPos(int(x), int(y))

    def _event(self, flags, data=0, dx=0, dy=0):
        self.user32.mouse_event(flags, dx, dy, data, 0)

    def down(self):
        self.buttons_down += 1
        self._event(self.LEFTDOWN)

    def up(self):
        self._event(self.LEFTUP)
        self.buttons_down = max(0, self.buttons_down - 1)

    def middle_down(self):
        self.buttons_down += 1
        self._event(self.MIDDLEDOWN)

    def middle_up(self):
        self._event(self.MIDDLEUP)
        self.buttons_down = max(0, self.buttons_down - 1)

    def right_click(self):
        self._event(self.RIGHTDOWN)
        self._event(self.RIGHTUP)

    def scroll(self, amount):
        self._event(self.WHEEL, int(amount))


class PyAutoGuiMouse:
    """Fallback for Mac/Linux."""
    def monitors(self):
        w, h = pyautogui.size()
        return [(0, 0, w, h)]

    def position(self):
        p = pyautogui.position()
        return p[0], p[1]

    def move(self, x, y):
        pyautogui.moveTo(int(x), int(y))

    def down(self):
        pyautogui.mouseDown()

    def up(self):
        pyautogui.mouseUp()

    def middle_down(self):
        pyautogui.mouseDown(button="middle")

    def middle_up(self):
        pyautogui.mouseUp(button="middle")

    def right_click(self):
        pyautogui.click(button="right")

    def scroll(self, amount):
        pyautogui.scroll(int(amount))


def target_rect(monitors, choice):
    """choice 0 = all monitors combined, 1..N = a single monitor."""
    if choice == 0:
        return (min(m[0] for m in monitors), min(m[1] for m in monitors),
                max(m[2] for m in monitors), max(m[3] for m in monitors))
    return monitors[choice - 1]


def at_outer_corner(pos, monitors, margin=2):
    """True only at a corner of the desktop that no other monitor touches."""
    x, y = pos
    for i, (l, t, r, b) in enumerate(monitors):
        if l <= x < r and t <= y < b:
            near_x = x <= l + margin or x >= r - 1 - margin
            near_y = y <= t + margin or y >= b - 1 - margin
            if not (near_x and near_y):
                return False
            pad = margin + 3
            for j, (l2, t2, r2, b2) in enumerate(monitors):
                if j != i and (l2 - pad <= x < r2 + pad and
                               t2 - pad <= y < b2 + pad):
                    return False
            return True
    return False


# ---------------------------------------------------------------------
#  Helpers
# ---------------------------------------------------------------------
def ip(p):
    return int(p[0]), int(p[1])


def ensure_model(path):
    """Download the hand-tracking model if it isn't already in the folder."""
    if os.path.exists(path) and os.path.getsize(path) > 0:
        return
    print("Downloading hand-tracking model (one time only, a few MB)...")
    tmp = path + ".part"
    try:
        urllib.request.urlretrieve(MODEL_URL, tmp)
        os.replace(tmp, path)
        print("Model downloaded.")
    except Exception as e:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise SystemExit(
            f"Could not download the model: {e}\n"
            f"Download it manually in your browser from:\n  {MODEL_URL}\n"
            f"and save it as:\n  {path}"
        )


def open_camera(index, far, width, height):
    if far:
        width, height = 1920, 1080
    cap = None
    if far and IS_WINDOWS:
        # DirectShow + MJPG is usually the fastest way to get 1080p on Windows
        cap = cv2.VideoCapture(index, cv2.CAP_DSHOW)
        if cap.isOpened():
            cap.set(cv2.CAP_PROP_FOURCC, cv2.VideoWriter_fourcc(*"MJPG"))
        else:
            cap.release()
            cap = None
    if cap is None:
        cap = cv2.VideoCapture(index)
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, width)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, height)
    return cap


def draw_hand(frame, pts, color, thickness=2):
    for a, b in HAND_CONNECTIONS:
        cv2.line(frame, ip(pts[a]), ip(pts[b]), color, thickness)
    for p in pts:
        cv2.circle(frame, ip(p), thickness + 1, (0, 128, 255), -1)


def open_thing(target):
    if IS_WINDOWS:
        os.startfile(target)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", target])
    else:
        subprocess.Popen(["xdg-open", target])


def run_action(action):
    """Press keys / type / open. Returns the action's 'do' value, if any."""
    label = action.get("label", "?")
    try:
        keys = as_key_list(action.get("keys"))
        if keys:
            pyautogui.hotkey(*keys)
        if action.get("type"):
            pyautogui.write(str(action["type"]), interval=0.01)
        if action.get("open"):
            open_thing(action["open"])
        print("Command:", label)
    except Exception as e:
        print(f"Could not run '{label}': {e}")
    return action.get("do")


def zoom(mouse, steps, mods=("ctrl",)):
    """Mouse wheel with optional held keys. Positive steps = wheel up."""
    mods = list(mods)
    for k in mods:
        pyautogui.keyDown(k)
    try:
        mouse.scroll(steps * SCROLL_UNIT)
    finally:
        for k in reversed(mods):
            pyautogui.keyUp(k)


# ---------------------------------------------------------------------
#  Hand shapes
# ---------------------------------------------------------------------
class Hand:
    """Measurements and shape checks for one detected hand."""
    def __init__(self, pts):
        self.pts = pts
        self.size = np.linalg.norm(pts[WRIST] - pts[MIDDLE_MCP]) + 1e-6
        self.index_up = self.finger_up(INDEX_TIP, INDEX_PIP)
        self.middle_up = self.finger_up(MIDDLE_TIP, MIDDLE_PIP)
        self.ring_up = self.finger_up(RING_TIP, RING_PIP)
        self.pinky_up = self.finger_up(PINKY_TIP, PINKY_PIP)
        self.curled = not (self.index_up or self.middle_up
                           or self.ring_up or self.pinky_up)
        self.is_open = (self.index_up and self.middle_up
                        and self.ring_up and self.pinky_up)
        self.d_thumb_index = self.dist(THUMB_TIP, INDEX_TIP)
        self.d_thumb_middle = self.dist(THUMB_TIP, MIDDLE_TIP)
        self.d_thumb_ring = self.dist(THUMB_TIP, RING_TIP)
        thumb_reach = self.dist(THUMB_TIP, INDEX_MCP)
        self.thumb_out = thumb_reach > THUMB_OUT
        self.index_only = (self.index_up and not self.middle_up
                           and not self.ring_up and not self.pinky_up)
        self.is_l = self.index_only and self.thumb_out
        self.is_point = self.index_only and thumb_reach < THUMB_IN
        self.is_peace = (self.index_up and self.middle_up and not self.ring_up
                         and not self.pinky_up
                         and self.dist(INDEX_TIP, MIDDLE_TIP) > PEACE_SPREAD)
        self.pinch_point = (pts[THUMB_TIP] + pts[INDEX_TIP]) / 2
        self.pose = self._command_pose()

    def dist(self, a, b):
        return np.linalg.norm(self.pts[a] - self.pts[b]) / self.size

    def finger_up(self, tip, pip):
        return self.dist(WRIST, tip) > self.dist(WRIST, pip)

    def _command_pose(self):
        others_y = self.pts[NON_THUMB, 1]
        thumb_y = self.pts[THUMB_TIP][1]
        margin = THUMB_MARGIN * self.size
        if self.curled and thumb_y < others_y.min() - margin:
            return "THUMBS_UP"
        if self.curled and thumb_y > others_y.max() + margin:
            return "THUMBS_DOWN"
        if self.index_up and self.pinky_up and not self.middle_up and not self.ring_up:
            return "ROCK"
        if (self.pinky_up and not self.index_up and not self.middle_up
                and not self.ring_up and self.thumb_out):
            return "SHAKA"
        if self.index_up and self.middle_up and self.ring_up and not self.pinky_up:
            return "THREE"
        return None


def fingers_up(hd):
    """How many of the four fingers (not the thumb) are straight."""
    return int(hd.index_up) + int(hd.middle_up) + int(hd.ring_up) + int(hd.pinky_up)


def pinch_vals(hd, ring_enabled=True):
    """Thumb-to-finger distances for index, middle and ring. Only the finger
    closest to the thumb counts (the others get 9.0) so pinches don't collide.
    A ring pinch only counts with the index and middle fingers up, so it can't
    be confused with the rock sign."""
    ring = hd.d_thumb_ring if (ring_enabled and hd.index_up and hd.middle_up) else 9.0
    d = (hd.d_thumb_index, hd.d_thumb_middle, ring)
    k = min(range(3), key=d.__getitem__)
    return tuple(v if i == k else 9.0 for i, v in enumerate(d))


TWO_HAND_HOLDS = [
    ("BOTH_L", lambda hd: hd.is_l),
    ("BOTH_POINT", lambda hd: hd.is_point),
    ("BOTH_PEACE", lambda hd: hd.is_peace),
]


# ---------------------------------------------------------------------
#  Filters and detectors
# ---------------------------------------------------------------------
class _LowPass:
    def __init__(self):
        self.y = None

    def __call__(self, x, alpha):
        self.y = x if self.y is None else alpha * x + (1 - alpha) * self.y
        return self.y


class OneEuroFilter:
    """Low jitter when still, low lag when moving fast."""
    def __init__(self, min_cutoff=1.0, beta=0.007, d_cutoff=1.0):
        self.min_cutoff, self.beta, self.d_cutoff = min_cutoff, beta, d_cutoff
        self.x_filt, self.dx_filt = _LowPass(), _LowPass()
        self.t_prev = None

    @staticmethod
    def _alpha(cutoff, dt):
        tau = 1.0 / (2 * math.pi * cutoff)
        return 1.0 / (1.0 + tau / dt)

    def __call__(self, x, t):
        if self.t_prev is None:
            self.t_prev = t
            self.dx_filt(0.0, 1.0)
            return self.x_filt(x, 1.0)
        dt = max(t - self.t_prev, 1e-6)
        self.t_prev = t
        dx = (x - self.x_filt.y) / dt
        edx = self.dx_filt(dx, self._alpha(self.d_cutoff, dt))
        cutoff = self.min_cutoff + self.beta * abs(edx)
        return self.x_filt(x, self._alpha(cutoff, dt))


class Hysteresis:
    """Pinch detector that doesn't flicker at the threshold."""
    def __init__(self, on, off):
        self.on, self.off, self.active = on, off, False

    def update(self, value):
        prev = self.active
        if not self.active and value < self.on:
            self.active = True
        elif self.active and value > self.off:
            self.active = False
        return self.active, (self.active and not prev), (prev and not self.active)

    def reset(self):
        self.active = False


class GestureTimer:
    """Command gestures only fire after being held steadily."""
    def __init__(self, hold, repeat_every, grace):
        self.hold, self.repeat_every, self.grace = hold, repeat_every, grace
        self.reset()

    def reset(self):
        self.pose = None
        self.start = self.last_fire = self.last_seen = 0.0
        self.fired = False

    def _progress(self, now):
        return 1.0 if self.fired else min((now - self.start) / self.hold, 1.0)

    def update(self, pose, now, repeat):
        if (pose is None and self.pose is not None
                and now - self.last_seen < self.grace):
            return self.pose, False, self._progress(now)
        if pose != self.pose:
            self.reset()
            self.pose, self.start = pose, now
        if pose is None:
            return None, False, 0.0
        self.last_seen = now
        if now - self.start < self.hold:
            return pose, False, self._progress(now)
        if not self.fired:
            self.fired, self.last_fire = True, now
            return pose, True, 1.0
        if repeat and now - self.last_fire >= self.repeat_every:
            self.last_fire = now
            return pose, True, 1.0
        return pose, False, 1.0


class SwipeDetector:
    """Sideways swipe detector.
    - Distance is measured in hand-lengths, so it works near or far.
    - Short tracking dropouts and blurry fingers don't spoil a swipe.
    - The movement back after a swipe is ignored, so it doesn't undo it.
    - progress (0..1) shows how close the current movement is to a swipe."""
    def __init__(self, distance=SWIPE_HANDS):
        self.distance = float(distance)
        self.hist = deque()
        self.last_t = 0.0
        self.ready_at = 0.0
        self.blocked = None
        self.blocked_until = 0.0
        self.progress = 0.0

    def reset(self):
        self.hist.clear()
        self.progress = 0.0

    def update(self, now, xs, ys, sizes, open_ok=True):
        if self.hist and now - self.last_t > SWIPE_GAP:
            self.hist.clear()
        self.last_t = now
        self.hist.append((now, tuple(xs), tuple(ys), bool(open_ok)))
        while now - self.hist[0][0] > SWIPE_WINDOW:
            self.hist.popleft()
        self.progress = 0.0
        if len(self.hist) < 3:
            return None

        open_share = sum(e[3] for e in self.hist) / len(self.hist)
        score = {}
        for direction in ("SWIPE_LEFT", "SWIPE_RIGHT"):
            worst = 9.0
            for i in range(len(xs)):
                px = [e[1][i] for e in self.hist]
                py = [e[2][i] for e in self.hist]
                if direction == "SWIPE_LEFT":
                    k = max(range(len(px)), key=px.__getitem__)
                    travel = px[k] - xs[i]      # moved left since the right-most point
                else:
                    k = min(range(len(px)), key=px.__getitem__)
                    travel = xs[i] - px[k]      # moved right since the left-most point
                rise = abs(ys[i] - py[k])
                ratio = travel / (self.distance * sizes[i])
                if travel <= 0 or rise > travel * 0.8:   # must be mostly sideways
                    ratio = 0.0
                worst = min(worst, ratio)
            score[direction] = worst

        direction = max(score, key=score.get)
        self.progress = min(score[direction], 1.0)
        if (score[direction] < 1.0 or open_share < 0.5 or now < self.ready_at
                or (direction == self.blocked and now < self.blocked_until)):
            return None
        self.ready_at = now + SWIPE_COOLDOWN
        self.blocked = "SWIPE_RIGHT" if direction == "SWIPE_LEFT" else "SWIPE_LEFT"
        self.blocked_until = now + SWIPE_RETURN_BLOCK
        self.hist.clear()
        self.progress = 0.0
        return direction


class CropTracker:
    """Far mode: digital zoom that follows your hands, and searches for
    them when they're lost."""
    SPOTS = [(0.5, 0.45), (0.25, 0.45), (0.75, 0.45)]

    def __init__(self, zoom_factor):
        self.zoom = max(1.0, float(zoom_factor))
        self.c = None
        self.last_seen = 0.0
        self.scan_i = 0
        self.next_scan = 0.0
        self.box = (0, 0, 1, 1)

    def crop(self, frame):
        h, w = frame.shape[:2]
        cw, ch = int(w / self.zoom), int(h / self.zoom)
        if self.c is None:
            self.c = [w / 2, h / 2]
        x0 = int(min(max(self.c[0] - cw / 2, 0), w - cw))
        y0 = int(min(max(self.c[1] - ch / 2, 0), h - ch))
        self.box = (x0, y0, cw, ch)
        return frame[y0:y0 + ch, x0:x0 + cw], x0, y0

    def update(self, points, now, w, h):
        if points:
            tx = sum(p[0] for p in points) / len(points)
            ty = sum(p[1] for p in points) / len(points)
            self.c[0] += (tx - self.c[0]) * 0.15
            self.c[1] += (ty - self.c[1]) * 0.15
            self.last_seen = now
        elif now - self.last_seen > 1.5 and now >= self.next_scan:
            sx, sy = self.SPOTS[self.scan_i % len(self.SPOTS)]
            self.scan_i += 1
            self.c = [sx * w, sy * h]
            self.next_scan = now + 0.5


class HotkeyWatcher:
    """Global hotkeys that work in any app."""
    COMBOS = {"present": (0x11, 0x12, 0x50),   # Ctrl+Alt+P
              "pause":   (0x11, 0x12, 0x48),   # Ctrl+Alt+H
              "clear":   (0x11, 0x12, 0x43)}   # Ctrl+Alt+C

    def __init__(self):
        self.down = {k: False for k in self.COMBOS}

    def poll(self):
        fired = []
        for name, keys in self.COMBOS.items():
            pressed = all(user32.GetAsyncKeyState(k) & 0x8000 for k in keys)
            if pressed and not self.down[name]:
                fired.append(name)
            self.down[name] = pressed
        return fired


class Overlay:
    """See-through, click-through window over every monitor, used for the
    laser pointer, drawings and short on-screen messages."""
    KEY = "#010101"   # this colour is see-through

    def __init__(self, bounds, laser_color, pen_color):
        import tkinter as tk
        L, T, R, B = bounds
        self.ox, self.oy = L, T
        self.root = tk.Tk()
        self.root.title("Hand Mouse overlay")
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.configure(bg=self.KEY)
        self.root.attributes("-transparentcolor", self.KEY)
        self.root.geometry(f"{R - L}x{B - T}+0+0")
        self.laser_color = self._color(laser_color, "#ff2020")
        self.pen_color = self._color(pen_color, "#ffd000")
        self.canvas = tk.Canvas(self.root, bg=self.KEY, highlightthickness=0, bd=0)
        self.canvas.pack(fill="both", expand=True)
        self.root.update()
        hwnd = user32.GetParent(self.root.winfo_id()) or self.root.winfo_id()
        ex = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        user32.SetWindowLongW(hwnd, GWL_EXSTYLE, ex | WS_EX_LAYERED | WS_EX_TRANSPARENT
                              | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE)
        user32.SetWindowPos(hwnd, None, L, T, R - L, B - T,
                            SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW)
        self.root.update()
        self.last_pen = None
        self.laser_drawn = None
        self.toast_until = 0.0

    def _color(self, color, fallback):
        try:
            self.root.winfo_rgb(color)
            return color
        except Exception:
            return fallback

    def set_laser(self, pos):
        key = None if pos is None else (int(pos[0]), int(pos[1]))
        if key == self.laser_drawn:
            return
        self.laser_drawn = key
        c = self.canvas
        c.delete("laser")
        if key is None:
            return
        x, y = key[0] - self.ox, key[1] - self.oy
        c.create_oval(x - 20, y - 20, x + 20, y + 20, outline=self.laser_color,
                      width=3, tags="laser")
        c.create_oval(x - 10, y - 10, x + 10, y + 10, fill=self.laser_color,
                      outline="", tags="laser")

    def pen_to(self, pos):
        x, y = pos[0] - self.ox, pos[1] - self.oy
        if self.last_pen is not None:
            self.canvas.create_line(self.last_pen[0], self.last_pen[1], x, y,
                                    fill=self.pen_color, width=6,
                                    capstyle="round", tags="ink")
        self.last_pen = (x, y)

    def pen_up(self):
        self.last_pen = None

    def clear_ink(self):
        self.canvas.delete("ink")
        self.last_pen = None

    def toast(self, text, pos, now):
        c = self.canvas
        c.delete("toast")
        x, y = pos[0] - self.ox, pos[1] - self.oy
        t = c.create_text(x, y, text=text, fill="white",
                          font=("Segoe UI", 16, "bold"), tags="toast")
        x1, y1, x2, y2 = c.bbox(t)
        c.create_rectangle(x1 - 14, y1 - 8, x2 + 14, y2 + 8, fill="#202020",
                           outline="#505050", tags="toast")
        c.tag_raise(t)
        self.toast_until = now + 1.2

    def pump(self, now):
        if self.toast_until and now > self.toast_until:
            self.canvas.delete("toast")
            self.toast_until = 0.0
        self.root.update()

    def close(self):
        try:
            self.root.destroy()
        except Exception:
            pass


# ---------------------------------------------------------------------
#  Main program
# ---------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Hand Mouse Pro")
    ap.add_argument("--camera", type=int, default=0, help="camera index")
    ap.add_argument("--width", type=int, default=640)
    ap.add_argument("--height", type=int, default=480)
    ap.add_argument("--far", action="store_true",
                    help="far / conference-room mode (HD camera + digital zoom)")
    ap.add_argument("--monitor", default="all",
                    help="'all' or a monitor number (1 = leftmost)")
    ap.add_argument("--mouse-hand", choices=["right", "left"], default=None)
    ap.add_argument("--pinch-on", type=float, default=0.22)
    ap.add_argument("--pinch-off", type=float, default=0.32)
    ap.add_argument("--scroll-speed", type=float, default=80.0)
    ap.add_argument("--min-cutoff", type=float, default=1.0)
    ap.add_argument("--beta", type=float, default=0.007)
    ap.add_argument("--no-commands", action="store_true")
    ap.add_argument("--no-steady-click", action="store_true")
    ap.add_argument("--no-preview", action="store_true")
    args = ap.parse_args()

    settings = load_settings()
    far = args.far or bool(settings.get("far_mode", False))
    mouse_side = args.mouse_hand or str(settings.get("mouse_hand", "right")).lower()
    if mouse_side not in ("right", "left"):
        mouse_side = "right"
    steady_on = STEADY_CLICK and not args.no_steady_click
    show_toasts = bool(settings.get("show_toasts", True))
    laser_gain = float(settings.get("laser_gain", 0.18))
    swipe_distance = float(settings.get("swipe_distance", SWIPE_HANDS))
    ring_middle = bool(settings.get("ring_pinch_middle", True))
    pres_acts = settings.get("presentation") or {}

    ensure_model(MODEL_PATH)

    # Monitors
    mouse = WindowsMouse() if IS_WINDOWS else PyAutoGuiMouse()
    monitors = sorted(mouse.monitors(), key=lambda m: (m[0], m[1]))
    if not monitors:
        raise SystemExit("No monitors found.")
    print(f"Found {len(monitors)} monitor(s), numbered left to right:")
    for i, (l, t, r, b) in enumerate(monitors, 1):
        tag = "  (main)" if (l, t) == (0, 0) else ""
        print(f"  {i}: {r - l} x {b - t} at position ({l}, {t}){tag}")
    primary = next((m for m in monitors if m[0] == 0 and m[1] == 0), monitors[0])

    if args.monitor.lower() == "all":
        choice = 0
    else:
        try:
            choice = int(args.monitor)
        except ValueError:
            choice = 0
        if not 0 <= choice <= len(monitors):
            print(f"No monitor {args.monitor}; using all monitors.")
            choice = 0
    target = target_rect(monitors, choice)

    # Camera
    cap = open_camera(args.camera, far, args.width, args.height)
    if not cap.isOpened():
        raise SystemExit(f"Could not open camera {args.camera}")
    cam_w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
    cam_h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
    print(f"Camera: {cam_w} x {cam_h}" + ("  (FAR MODE)" if far else ""))
    if far and cam_w < 1280:
        print("  Note: this camera didn't give an HD picture; far mode works best at 1080p.")

    # Hand tracker: up to two hands
    with open(MODEL_PATH, "rb") as f:
        model_bytes = f.read()
    options = vision.HandLandmarkerOptions(
        base_options=mp_python.BaseOptions(model_asset_buffer=model_bytes),
        running_mode=vision.RunningMode.VIDEO,
        num_hands=2,
        min_hand_detection_confidence=0.5 if far else 0.7,
        min_hand_presence_confidence=0.5 if far else 0.6,
        min_tracking_confidence=0.5 if far else 0.6,
    )
    landmarker = vision.HandLandmarker.create_from_options(options)
    crop = CropTracker(settings.get("far_zoom", 2.5)) if far else None

    overlay = None
    if IS_WINDOWS:
        try:
            overlay = Overlay(target_rect(monitors, 0),
                              settings.get("laser_color", "#ff2020"),
                              settings.get("pen_color", "#ffd000"))
        except Exception as e:
            print(f"On-screen overlay unavailable ({e}). "
                  "The laser will use the mouse pointer instead.")
    hotkeys = HotkeyWatcher() if IS_WINDOWS else None

    filt_x = OneEuroFilter(args.min_cutoff, args.beta)
    filt_y = OneEuroFilter(args.min_cutoff, args.beta)
    left_pinch = Hysteresis(args.pinch_on, args.pinch_off)
    right_pinch = Hysteresis(args.pinch_on, args.pinch_off)
    middle_pinch = Hysteresis(args.pinch_on, args.pinch_off)
    steady_left = Hysteresis(args.pinch_on, args.pinch_off)
    steady_right = Hysteresis(args.pinch_on, args.pinch_off)
    pen_pinch = Hysteresis(args.pinch_on, args.pinch_off)
    gesture = GestureTimer(HOLD_SECONDS, REPEAT_SECONDS, GRACE_SECONDS)
    two_timer = GestureTimer(TWO_HAND_HOLD, REPEAT_SECONDS, GRACE_SECONDS)
    swipe1 = SwipeDetector(swipe_distance)         # presentation: one hand
    swipe2 = SwipeDetector(swipe_distance + 0.4)   # normal mode: both hands

    state = {"mouse_down": False, "middle_down": False, "held_mods": [],
             "steady": None, "steady_t": 0.0,
             "mouse_block": False, "steady_block": False}
    last_right_click = 0.0
    scroll_ref_y = None
    scroll_acc = 0.0
    zoom_active = False
    zoom_ref = 0.0
    flash_text, flash_until = "", 0.0
    enabled = True
    commands_on = not args.no_commands
    last_cmd_src = None
    window_placed = False
    last_pos = None
    takeover_until = 0.0
    prev_t = time.time()
    fps = 0.0
    t0 = time.monotonic()
    last_ts = -1
    frame = None
    w = h = 1
    now = dt = 0.0

    # App / presentation state
    presenting = False
    pres_override = None
    pres_override_auto = False
    auto_pres = None
    cur_profile = None
    acts = actions_for(settings, None)
    app_name = ""
    pres_monitor = primary
    next_ctx = 0.0
    pointer_on = False
    point_since = None
    anchor = None
    last_point_t = 0.0
    laser_pos = None
    laser_fx = laser_fy = None

    def hold_mods(mods):
        """Hold modifier keys (e.g. Ctrl) while a mouse button is down."""
        if mods and not state["held_mods"]:
            for k in mods:
                pyautogui.keyDown(k)
            state["held_mods"] = list(mods)

    def drop_mods():
        if (state["held_mods"] and not state["mouse_down"]
                and not state["middle_down"]):
            for k in reversed(state["held_mods"]):
                pyautogui.keyUp(k)
            state["held_mods"] = []

    def release_all():
        if state["mouse_down"]:
            mouse.up()
            state["mouse_down"] = False
        if state["middle_down"]:
            mouse.middle_up()
            state["middle_down"] = False
        for k in reversed(state["held_mods"]):
            pyautogui.keyUp(k)
        state["held_mods"] = []
        left_pinch.reset()
        right_pinch.reset()
        middle_pinch.reset()
        steady_left.reset()
        steady_right.reset()
        state["steady"] = None

    def toast(text):
        if overlay and show_toasts and text:
            pl, pt_, pr, pb = primary
            overlay.toast(text, ((pl + pr) / 2, pb - 90), time.time())

    def toggle_enabled():
        nonlocal enabled
        enabled = not enabled
        release_all()
        print("Hand Mouse", "ON" if enabled else "OFF")
        toast("Hand Mouse ON" if enabled else "Hand Mouse OFF")

    def toggle_presentation():
        nonlocal pres_override, pres_override_auto, next_ctx
        pres_override = not presenting
        pres_override_auto = bool(auto_pres)
        next_ctx = 0.0

    def fire(action):
        if not action:
            return
        do = run_action(action)
        toast(action.get("label", ""))
        if do == "clear_ink" and overlay:
            overlay.clear_ink()
        elif do == "toggle_presentation":
            toggle_presentation()
        elif do == "toggle_enabled":
            toggle_enabled()

    def do_mouse(hd, clicks=True, mods=()):
        """Normal mouse control with one hand. Returns the mode text.
        mods = keys to hold (e.g. Ctrl) when a button is pressed."""
        nonlocal scroll_ref_y, scroll_acc, last_pos, last_right_click
        if hd.curled and not (state["mouse_down"] or state["middle_down"]):
            release_all()
            scroll_ref_y = None
            return "PAUSED (fist)"

        if (not state["mouse_down"] and not state["middle_down"]
                and hd.index_up and hd.middle_up and not hd.ring_up and not hd.pinky_up
                and hd.dist(INDEX_TIP, MIDDLE_TIP) < 0.35
                and hd.d_thumb_index > args.pinch_off
                and hd.d_thumb_middle > args.pinch_off):
            release_all()
            y = hd.pts[MIDDLE_MCP][1] / h
            if scroll_ref_y is None:
                scroll_ref_y = y
                scroll_acc = 0.0
            disp = scroll_ref_y - y
            dead = 0.03
            if abs(disp) > dead:
                rate = math.copysign(abs(disp) - dead, disp) * args.scroll_speed
                scroll_acc += rate * dt * SCROLL_UNIT
                n = int(scroll_acc)
                if n:
                    mouse.scroll(n)
                    scroll_acc -= n
            ry = int(scroll_ref_y * h)
            cv2.line(frame, (0, ry), (w, ry), (255, 200, 0), 2)
            return "SCROLL"

        scroll_ref_y = None
        L, T, R, B = target
        sx = np.interp(hd.pts[INDEX_MCP][0] / w, REGION_X, (L, R - 1))
        sy = np.interp(hd.pts[INDEX_MCP][1] / h, REGION_Y, (T, B - 1))
        sx = min(max(filt_x(sx, now), L), R - 1)
        sy = min(max(filt_y(sy, now), T), B - 1)
        mouse.move(sx, sy)
        last_pos = mouse.position()

        if state["mouse_block"]:
            if hd.d_thumb_index > args.pinch_off:
                state["mouse_block"] = False
            else:
                clicks = False
        if not clicks:
            left_pinch.reset()
            right_pinch.reset()
            middle_pinch.reset()
            cv2.circle(frame, ip(hd.pts[INDEX_MCP]), 8, (0, 255, 0), -1)
            return "MOVE"

        labels = (cur_profile or {}).get("labels") or {}
        di, dm, dr = pinch_vals(hd, ring_middle)
        l_active, l_press, l_release = left_pinch.update(di)
        _, m_press, m_release = middle_pinch.update(dr)

        # Left button (thumb + index)
        if l_press and not state["middle_down"]:
            hold_mods(mods)
            mouse.down()
            state["mouse_down"] = True
        elif l_release and state["mouse_down"]:
            mouse.up()
            state["mouse_down"] = False
            drop_mods()

        # Middle button (thumb + ring)
        if m_press and not state["mouse_down"]:
            hold_mods(mods)
            mouse.middle_down()
            state["middle_down"] = True
        elif m_release and state["middle_down"]:
            mouse.middle_up()
            state["middle_down"] = False
            drop_mods()

        held = ("+".join(k.upper() for k in state["held_mods"]))
        if state["mouse_down"]:
            name = labels.get("left_mod" if held else "left")
            mode = name or ("LEFT DOWN / DRAG" + (f" + {held}" if held else ""))
            right_pinch.reset()
            color = (0, 0, 255)
        elif state["middle_down"]:
            name = labels.get("middle_mod" if held else "middle")
            mode = name or ("MIDDLE DRAG" + (f" + {held}" if held else ""))
            right_pinch.reset()
            color = (255, 0, 255)
        else:
            mode = "MOVE"
            if mods:
                mode += f"  (fist = {'+'.join(k.upper() for k in mods)})"
            _, r_press, _ = right_pinch.update(dm)
            if r_press and now - last_right_click > 0.5:
                hold_mods(mods)
                mouse.right_click()
                drop_mods()
                last_right_click = now
                mode = "RIGHT CLICK"
            color = (0, 255, 0)
        cv2.circle(frame, ip(hd.pts[INDEX_MCP]), 8, color, -1)
        return mode

    def do_steady(hd):
        """Steady click: the command hand clicks while the mouse hand aims."""
        nonlocal last_right_click
        if state["steady_block"]:
            if hd.d_thumb_index > args.pinch_off:
                state["steady_block"] = False
            else:
                return ""
        ti, tm = hd.d_thumb_index, hd.d_thumb_middle
        left_val = ti if ti <= tm else 9.0
        right_val = tm if tm < ti else 9.0
        l_active, l_press, l_release = steady_left.update(left_val)
        if l_press and not state["middle_down"] and not state["mouse_down"]:
            state["steady"], state["steady_t"] = "pending", now
        if state["steady"] == "pending":
            if l_release:
                mouse.down()
                mouse.up()
                state["mouse_down"] = False
                state["steady"] = None
                return "STEADY CLICK"
            if now - state["steady_t"] >= STEADY_DELAY:
                mouse.down()
                state["mouse_down"] = True
                state["steady"] = "down"
            else:
                cv2.circle(frame, ip(hd.pinch_point), 8, (0, 255, 255), -1)
                return "STEADY CLICK..."
        if state["steady"] == "down":
            if l_release:
                mouse.up()
                state["mouse_down"] = False
                state["steady"] = None
                return "STEADY RELEASE"
            cv2.circle(frame, ip(hd.pinch_point), 8, (0, 0, 255), -1)
            return "STEADY HOLD / DRAG"
        if l_active:
            steady_right.reset()
            return ""
        _, r_press, _ = steady_right.update(right_val)
        if r_press and now - last_right_click > 0.5:
            mouse.right_click()
            last_right_click = now
            return "STEADY RIGHT CLICK"
        return ""

    # Startup summary
    print("\nHand Mouse PRO running.")
    print("  Hotkeys (any app): Ctrl+Alt+P presentation mode, "
          "Ctrl+Alt+H on/off, Ctrl+Alt+C clear drawings")
    print("  Preview window keys: 0-9 screens, H swap hands, G commands, "
          "P presentation, SPACE on/off, Q quit")
    print(f"  Mouse hand: {mouse_side}   Swipe distance: {swipe_distance}   "
          f"Ring pinch = middle button: {'on' if ring_middle else 'off'}")
    print("Normal gestures:")
    for name, action in (settings.get("global") or {}).items():
        if action:
            print(f"  {POSE_NAMES.get(name, name):40s} -> {action.get('label', '')}")
    print("Presentation gestures:")
    print(f"  {'Point index finger':40s} -> Laser pointer")
    print(f"  {'Point, then pinch thumb + index':40s} -> Draw")
    for name, action in pres_acts.items():
        if action:
            print(f"  {PRES_NAMES.get(name, name):40s} -> {action.get('label', '')}")
    print("App profiles:", ", ".join(p.get("name", "?")
                                     for p in settings.get("profiles", [])
                                     if isinstance(p, dict)))

    try:
        while True:
            ok, frame = cap.read()
            if not ok:
                print("Camera frame grab failed.")
                break

            frame = cv2.flip(frame, 1)
            h, w = frame.shape[:2]
            if crop is not None:
                sub, x0, y0 = crop.crop(frame)
            else:
                sub, x0, y0 = frame, 0, 0
            sh, sw = sub.shape[:2]
            rgb = cv2.cvtColor(sub, cv2.COLOR_BGR2RGB)
            mp_image = mp.Image(image_format=mp.ImageFormat.SRGB, data=rgb)

            ts = int((time.monotonic() - t0) * 1000)
            if ts <= last_ts:
                ts = last_ts + 1
            last_ts = ts
            result = landmarker.detect_for_video(mp_image, ts)

            now = time.time()
            dt = max(now - prev_t, 1e-6)
            prev_t = now
            fps = 0.9 * fps + 0.1 * (1.0 / dt)
            mode = "no hand"
            gesture_progress = 0.0
            show_laser = False
            drawing = False
            pres_branch = False

            # Hands in full-frame pixel coordinates
            hands = [Hand(np.array([(x0 + p.x * sw, y0 + p.y * sh) for p in lm]))
                     for lm in result.hand_landmarks[:2]]
            if crop is not None:
                crop.update([hd.pts[MIDDLE_MCP] for hd in hands], now, w, h)

            # Which app is in front? Presenting?
            if now >= next_ctx:
                next_ctx = now + 0.3
                win = foreground_window()
                profile = match_profile(win, settings.get("profiles"))
                auto_pres = detect_presenting(win, monitors)
                if pres_override is not None and bool(auto_pres) != pres_override_auto:
                    pres_override = None
                new_presenting = (pres_override if pres_override is not None
                                  else bool(auto_pres))
                if new_presenting != presenting:
                    presenting = new_presenting
                    release_all()
                    gesture.reset()
                    two_timer.reset()
                    swipe1.reset()
                    swipe2.reset()
                    pointer_on = False
                    point_since = None
                    zoom_active = False
                    if overlay:
                        overlay.clear_ink()
                    print("Presentation mode", "ON" if presenting else "OFF")
                    toast("Presentation mode ON" if presenting
                          else "Presentation mode OFF")
                if win:
                    pres_monitor = monitor_of(win["rect"], monitors, primary)
                if profile is not cur_profile and (state["mouse_down"]
                                                   or state["middle_down"]):
                    release_all()
                cur_profile = profile
                acts = actions_for(settings, profile)
                app_name = auto_pres or (profile.get("name", "") if profile else "")

            # Roles: mouse/presenter hand and command hand
            hands.sort(key=lambda hd: hd.pts[WRIST][0])
            mouse_hand = other_hand = None
            if len(hands) == 2:
                if mouse_side == "right":
                    other_hand, mouse_hand = hands
                else:
                    mouse_hand, other_hand = hands
            elif hands:
                mouse_hand = hands[0]

            thick = 2 if w <= 1000 else 4
            for hd in hands:
                is_mouse = hd is mouse_hand
                color = (255, 255, 255) if is_mouse else (255, 200, 120)
                draw_hand(frame, hd.pts, color, thick)
                if other_hand is not None:
                    wx, wy = ip(hd.pts[WRIST])
                    cv2.putText(frame, "MOUSE" if is_mouse else "COMMANDS",
                                (max(wx - 50, 0), min(wy + 30, h - 10)),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.35 * thick, color, thick)

            # Global hotkeys
            if hotkeys:
                for name in hotkeys.poll():
                    if name == "present":
                        toggle_presentation()
                    elif name == "pause":
                        toggle_enabled()
                    elif name == "clear" and overlay:
                        overlay.clear_ink()

            # Did the user move the real mouse?
            cur = mouse.position()
            if last_pos is not None and (abs(cur[0] - last_pos[0]) +
                                         abs(cur[1] - last_pos[1])) > TAKEOVER_PX:
                takeover_until = now + TAKEOVER_SECONDS
                release_all()
            last_pos = cur
            user_has_mouse = now < takeover_until
            if user_has_mouse and at_outer_corner(cur, monitors):
                print("Real mouse moved into a screen corner. Exiting.")
                break

            # Two-hand hold gestures (both fists always works)
            holds = pres_acts if presenting else acts
            busy_button = state["mouse_down"] or state["middle_down"]
            two_pose = None
            if other_hand is not None and not user_has_mouse:
                if (mouse_hand.curled and other_hand.curled and not busy_button
                        and mouse_hand.pose is None and other_hand.pose is None):
                    two_pose = "BOTH_FISTS"
                elif enabled and (presenting or commands_on) and not busy_button:
                    for name, test in TWO_HAND_HOLDS:
                        if holds.get(name) and test(mouse_hand) and test(other_hand):
                            two_pose = name
                            break
            two_active, two_fire, two_prog = two_timer.update(two_pose, now, False)
            if two_fire:
                if two_active == "BOTH_FISTS":
                    toggle_enabled()
                    gesture.reset()
                else:
                    fire(holds.get(two_active))

            # Normal mode two-hand swipe
            if (not presenting and other_hand is not None and enabled and commands_on
                    and not user_has_mouse and not two_active and not busy_button):
                d = swipe2.update(
                    now,
                    [mouse_hand.pts[MIDDLE_MCP][0], other_hand.pts[MIDDLE_MCP][0]],
                    [mouse_hand.pts[MIDDLE_MCP][1], other_hand.pts[MIDDLE_MCP][1]],
                    [mouse_hand.size, other_hand.size],
                    open_ok=fingers_up(mouse_hand) >= 3 and fingers_up(other_hand) >= 3)
                if d and acts.get(d):
                    fire(acts[d])
                    flash_text = acts[d].get("label", "")
                    flash_until = now + 0.8
            else:
                swipe2.reset()

            # Two-hand pinch zoom state (normal mode only)
            if presenting or other_hand is None or not enabled or user_has_mouse:
                zoom_active = False
            elif not (mouse_hand.d_thumb_index < args.pinch_off
                      and other_hand.d_thumb_index < args.pinch_off):
                zoom_active = False
            start_zoom = (not presenting and ZOOM_ENABLED and other_hand is not None
                          and not zoom_active and state["steady"] != "down"
                          and not state["middle_down"]
                          and mouse_hand.d_thumb_index < args.pinch_on
                          and other_hand.d_thumb_index < args.pinch_on)

            if user_has_mouse:
                mode = "REAL MOUSE - paused"
                gesture.reset()
                scroll_ref_y = None

            elif two_active:
                release_all()
                gesture.reset()
                scroll_ref_y = None
                zoom_active = False
                gesture_progress = two_prog
                if two_active == "BOTH_FISTS":
                    if two_prog < 1.0:
                        mode = f"Both fists: turn {'OFF' if enabled else 'ON'} (hold)"
                    else:
                        mode = f"Hand Mouse {'ON' if enabled else 'OFF'}"
                else:
                    label = (holds.get(two_active) or {}).get("label", two_active)
                    mode = label if two_prog >= 1.0 else f"{label} (hold)"

            elif not enabled:
                mode = "OFF (both fists, Space or Ctrl+Alt+H)"
                release_all()
                gesture.reset()
                scroll_ref_y = None

            elif presenting:
                # PRESENTATION MODE: never touches the mouse
                pres_branch = True
                release_all()
                scroll_ref_y = None
                hd = mouse_hand
                if hd is None:
                    gesture.reset()
                    pen_pinch.reset()
                    point_since = None
                    if pointer_on and now - last_point_t < LASER_GRACE:
                        show_laser = True
                    else:
                        pointer_on = False
                    mode = "PRESENTING"
                else:
                    pen_on, _, _ = pen_pinch.update(hd.d_thumb_index)
                    # The finger must point steadily for a moment before the laser
                    # appears, so a blurry swipe isn't mistaken for pointing
                    if hd.index_only:
                        if point_since is None:
                            point_since = now
                    else:
                        point_since = None
                    pointing = ((point_since is not None
                                 and now - point_since >= POINT_DELAY)
                                or (pointer_on and (pen_on or hd.index_only)))
                    if pointing:
                        knuckle = hd.pts[INDEX_MCP]
                        L, T, R, B = pres_monitor
                        span = (R - L) * laser_gain
                        if not pointer_on:
                            pointer_on = True
                            anchor = knuckle.copy()
                            laser_fx = OneEuroFilter(0.8, 0.02)
                            laser_fy = OneEuroFilter(0.8, 0.02)
                        last_point_t = now
                        rx = (L + R) / 2 + (knuckle[0] - anchor[0]) / hd.size * span
                        ry = (T + B) / 2 + (knuckle[1] - anchor[1]) / hd.size * span
                        cx = min(max(rx, L), R - 1)
                        cy = min(max(ry, T), B - 1)
                        # Pushing past an edge moves the anchor, so coming back
                        # responds straight away
                        anchor[0] += (rx - cx) / span * hd.size
                        anchor[1] += (ry - cy) / span * hd.size
                        laser_pos = (laser_fx(cx, now), laser_fy(cy, now))
                        show_laser = True
                        if pen_on and overlay:
                            overlay.pen_to(laser_pos)
                            drawing = True
                            mode = "DRAWING"
                        else:
                            mode = "LASER"
                        if overlay is None:
                            mouse.move(*laser_pos)
                            last_pos = mouse.position()
                        gesture.reset()
                        swipe1.reset()
                    else:
                        if pointer_on and now - last_point_t < LASER_GRACE:
                            show_laser = True
                        else:
                            pointer_on = False
                        mode = "PRESENTING"

                        # Swipes: tolerant of blurry fingers and short dropouts
                        d = swipe1.update(now, [hd.pts[MIDDLE_MCP][0]],
                                          [hd.pts[MIDDLE_MCP][1]], [hd.size],
                                          open_ok=fingers_up(hd) >= 3)
                        if d and pres_acts.get(d):
                            fire(pres_acts[d])
                            flash_text = pres_acts[d].get("label", "")
                            flash_until = now + 0.8

                        # Hold gestures (thumbs up/down, three fingers, rock)
                        pose = hd.pose or (other_hand.pose if other_hand else None)
                        if pose and not pres_acts.get(pose):
                            pose = None
                        act = pres_acts.get(pose) if pose else None
                        active, fire_now, gesture_progress = gesture.update(
                            pose, now, bool(act and act.get("repeat")))
                        if fire_now:
                            fire(pres_acts.get(active))
                        if active:
                            lab = (pres_acts.get(active) or {}).get("label", active)
                            mode = lab if gesture_progress >= 1.0 else f"{lab} (hold)"
                        elif swipe1.progress > 0.15:
                            gesture_progress = swipe1.progress
                            mode = "Swiping..."

            elif mouse_hand is None:
                release_all()
                gesture.reset()
                scroll_ref_y = None

            elif zoom_active or start_zoom:
                sep = float(np.linalg.norm(mouse_hand.pinch_point
                                           - other_hand.pinch_point)) / w
                if not zoom_active:
                    zoom_active, zoom_ref = True, sep
                    release_all()
                    gesture.reset()
                    scroll_ref_y = None
                state["mouse_block"] = state["steady_block"] = True
                steps = 0
                while sep - zoom_ref >= ZOOM_STEP:
                    steps += 1
                    zoom_ref += ZOOM_STEP
                while zoom_ref - sep >= ZOOM_STEP:
                    steps -= 1
                    zoom_ref -= ZOOM_STEP
                if steps:
                    zp = cur_profile or {}
                    zmods = as_key_list(zp.get("zoom_modifier", ["ctrl"]))
                    if zp.get("zoom_reverse"):
                        steps = -steps
                    zoom(mouse, steps, zmods)
                mode = "ZOOM (apart = in, together = out)"
                cv2.line(frame, ip(mouse_hand.pinch_point),
                         ip(other_hand.pinch_point), (0, 255, 255), thick)

            else:
                # NORMAL MODE
                if other_hand is None and state["steady"]:
                    release_all()
                steady_mode = ""
                steady_busy = False
                if other_hand is not None and steady_on:
                    steady_mode = do_steady(other_hand)
                    steady_busy = (state["steady"] is not None
                                   or steady_left.active or steady_right.active)

                # Command-hand fist = hold a key (e.g. Ctrl in SolidWorks)
                fist_mods = []
                if other_hand is not None and cur_profile:
                    fm = as_key_list(cur_profile.get("command_fist_modifier"))
                    if (fm and other_hand.curled and other_hand.pose is None
                            and other_hand.d_thumb_index > args.pinch_off):
                        fist_mods = fm

                cmd_src = "other" if other_hand is not None else "mouse"
                if cmd_src != last_cmd_src:
                    gesture.reset()
                    last_cmd_src = cmd_src
                cmd_hand = other_hand if other_hand is not None else mouse_hand
                buttons_held = state["mouse_down"] or state["middle_down"]
                pose = (cmd_hand.pose if (commands_on and not steady_busy
                                          and not buttons_held) else None)
                if pose and not acts.get(pose):
                    pose = None
                action = acts.get(pose) if pose else None
                active, fire_now, gesture_progress = gesture.update(
                    pose, now, bool(action and action.get("repeat")))
                if fire_now:
                    fire(acts.get(active))
                label = ""
                if active:
                    lab = (acts.get(active) or {}).get("label", active)
                    label = lab if gesture_progress >= 1.0 else f"{lab} (hold)"

                if active and other_hand is None:
                    release_all()
                    scroll_ref_y = None
                    mode = label
                else:
                    mode = do_mouse(mouse_hand, clicks=not steady_busy, mods=fist_mods)
                    if steady_mode:
                        mode = steady_mode
                    if label:
                        mode = f"{mode} | {label}"

            if not pres_branch:
                pointer_on = False
                point_since = None
            if now < flash_until:
                mode = f"{mode} | {flash_text}"

            # Overlay
            if overlay:
                overlay.set_laser(laser_pos if show_laser else None)
                if not drawing:
                    overlay.pen_up()
                overlay.pump(now)

            # Preview window
            if not args.no_preview:
                if crop is not None:
                    bx0, by0, bcw, bch = crop.box
                    cv2.rectangle(frame, (bx0, by0), (bx0 + bcw, by0 + bch),
                                  (0, 200, 255), thick)
                if not presenting:
                    x_0, x_1 = int(REGION_X[0] * w), int(REGION_X[1] * w)
                    y_0, y_1 = int(REGION_Y[0] * h), int(REGION_Y[1] * h)
                    cv2.rectangle(frame, (x_0, y_0), (x_1, y_1), (255, 0, 255), thick)

                pw = 480
                ph = int(h * pw / w)
                preview = cv2.resize(frame, (pw, ph), interpolation=cv2.INTER_AREA)
                cv2.putText(preview, mode, (8, 20),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 255), 1)
                if gesture_progress > 0:
                    bx, by, bw = 8, 28, 110
                    cv2.rectangle(preview, (bx, by),
                                  (bx + int(bw * gesture_progress), by + 6),
                                  (0, 255, 0), -1)
                    cv2.rectangle(preview, (bx, by), (bx + bw, by + 6),
                                  (255, 255, 255), 1)
                app_line = f"App: {app_name or '-'}"
                if presenting:
                    app_line += "  |  PRESENTING"
                cv2.putText(preview, app_line, (8, 52),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.45,
                            (0, 165, 255) if presenting else (255, 255, 0), 1)
                screen_label = "all" if choice == 0 else str(choice)
                cv2.putText(preview, f"Screen: {screen_label}   Commands: "
                            f"{'ON' if commands_on else 'OFF'} (G)",
                            (8, 70), cv2.FONT_HERSHEY_SIMPLEX, 0.4, (255, 255, 0), 1)
                cv2.putText(preview, f"Mouse hand: {mouse_side} (H)",
                            (8, 86), cv2.FONT_HERSHEY_SIMPLEX, 0.4, (255, 255, 0), 1)
                cv2.putText(preview, f"FPS {fps:4.1f}" + ("   FAR" if far else ""),
                            (8, ph - 8), cv2.FONT_HERSHEY_SIMPLEX, 0.4,
                            (200, 200, 200), 1)
                cv2.imshow("Hand Mouse Pro", preview)

                if not window_placed:
                    pl, pt_, pr, pb = primary
                    cv2.moveWindow("Hand Mouse Pro",
                                   pl + (pr - pl - pw) // 2, pt_ + (pb - pt_ - ph) // 2)
                    window_placed = True

                key = cv2.waitKey(1) & 0xFF
                if key in (ord("q"), 27):
                    break
                if key == ord(" "):
                    toggle_enabled()
                if key in (ord("g"), ord("G")):
                    commands_on = not commands_on
                    gesture.reset()
                    print("Command gestures:", "ON" if commands_on else "OFF")
                if key in (ord("p"), ord("P")):
                    toggle_presentation()
                if key in (ord("h"), ord("H")):
                    mouse_side = "left" if mouse_side == "right" else "right"
                    release_all()
                    gesture.reset()
                    filt_x = OneEuroFilter(args.min_cutoff, args.beta)
                    filt_y = OneEuroFilter(args.min_cutoff, args.beta)
                    print("Mouse hand:", mouse_side)
                if ord("0") <= key <= ord("9"):
                    c = key - ord("0")
                    if c <= len(monitors):
                        choice = c
                        target = target_rect(monitors, choice)
                        release_all()
                        filt_x = OneEuroFilter(args.min_cutoff, args.beta)
                        filt_y = OneEuroFilter(args.min_cutoff, args.beta)
                        print("Now controlling:",
                              "all monitors" if c == 0 else f"monitor {c}")

    except KeyboardInterrupt:
        pass
    finally:
        release_all()
        if overlay:
            overlay.close()
        landmarker.close()
        cap.release()
        cv2.destroyAllWindows()


if __name__ == "__main__":
    main()
