#!/usr/bin/env python3
"""Windows Python port of easyMoney.swift.

This is a practical first Windows version, not a one-to-one Swift rewrite.
The UI layer uses Windows UI Automation plus coordinate fallbacks; LLM commands
are kept usable without loading UI automation dependencies.
"""

from __future__ import annotations

import base64
import ctypes
import importlib
import os
import re
import sys
import time
from ctypes import wintypes
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Iterable, Optional


APP_NAME = "easyMoney Windows"
APP_VERSION = "0.1.0"
COMMENT_REFRESH_WAIT_SECONDS = 0.1
COMMENT_REFRESH_CAPTURE_INTERVAL_SECONDS = 0.012
COMMENT_REFRESH_IDLE_SECONDS = 0.003

INPUT_MOUSE = 0
INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004
DIRECT_TEXT_ENTRY_MAX_UTF16_UNITS = 512
DIRECT_TEXT_ENTRY_CHUNK_UTF16_UNITS = 64
MOUSEEVENTF_MOVE = 0x0001
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_ABSOLUTE = 0x8000
MOUSEEVENTF_VIRTUALDESK = 0x4000
SM_XVIRTUALSCREEN = 76
SM_YVIRTUALSCREEN = 77
SM_CXVIRTUALSCREEN = 78
SM_CYVIRTUALSCREEN = 79
ULONG_PTR = ctypes.c_ulonglong if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_ulong


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", wintypes.LONG),
        ("dy", wintypes.LONG),
        ("mouseData", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]


class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ("uMsg", wintypes.DWORD),
        ("wParamL", wintypes.WORD),
        ("wParamH", wintypes.WORD),
    ]


class INPUTUNION(ctypes.Union):
    _fields_ = [
        ("ki", KEYBDINPUT),
        ("mi", MOUSEINPUT),
        ("hi", HARDWAREINPUT),
    ]


class INPUT(ctypes.Structure):
    _fields_ = [
        ("type", wintypes.DWORD),
        ("union", INPUTUNION),
    ]


class EasyMoneyError(RuntimeError):
    pass


class WindowPositionUnavailable(EasyMoneyError):
    pass


class CaptureUnavailable(EasyMoneyError):
    pass


class UIAListItemUnavailable(EasyMoneyError):
    pass


@dataclass(frozen=True)
class Point:
    x: float
    y: float

    def rounded(self) -> tuple[int, int]:
        return int(round(self.x)), int(round(self.y))


@dataclass(frozen=True)
class Rect:
    left: float
    top: float
    right: float
    bottom: float

    @property
    def width(self) -> float:
        return max(0.0, self.right - self.left)

    @property
    def height(self) -> float:
        return max(0.0, self.bottom - self.top)

    @property
    def center(self) -> Point:
        return Point(self.left + self.width / 2, self.top + self.height / 2)

    def contains(self, point: Point) -> bool:
        return self.left <= point.x <= self.right and self.top <= point.y <= self.bottom

    def intersects_y(self, y: float, tolerance: float = 0) -> bool:
        return self.top - tolerance <= y <= self.bottom + tolerance

    def inset(self, dx: float, dy: float) -> "Rect":
        return Rect(self.left + dx, self.top + dy, self.right - dx, self.bottom - dy)

    def expanded(self, dx: float, dy: float) -> "Rect":
        return Rect(self.left - dx, self.top - dy, self.right + dx, self.bottom + dy)

    def clamp_to(self, outer: "Rect") -> "Rect":
        return Rect(
            max(self.left, outer.left),
            max(self.top, outer.top),
            min(self.right, outer.right),
            min(self.bottom, outer.bottom),
        )

    def to_mss(self) -> dict[str, int]:
        return {
            "left": int(round(self.left)),
            "top": int(round(self.top)),
            "width": max(1, int(round(self.width))),
            "height": max(1, int(round(self.height))),
        }

    def describe(self) -> str:
        return f"({int(self.left)},{int(self.top)}) {int(self.width)}x{int(self.height)}"


@dataclass(frozen=True)
class CaptureFrame:
    width: int
    height: int
    rgb: bytes


@dataclass
class CommentConfig:
    comment_from_action: Point
    send_x_ratio: float = 0.8
    send_from_action: Optional[Point] = None
    fixed_send_action_y_threshold: Optional[float] = None
    fixed_send_window_offset: Optional[Point] = None


@dataclass
class LocalLLMConfig:
    provider: str
    endpoint: str
    model: str
    api_key: Optional[str]
    timeout_seconds: float


@dataclass
class SolvedQuestion:
    answer: str
    evidence: str = ""
    confidence: float = 0.0
    source: str = "unknown"


@dataclass
class MomentPostResolution:
    body_frame: Rect
    action_point: Point
    text: str
    source: str


@dataclass
class UIAListItemResolution:
    item_index: int
    body_frame: Rect
    action_point: Point
    text: str
    expected_user_id: str
    detected_prefix: str
    elapsed_ms: int


@dataclass
class CommentOptions:
    comment_text: Optional[str] = None
    requested_user: str = ""
    solve_question: bool = False
    use_llm: bool = False
    use_vision: bool = False
    debug: bool = False
    save_post_image: bool = False
    save_path: Optional[Path] = None
    click_post_image: bool = False
    test_image_crop: bool = False
    rounds: int = 30
    submit_comment_keys_override: Optional[tuple[str, ...]] = None
    open_comment_mode: str = "keys"
    submit_mode: str = "click"


@dataclass
class CommentSendPlan:
    text: str
    action_point: Point
    comment_point: Point
    send_point: Point
    send_point_method: str
    open_comment_mode: str
    submit_mode: str
    open_comment_keys: tuple[str, ...]
    submit_comment_keys: tuple[str, ...]
    comment_open_method: str
    submit_method: str


@dataclass
class CommentSendResult:
    text_input_method: str
    action_click_ms: int
    open_comment_ms: int
    text_input_ms: int
    send_submit_ms: int
    total_send_ms: int
    send_step_label: str


HOME = Path.home()
EASYMONEY_DIR = HOME / ".easyMoney"
CONFIG_REFRESH = HOME / ".wechat_refresh_offset"
CONFIG_COMMENT = HOME / ".wechat_comment_config"
CONFIG_POST_IMAGE_TAP_OFFSET = HOME / ".wechat_post_image_tap_offset"
CONFIG_POST_IMAGE_TAP_X_OFFSET = HOME / ".wechat_post_image_tap_x_offset"
ACTION_TEMPLATE = HOME / ".wechat_action_tpl.png"
DEBUG_DIR = Path(os.environ.get("EASYMONEY_DEBUG_DIR", str(HOME / "test")))


def enable_dpi_awareness() -> None:
    if os.name != "nt":
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-monitor DPI aware.
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def require_module(module_name: str, pip_name: Optional[str] = None) -> Any:
    try:
        return importlib.import_module(module_name)
    except ImportError as exc:
        package = pip_name or module_name
        raise EasyMoneyError(f"缺少依赖 `{package}`，请先运行: python -m pip install -r requirements.txt") from exc


def expand_path(raw: str | Path | None) -> Optional[Path]:
    if raw is None:
        return None
    text = str(raw).strip()
    if not text:
        return None
    return Path(os.path.expandvars(os.path.expanduser(text)))


def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def atomic_write_text(path: Path, text: str) -> None:
    ensure_parent(path)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    tmp.replace(path)


def parse_point_text(text: str) -> Point:
    parts = [p.strip() for p in text.split(",")]
    if len(parts) != 2:
        raise EasyMoneyError(f"坐标格式应为 x,y: {text}")
    return Point(float(parts[0]), float(parts[1]))


def load_point(path: Path) -> Optional[Point]:
    try:
        return parse_point_text(path.read_text(encoding="utf-8").strip())
    except Exception:
        return None


def save_point(path: Path, point: Point) -> None:
    atomic_write_text(path, f"{point.x},{point.y}")


def load_float(path: Path) -> Optional[float]:
    try:
        return float(path.read_text(encoding="utf-8").strip())
    except Exception:
        return None


def save_float(path: Path, value: float) -> None:
    atomic_write_text(path, str(value))


def load_comment_config() -> Optional[CommentConfig]:
    if not CONFIG_COMMENT.exists():
        return None
    try:
        parts = [p.strip() for p in CONFIG_COMMENT.read_text(encoding="utf-8").strip().split(",")]
        values = [float(p) for p in parts if p != ""]
    except Exception:
        return None
    if len(values) < 2:
        return None
    comment_from_action = Point(values[0], values[1])
    send_x_ratio = values[2] if len(values) >= 3 else 0.8
    send_from_action = Point(values[3], values[4]) if len(values) >= 5 else None
    fixed_threshold = None
    fixed_offset = None
    if len(values) >= 8:
        fixed_threshold = values[5]
        fixed_offset = Point(values[6], values[7])
    elif len(values) >= 6 and send_from_action is None:
        fixed_threshold = values[3]
        fixed_offset = Point(values[4], values[5])
    return CommentConfig(
        comment_from_action=comment_from_action,
        send_x_ratio=send_x_ratio,
        send_from_action=send_from_action,
        fixed_send_action_y_threshold=fixed_threshold,
        fixed_send_window_offset=fixed_offset,
    )


def save_comment_config(config: CommentConfig) -> None:
    values = [
        config.comment_from_action.x,
        config.comment_from_action.y,
        config.send_x_ratio,
    ]
    if config.send_from_action is not None:
        values.extend([config.send_from_action.x, config.send_from_action.y])
    if config.fixed_send_action_y_threshold is not None and config.fixed_send_window_offset is not None:
        values.extend(
            [
                config.fixed_send_action_y_threshold,
                config.fixed_send_window_offset.x,
                config.fixed_send_window_offset.y,
            ]
        )
    atomic_write_text(CONFIG_COMMENT, ",".join(str(v) for v in values))


class InputBackend:
    def __init__(self) -> None:
        self.pyautogui = require_module("pyautogui")
        self.pyperclip = require_module("pyperclip")
        self.pyautogui.PAUSE = 0.0
        self.native = os.name == "nt"
        self.user32 = ctypes.windll.user32 if self.native else None
        self._key_sequence_cache: dict[tuple[str, ...], tuple[int, Any]] = {}
        self._key_event_sequence_cache: dict[tuple[str, ...], tuple[INPUT, ...]] = {}
        self._vk_event_cache: dict[int, tuple[INPUT, INPUT]] = {}
        self._unicode_event_cache: dict[int, tuple[INPUT, INPUT]] = {}
        self._mouse_click_cache: dict[int, tuple[INPUT, ...]] = {}
        self._virtual_screen_metrics: Optional[tuple[int, int, int, int]] = None
        if self.user32 is not None:
            try:
                self.user32.SendInput.argtypes = (wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int)
                self.user32.SendInput.restype = wintypes.UINT
            except Exception:
                pass

    @staticmethod
    def _vk(key: str) -> int:
        mapping = {
            "tab": 0x09,
            "enter": 0x0D,
            "return": 0x0D,
            "esc": 0x1B,
            "escape": 0x1B,
            "ctrl": 0x11,
            "control": 0x11,
            "shift": 0x10,
            "alt": 0x12,
            " ": 0x20,
            "space": 0x20,
        }
        lowered = key.lower()
        if lowered in mapping:
            return mapping[lowered]
        if len(key) == 1:
            return ord(key.upper())
        raise EasyMoneyError(f"不支持的按键: {key}")

    def prepare_key_sequence(self, keys: Iterable[str]) -> None:
        key_tuple = tuple(keys)
        if not key_tuple or not (self.native and self.user32 is not None):
            return
        if key_tuple in self._key_sequence_cache:
            return
        events = self._key_input_events_for_sequence(key_tuple)
        array_type = INPUT * len(events)
        self._key_sequence_cache[key_tuple] = (len(events), array_type(*events))

    def press_sequence_atomic(self, keys: Iterable[str]) -> None:
        key_tuple = tuple(keys)
        if not key_tuple:
            return
        if self.native and self.user32 is not None:
            self.prepare_key_sequence(key_tuple)
            cached = self._key_sequence_cache.get(key_tuple)
            if cached is not None:
                count, event_array = cached
                sent = self.user32.SendInput(count, event_array, ctypes.sizeof(INPUT))
                if sent == count:
                    return
        self.press_sequence(key_tuple)

    def _key_input_events_for_vk(self, vk: int) -> tuple[INPUT, INPUT]:
        cached = self._vk_event_cache.get(vk)
        if cached is not None:
            return cached
        down = INPUT()
        down.type = INPUT_KEYBOARD
        down.union.ki = KEYBDINPUT(vk, 0, 0, 0, 0)
        up = INPUT()
        up.type = INPUT_KEYBOARD
        up.union.ki = KEYBDINPUT(vk, 0, KEYEVENTF_KEYUP, 0, 0)
        cached = (down, up)
        self._vk_event_cache[vk] = cached
        return cached

    def _key_input_events_for_sequence(self, keys: Iterable[str]) -> tuple[INPUT, ...]:
        key_tuple = tuple(keys)
        cached = self._key_event_sequence_cache.get(key_tuple)
        if cached is not None:
            return cached
        events: list[INPUT] = []
        for key in key_tuple:
            events.extend(self._key_input_events_for_vk(self._vk(key)))
        cached = tuple(events)
        self._key_event_sequence_cache[key_tuple] = cached
        return cached

    def _unicode_input_events_for_unit(self, unit: int) -> tuple[INPUT, INPUT]:
        cached = self._unicode_event_cache.get(unit)
        if cached is not None:
            return cached
        down = INPUT()
        down.type = INPUT_KEYBOARD
        down.union.ki = KEYBDINPUT(0, unit, KEYEVENTF_UNICODE, 0, 0)
        up = INPUT()
        up.type = INPUT_KEYBOARD
        up.union.ki = KEYBDINPUT(0, unit, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP, 0, 0)
        cached = (down, up)
        self._unicode_event_cache[unit] = cached
        return cached

    def _mouse_click_input_events(self, clicks: int = 1) -> tuple[INPUT, ...]:
        cached = self._mouse_click_cache.get(clicks)
        if cached is not None:
            return cached
        events: list[INPUT] = []
        for _ in range(clicks):
            down = INPUT()
            down.type = INPUT_MOUSE
            down.union.mi = MOUSEINPUT(0, 0, 0, MOUSEEVENTF_LEFTDOWN, 0, 0)
            up = INPUT()
            up.type = INPUT_MOUSE
            up.union.mi = MOUSEINPUT(0, 0, 0, MOUSEEVENTF_LEFTUP, 0, 0)
            events.extend([down, up])
        cached = tuple(events)
        self._mouse_click_cache[clicks] = cached
        return cached

    def _get_virtual_screen_metrics(self) -> tuple[int, int, int, int]:
        if self._virtual_screen_metrics is not None:
            return self._virtual_screen_metrics
        left = self.user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
        top = self.user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
        width = max(1, self.user32.GetSystemMetrics(SM_CXVIRTUALSCREEN))
        height = max(1, self.user32.GetSystemMetrics(SM_CYVIRTUALSCREEN))
        self._virtual_screen_metrics = (left, top, width, height)
        return self._virtual_screen_metrics

    def _absolute_mouse_coords(self, point: Point) -> tuple[int, int]:
        left, top, width, height = self._get_virtual_screen_metrics()
        x, y = point.rounded()
        dx = int(round((x - left) * 65535 / max(width - 1, 1)))
        dy = int(round((y - top) * 65535 / max(height - 1, 1)))
        return max(0, min(65535, dx)), max(0, min(65535, dy))

    def _send_input_events(self, events: list[INPUT]) -> bool:
        if not events or not (self.native and self.user32 is not None):
            return False
        array_type = INPUT * len(events)
        event_array = array_type(*events)
        sent = self.user32.SendInput(len(events), event_array, ctypes.sizeof(INPUT))
        return sent == len(events)

    def position(self) -> Point:
        if self.native and self.user32 is not None:
            point = wintypes.POINT()
            if self.user32.GetCursorPos(ctypes.byref(point)):
                return Point(float(point.x), float(point.y))
        x, y = self.pyautogui.position()
        return Point(float(x), float(y))

    def click(self, point: Point, clicks: int = 1, interval: float = 0.04) -> None:
        x, y = point.rounded()
        if self.native and self.user32 is not None:
            self.user32.SetCursorPos(x, y)
            if interval <= 0:
                if self._send_input_events(self._mouse_click_input_events(clicks)):
                    return
            for index in range(clicks):
                self.user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                self.user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                if interval > 0 and index < clicks - 1:
                    precise_delay(interval)
            return
        self.pyautogui.click(x=x, y=y, clicks=clicks, interval=interval, button="left")

    def move_to(self, point: Point) -> None:
        x, y = point.rounded()
        if self.native and self.user32 is not None:
            self.user32.SetCursorPos(x, y)
            return
        self.pyautogui.moveTo(x=x, y=y)

    def press(self, key: str) -> None:
        if self.native and self.user32 is not None:
            vk = self._vk(key)
            self.user32.keybd_event(vk, 0, 0, 0)
            self.user32.keybd_event(vk, 0, 0x0002, 0)
            return
        self.pyautogui.press(key)

    def press_sequence(self, keys: Iterable[str], gap: float = 0.0) -> None:
        key_list = list(keys)
        if not key_list:
            return
        if self.native and self.user32 is not None and gap <= 0:
            events = self._key_input_events_for_sequence(key_list)
            array_type = INPUT * len(events)
            event_array = array_type(*events)
            sent = self.user32.SendInput(len(events), event_array, ctypes.sizeof(INPUT))
            if sent == len(events):
                return
        for index, key in enumerate(key_list):
            self.press(key)
            if gap > 0 and index < len(key_list) - 1:
                precise_delay(gap)

    def hotkey(self, *keys: str) -> None:
        if self.native and self.user32 is not None:
            vks = [self._vk(key) for key in keys]
            for vk in vks:
                self.user32.keybd_event(vk, 0, 0, 0)
            for vk in reversed(vks):
                self.user32.keybd_event(vk, 0, 0x0002, 0)
            return
        self.pyautogui.hotkey(*keys)

    @staticmethod
    def can_type_directly(text: str) -> bool:
        return bool(text) and text.isascii() and all(ch.isalnum() or ch == " " for ch in text)

    @staticmethod
    def _utf16_units(text: str) -> list[int]:
        data = text.encode("utf-16-le", errors="surrogatepass")
        return [data[index] | (data[index + 1] << 8) for index in range(0, len(data), 2)]

    @staticmethod
    def can_type_unicode_directly(text: str) -> bool:
        if not text:
            return False
        units = len(text.encode("utf-16-le", errors="surrogatepass")) // 2
        if units <= 0 or units > DIRECT_TEXT_ENTRY_MAX_UTF16_UNITS:
            return False
        return all(ord(ch) >= 32 and ch not in "\r\n\t" for ch in text)

    def can_type_text_directly(self, text: str) -> bool:
        if self.native and self.user32 is not None:
            return self.can_type_unicode_directly(text)
        return self.can_type_directly(text)

    def type_unicode_text_directly(self, text: str) -> str:
        if not (self.native and self.user32 is not None):
            raise EasyMoneyError("当前平台不支持 Unicode 直接输入")
        if not self.can_type_unicode_directly(text):
            raise EasyMoneyError("当前文本不适合 Unicode 直接输入")
        units = self._utf16_units(text)
        chunk_size = max(1, DIRECT_TEXT_ENTRY_CHUNK_UTF16_UNITS)
        for offset in range(0, len(units), chunk_size):
            events: list[INPUT] = []
            for unit in units[offset : offset + chunk_size]:
                down = INPUT()
                down.type = INPUT_KEYBOARD
                down.union.ki = KEYBDINPUT(0, unit, KEYEVENTF_UNICODE, 0, 0)
                up = INPUT()
                up.type = INPUT_KEYBOARD
                up.union.ki = KEYBDINPUT(0, unit, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP, 0, 0)
                events.extend([down, up])
            array_type = INPUT * len(events)
            event_array = array_type(*events)
            sent = self.user32.SendInput(len(events), event_array, ctypes.sizeof(INPUT))
            if sent != len(events):
                raise EasyMoneyError("Unicode 直接输入失败")
        return "Unicode直接输入"

    def type_text_directly(self, text: str, interval: float = 0.0) -> str:
        if not self.can_type_text_directly(text):
            raise EasyMoneyError("当前文本不适合直接键盘输入")
        if self.native and self.user32 is not None:
            try:
                return self.type_unicode_text_directly(text)
            except EasyMoneyError:
                if not self.can_type_directly(text):
                    raise
                self.press_sequence(text, gap=interval)
                return "直接键盘输入"
        self.pyautogui.write(text, interval=interval)
        return "直接键盘输入"

    def paste_text(
        self,
        text: str,
        restore_clipboard: bool = True,
        before_paste_delay: float = 0.03,
        after_paste_delay: float = 0.06,
    ) -> str:
        old_text: Optional[str]
        try:
            old_text = self.pyperclip.paste()
        except Exception:
            old_text = None
        self.pyperclip.copy(text)
        if before_paste_delay > 0:
            time.sleep(before_paste_delay)
        self.hotkey("ctrl", "v")
        if after_paste_delay > 0:
            time.sleep(after_paste_delay)
        if restore_clipboard and old_text is not None:
            try:
                self.pyperclip.copy(old_text)
            except Exception:
                pass
        return "剪贴板粘贴"


def parse_key_sequence_text(raw: str, option_name: str = "key sequence") -> tuple[str, ...]:
    text = raw.strip()
    if not text:
        raise EasyMoneyError(f"{option_name} 不能为空")
    parts = [part.strip().lower() for part in re.split(r"[\s,]+", text) if part.strip()]
    if not parts:
        raise EasyMoneyError(f"{option_name} 不能为空")
    valid_names = {"tab", "enter", "return", "esc", "escape", "ctrl", "control", "shift", "alt", "space"}
    for part in parts:
        if part not in valid_names and len(part) != 1:
            raise EasyMoneyError(f"不支持的按键: {part}")
    return tuple(parts)


def format_key_sequence(keys: Iterable[str]) -> str:
    labels = {
        "tab": "Tab",
        "enter": "Enter",
        "return": "Enter",
        "esc": "Esc",
        "escape": "Esc",
        "ctrl": "Ctrl",
        "control": "Ctrl",
        "shift": "Shift",
        "alt": "Alt",
        "space": "Space",
        " ": "Space",
    }
    return "+".join(labels.get(key.lower(), key.upper() if len(key) == 1 else key) for key in keys)


class CaptureBackend:
    def __init__(self, backend: Optional[str] = None) -> None:
        self.Image = require_module("PIL.Image", "Pillow")
        requested_backend = (backend or os.environ.get("EASYMONEY_CAPTURE_BACKEND") or "auto").strip().lower()
        self.backend = requested_backend
        self._allow_mss_fallback = requested_backend == "auto"
        self._dxcam_mod = None
        self._dx_camera = None
        self._dx_stream_region: Optional[tuple[int, int, int, int]] = None
        self.mss_mod = None
        self._sct = None
        dxgi_ready = False
        if self.backend in {"auto", "dxgi", "dxcam"}:
            try:
                self._dxcam_mod = importlib.import_module("dxcam")
                output_idx = int(os.environ.get("EASYMONEY_DXGI_OUTPUT", "0"))
                self._dx_camera = self._dxcam_mod.create(output_idx=output_idx, output_color="RGB")
                if self._dx_camera is not None:
                    self.backend = "dxgi"
                    dxgi_ready = True
            except ImportError as exc:
                if self.backend in {"dxgi", "dxcam"}:
                    raise EasyMoneyError("缺少 DXGI 依赖 `dxcam`，请运行: python -m pip install -r requirements.txt") from exc
            except Exception as exc:
                if self.backend in {"dxgi", "dxcam"}:
                    raise EasyMoneyError(f"DXGI 捕获初始化失败: {exc}") from exc
        if dxgi_ready and not self._allow_mss_fallback:
            return
        self._init_mss(required=not dxgi_ready)
        if not dxgi_ready:
            self.backend = "mss"

    def _init_mss(self, required: bool = True) -> None:
        try:
            self.mss_mod = require_module("mss")
        except EasyMoneyError:
            if required:
                raise
            return
        if hasattr(self.mss_mod, "MSS"):
            self._sct = self.mss_mod.MSS()
        else:
            self._sct = self.mss_mod.mss()

    def _grab_mss(self, rect: Rect) -> CaptureFrame:
        if self._sct is None:
            raise CaptureUnavailable("截图后端未初始化")
        try:
            shot = self._sct.grab(rect.to_mss())
        except Exception as exc:
            raise CaptureUnavailable(f"MSS 截图失败: {rect.describe()} ({exc})") from exc
        return CaptureFrame(width=int(shot.width), height=int(shot.height), rgb=shot.rgb)

    def _grab_mss_fallback(self, rect: Rect, exc: Exception) -> CaptureFrame:
        if not self._allow_mss_fallback:
            raise CaptureUnavailable(
                f"DXGI 截图区域无效: {rect.describe()} ({exc})；"
                "窗口可能在副屏或跨屏区域，请移动到主屏，或设置 EASYMONEY_CAPTURE_BACKEND=mss"
            ) from exc
        if self._sct is None:
            self._init_mss(required=True)
        self._stop_dx_stream()
        return self._grab_mss(rect)

    def grab(self, rect: Rect) -> CaptureFrame:
        if rect.width <= 0 or rect.height <= 0:
            raise EasyMoneyError(f"截图区域无效: {rect.describe()}")
        if self.backend == "dxgi" and self._dx_camera is not None:
            region = (
                int(round(rect.left)),
                int(round(rect.top)),
                int(round(rect.right)),
                int(round(rect.bottom)),
            )
            if getattr(self._dx_camera, "is_capturing", False) and self._dx_stream_region != region:
                self._stop_dx_stream()
            try:
                frame = self._dx_camera.grab(region=region)
            except Exception as exc:
                return self._grab_mss_fallback(rect, exc)
            if frame is None:
                raise EasyMoneyError(f"DXGI 截图失败: {rect.describe()}")
            height, width = int(frame.shape[0]), int(frame.shape[1])
            return CaptureFrame(width=width, height=height, rgb=frame.tobytes())
        return self._grab_mss(rect)

    def grab_stream(self, rect: Rect) -> CaptureFrame:
        if self.backend != "dxgi" or self._dx_camera is None:
            return self.grab(rect)
        if rect.width <= 0 or rect.height <= 0:
            raise EasyMoneyError(f"截图区域无效: {rect.describe()}")
        region = (
            int(round(rect.left)),
            int(round(rect.top)),
            int(round(rect.right)),
            int(round(rect.bottom)),
        )
        if self._dx_stream_region != region or not getattr(self._dx_camera, "is_capturing", False):
            self._stop_dx_stream()
            fps = int(os.environ.get("EASYMONEY_DXGI_STREAM_FPS", "240"))
            try:
                self._dx_camera.start(region=region, target_fps=fps, video_mode=True)
            except Exception as exc:
                return self._grab_mss_fallback(rect, exc)
            self._dx_stream_region = region
        try:
            frame = self._dx_camera.get_latest_frame(copy=True)
        except Exception as exc:
            return self._grab_mss_fallback(rect, exc)
        if frame is None:
            raise EasyMoneyError(f"DXGI 流采帧失败: {rect.describe()}")
        height, width = int(frame.shape[0]), int(frame.shape[1])
        return CaptureFrame(width=width, height=height, rgb=frame.tobytes())

    def _stop_dx_stream(self) -> None:
        stop = getattr(self._dx_camera, "stop", None)
        if callable(stop):
            try:
                stop()
            except Exception:
                pass
        self._dx_stream_region = None

    def screenshot(self, rect: Rect):
        shot = self.grab(rect)
        return self.Image.frombytes("RGB", (shot.width, shot.height), shot.rgb)

    def save(self, image: Any, path: Path) -> Path:
        ensure_parent(path)
        image.save(path)
        return path

    def close(self) -> None:
        self._stop_dx_stream()
        close = getattr(self._sct, "close", None)
        if callable(close):
            close()

    def __del__(self) -> None:
        try:
            self.close()
        except Exception:
            pass


def quick_capture_fingerprint(capture: CaptureBackend, rect: Rect) -> Optional[int]:
    try:
        shot = capture.grab(rect)
    except Exception:
        return None
    width, height = int(shot.width), int(shot.height)
    if width <= 0 or height <= 0:
        return None
    data = shot.rgb
    samples: list[int] = []
    for row in range(8):
        y = min(height - 1, max(0, int((row + 0.5) * height / 8)))
        for col in range(8):
            x = min(width - 1, max(0, int((col + 0.5) * width / 8)))
            idx = (y * width + x) * 3
            r, g, b = data[idx], data[idx + 1], data[idx + 2]
            samples.append((int(r) * 30 + int(g) * 59 + int(b) * 11) // 100)
    avg = sum(samples) / max(1, len(samples))
    fingerprint = 0
    for value in samples:
        fingerprint = (fingerprint << 1) | (1 if value >= avg else 0)
    return fingerprint


def fingerprint_distance(lhs: int, rhs: int) -> int:
    return (lhs ^ rhs).bit_count()


def wait_for_region_refresh(
    capture: CaptureBackend,
    region: Rect,
    baseline_fingerprint: Optional[int],
    timeout_seconds: float = COMMENT_REFRESH_WAIT_SECONDS,
) -> bool:
    deadline = time.perf_counter() + max(0.001, timeout_seconds)
    next_check = time.perf_counter()
    while time.perf_counter() < deadline:
        now = time.perf_counter()
        if now >= next_check:
            current = quick_capture_fingerprint(capture, region)
            if current is not None:
                if baseline_fingerprint is None:
                    return True
                if fingerprint_distance(baseline_fingerprint, current) >= 10:
                    return True
            next_check = now + COMMENT_REFRESH_CAPTURE_INTERVAL_SECONDS
        else:
            time.sleep(COMMENT_REFRESH_IDLE_SECONDS)
    return False


def refresh_observation_region(window_rect: Rect) -> Rect:
    return Rect(
        window_rect.left,
        window_rect.top + min(60, max(0, window_rect.height * 0.08)),
        window_rect.left + max(1, window_rect.width / 7),
        window_rect.top + max(80, window_rect.height * 0.62),
    ).clamp_to(window_rect)


_UIA_CONTROL_TYPE_NAMES: Optional[dict[int, str]] = None


def uia_control_type_name(control_type: Any) -> str:
    try:
        control_type_id = int(control_type)
    except (TypeError, ValueError):
        return str(control_type or "")
    global _UIA_CONTROL_TYPE_NAMES
    if _UIA_CONTROL_TYPE_NAMES is None:
        names: dict[int, str] = {}
        try:
            comtypes_client = require_module("comtypes.client", "comtypes")
            comtypes_client.GetModule("UIAutomationCore.dll")
            from comtypes.gen import UIAutomationClient as UIA  # type: ignore

            for name in dir(UIA):
                if name.startswith("UIA_") and name.endswith("ControlTypeId"):
                    value = getattr(UIA, name)
                    if isinstance(value, int):
                        names[value] = name.removeprefix("UIA_").removesuffix("ControlTypeId")
        except Exception:
            names = {}
        _UIA_CONTROL_TYPE_NAMES = names
    return _UIA_CONTROL_TYPE_NAMES.get(control_type_id, str(control_type_id))


class WindowBackend:
    def __init__(self) -> None:
        self.pywinauto = require_module("pywinauto")
        try:
            timings = require_module("pywinauto.timings", "pywinauto")
            timings.Timings.window_find_timeout = float(os.environ.get("EASYMONEY_UIA_SEARCH_TIMEOUT", "1"))
        except Exception:
            pass
        self.desktop = self.pywinauto.Desktop(backend="uia")
        self._comtypes_client: Optional[Any] = None
        self._uia_module: Optional[Any] = None
        self._automation: Optional[Any] = None
        self._sns_list_cache: dict[Any, Any] = {}

    @staticmethod
    def _safe_text(control: Any) -> str:
        try:
            text = control.window_text()
            if text:
                return str(text)
        except Exception:
            pass
        for attr in ("Name", "CurrentName"):
            try:
                value = getattr(control, attr)
                return str(value() if callable(value) else value or "")
            except Exception:
                pass
        try:
            return control.element_info.name or ""
        except Exception:
            return ""

    @staticmethod
    def _control_type(control: Any) -> str:
        try:
            control_type = control.element_info.control_type
            if control_type:
                return str(control_type)
        except Exception:
            pass
        for attr in ("ControlTypeName", "CurrentControlType", "CurrentLocalizedControlType"):
            try:
                value = getattr(control, attr)
                value = value() if callable(value) else value
                if value:
                    if isinstance(value, int):
                        return uia_control_type_name(value)
                    text = str(value)
                    if text.endswith("Control") and len(text) > len("Control"):
                        return text[: -len("Control")]
                    return text
            except Exception:
                pass
        try:
            return str(control.element_info.control_type or "")
        except Exception:
            return ""

    @staticmethod
    def _automation_id(control: Any) -> str:
        try:
            automation_id = control.element_info.automation_id
            if automation_id:
                return str(automation_id)
        except Exception:
            pass
        for attr in ("AutomationId", "CurrentAutomationId"):
            try:
                value = getattr(control, attr)
                return str(value() if callable(value) else value or "")
            except Exception:
                pass
        try:
            return str(control.element_info.automation_id or "")
        except Exception:
            return ""

    @staticmethod
    def _class_name(control: Any) -> str:
        try:
            class_name = control.element_info.class_name
            if class_name:
                return str(class_name)
        except Exception:
            pass
        for attr in ("ClassName", "CurrentClassName"):
            try:
                value = getattr(control, attr)
                return str(value() if callable(value) else value or "")
            except Exception:
                pass
        try:
            return str(control.element_info.class_name or "")
        except Exception:
            return ""

    @staticmethod
    def rect(control: Any) -> Optional[Rect]:
        for attr in ("BoundingRectangle", "CurrentBoundingRectangle"):
            try:
                r = getattr(control, attr)
                r = r() if callable(r) else r
                left, top, right, bottom = float(r.left), float(r.top), float(r.right), float(r.bottom)
                if right <= left or bottom <= top:
                    return None
                return Rect(left, top, right, bottom)
            except Exception:
                pass
        try:
            r = control.rectangle()
            if r.right <= r.left or r.bottom <= r.top:
                return None
            return Rect(float(r.left), float(r.top), float(r.right), float(r.bottom))
        except Exception:
            return None

    def windows(self) -> list[Any]:
        try:
            return list(self.desktop.windows())
        except Exception:
            pass
        try:
            return list(self.desktop.children())
        except Exception as exc:
            raise EasyMoneyError(f"读取窗口列表失败: {exc}") from exc

    def _ensure_automation(self) -> tuple[Any, Any]:
        if self._automation is None or self._uia_module is None:
            self._comtypes_client = require_module("comtypes.client", "comtypes")
            self._comtypes_client.GetModule("UIAutomationCore.dll")
            from comtypes.gen import UIAutomationClient as UIA  # type: ignore

            self._uia_module = UIA
            self._automation = self._comtypes_client.CreateObject(UIA.CUIAutomation, interface=UIA.IUIAutomation)
        return self._automation, self._uia_module

    def _com_element(self, control: Any) -> Any:
        try:
            element = control.element_info.element
            if element:
                return element
        except Exception:
            pass
        for attr in ("NativeWindowHandle", "CurrentNativeWindowHandle"):
            try:
                handle = getattr(control, attr)
                handle = handle() if callable(handle) else handle
                if handle:
                    automation, _ = self._ensure_automation()
                    return automation.ElementFromHandle(ctypes.c_void_p(int(handle)))
            except Exception:
                pass
        for attr in ("Element", "_element"):
            try:
                element = getattr(control, attr)
                element = element() if callable(element) else element
                if element:
                    return element
            except Exception:
                pass
        return control

    def _walker_for_view(self, view: str) -> Any:
        automation, _ = self._ensure_automation()
        if view == "raw":
            return automation.RawViewWalker
        if view == "content":
            return automation.ContentViewWalker
        return automation.ControlViewWalker

    def find_sns_list_fast(self, root: Any) -> Optional[Any]:
        cache_key = self._control_identity(root)
        cached = self._sns_list_cache.get(cache_key)
        if cached is not None and is_sns_list_control(self, cached):
            return cached

        try:
            automation, UIA = self._ensure_automation()
            root_element = self._com_element(root)
            automation_id_condition = automation.CreatePropertyCondition(
                UIA.UIA_AutomationIdPropertyId,
                "sns_list",
            )
            control_type_condition = automation.CreatePropertyCondition(
                UIA.UIA_ControlTypePropertyId,
                UIA.UIA_ListControlTypeId,
            )
            condition = automation.CreateAndCondition(automation_id_condition, control_type_condition)
            element = root_element.FindFirst(UIA.TreeScope_Subtree, condition)
            if element is not None and is_sns_list_control(self, element):
                self._sns_list_cache[cache_key] = element
                return element
        except Exception:
            pass

        try:
            child_window = getattr(root, "child_window", None)
            if callable(child_window):
                spec = child_window(auto_id="sns_list", control_type="List")
                wrapper = spec.wrapper_object()
                if is_sns_list_control(self, wrapper):
                    self._sns_list_cache[cache_key] = wrapper
                    return wrapper
        except Exception:
            pass
        return None

    def moments_window(self) -> Any:
        candidates = []
        fallback_candidates = []
        for win in self.windows():
            title = self._safe_text(win).strip()
            class_name = self._class_name(win)
            if not title and not class_name:
                continue
            if "朋友圈" in title or class_name == "mmui::SNSWindow":
                return win
            if class_name == "mmui::MainWindow" or title in {"微信", "Weixin"}:
                candidates.append(win)
            elif "微信" in title or "WeChat" in title or "wechat" in title.lower():
                fallback_candidates.append(win)
        if candidates:
            print("警告: 未找到标题为“朋友圈”的窗口，使用疑似微信窗口；如定位异常，请先打开朋友圈窗口。")
            return candidates[0]
        if fallback_candidates:
            print("警告: 未找到标题为“朋友圈”的窗口，使用疑似微信窗口；如定位异常，请先打开朋友圈窗口。")
            return fallback_candidates[0]
        raise EasyMoneyError("未找到微信/朋友圈窗口，请先打开微信桌面版并进入朋友圈")

    def moments_window_rect(self) -> Rect:
        win = self.moments_window()
        rect = self.rect(win)
        if rect is None:
            raise WindowPositionUnavailable("无法读取朋友圈窗口位置")
        return rect

    def activate(self, control: Any) -> None:
        try:
            control.restore()
        except Exception:
            pass
        try:
            control.set_focus()
            return
        except Exception:
            pass
        try:
            control.SetActive()
            return
        except Exception:
            pass
        try:
            control.SetFocus()
            return
        except Exception:
            pass
        try:
            control.set_focus()
        except Exception:
            try:
                control.wrapper_object().set_focus()
            except Exception:
                pass

    def children(self, control: Any) -> list[Any]:
        try:
            return list(control.children())
        except Exception:
            pass
        try:
            return list(control.GetChildren())
        except Exception:
            pass
        try:
            walker = self._walker_for_view("raw")
            children: list[Any] = []
            child = walker.GetFirstChildElement(self._com_element(control))
            while child:
                children.append(child)
                child = walker.GetNextSiblingElement(child)
            return children
        except Exception:
            return []

    def listitem_children(self, control: Any, limit: Optional[int] = None) -> list[Any]:
        try:
            children = list(control.children(control_type="ListItem"))
            if children:
                return children[:limit] if limit is not None else children
        except Exception:
            pass
        return []

    @staticmethod
    def _control_identity(control: Any) -> Any:
        try:
            runtime_id = control.element_info.element.GetRuntimeId()
            values = tuple(int(item) for item in runtime_id)
            if values:
                return ("runtime", values)
        except Exception:
            pass
        try:
            runtime_id = control.GetRuntimeId()
            values = tuple(int(item) for item in runtime_id)
            if values:
                return ("runtime", values)
        except Exception:
            pass
        for attr in ("NativeWindowHandle", "CurrentNativeWindowHandle", "handle"):
            try:
                value = getattr(control, attr)
                value = value() if callable(value) else value
                if value:
                    return (attr, int(value))
            except Exception:
                pass
        return ("object", id(control))

    def iter_tree(self, control: Any, max_depth: int = 10) -> Iterable[tuple[Any, int]]:
        stack: list[tuple[Any, int]] = [(control, 0)]
        seen: set[Any] = set()
        while stack:
            item, depth = stack.pop()
            marker = self._control_identity(item)
            if marker in seen or depth > max_depth:
                continue
            seen.add(marker)
            yield item, depth
            kids = self.children(item)
            for child in reversed(kids):
                stack.append((child, depth + 1))

    def iter_tree_view(self, control: Any, max_depth: int = 10, view: str = "control") -> Iterable[tuple[Any, int]]:
        if view in {"", "default", "pywinauto"}:
            yield from self.iter_tree(control, max_depth=max_depth)
            return
        walker = self._walker_for_view(view)
        root = self._com_element(control)
        stack: list[tuple[Any, int]] = [(root, 0)]
        seen: set[Any] = set()
        while stack:
            item, depth = stack.pop()
            marker = self._control_identity(item)
            if marker in seen or depth > max_depth:
                continue
            seen.add(marker)
            yield item, depth
            if depth >= max_depth:
                continue
            children: list[Any] = []
            try:
                child = walker.GetFirstChildElement(item)
            except Exception:
                child = None
            while child:
                children.append(child)
                try:
                    child = walker.GetNextSiblingElement(child)
                except Exception:
                    break
            for child in reversed(children):
                stack.append((child, depth + 1))

    def dump_tree(self, control: Any, max_depth: int = 10, buttons_only: bool = False, view: str = "default") -> tuple[int, bool, bool]:
        count = 0
        found_sns_list = False
        found_list_item = False
        for node, depth in self.iter_tree_view(control, max_depth=max_depth, view=view):
            control_type = self._control_type(node)
            automation_id = self._automation_id(node)
            if automation_id == "sns_list" and control_type.lower() == "list":
                found_sns_list = True
            if control_type.lower() == "listitem":
                found_list_item = True
            if buttons_only and control_type.lower() != "button":
                continue
            print_uia_node(self, node, depth)
            count += 1
        if buttons_only:
            print(f"\n按钮数量: {count}")
        return count, found_sns_list, found_list_item

    def find_buttons(self, control: Any, max_depth: int = 12) -> list[Any]:
        return [node for node, _ in self.iter_tree(control, max_depth=max_depth) if self._control_type(node).lower() == "button"]

    def click_control(
        self,
        control: Any,
        input_backend: Optional[InputBackend] = None,
        prefer_coordinate: bool = False,
    ) -> bool:
        rect = self.rect(control)
        if prefer_coordinate and rect is not None:
            if input_backend is None:
                input_backend = InputBackend()
            input_backend.click(rect.center)
            return True
        try:
            pattern = control.GetInvokePattern()
            if pattern:
                pattern.Invoke()
                return True
        except Exception:
            pass
        try:
            control.invoke()
            return True
        except Exception:
            pass
        try:
            control.Click()
            return True
        except Exception:
            pass
        if rect is None:
            return False
        if input_backend is None:
            input_backend = InputBackend()
        input_backend.click(rect.center)
        return True


def countdown(seconds: int = 3) -> None:
    for i in range(seconds, 0, -1):
        print(f"  {i}...")
        time.sleep(1)


def precise_delay(seconds: float) -> None:
    if seconds <= 0:
        return
    if seconds >= 0.01:
        time.sleep(seconds)
        return
    deadline = time.perf_counter() + seconds
    while time.perf_counter() < deadline:
        pass


def current_timestamp_ms() -> str:
    now = time.time()
    millis = int((now % 1) * 1000)
    return f"{time.strftime('%H:%M:%S', time.localtime(now))}.{millis:03d}"


BLOCKED_TEXT_SNIPPETS = {
    "评论",
    "赞",
    "分钟前",
    "小时前",
    "刚刚",
    "昨天",
    "前天",
    "微信",
    "删除",
    "详情",
    "收起",
}


def clean_post_text(text: str) -> str:
    lines = []
    for raw in text.replace("\r", "\n").split("\n"):
        line = raw.strip()
        if not line:
            continue
        if any(snippet in line for snippet in BLOCKED_TEXT_SNIPPETS) and len(line) < 12:
            continue
        lines.append(line)
    return "\n".join(lines).strip()


def uia_list_item_prefix(text: str) -> str:
    for raw in text.replace("\r", "\n").split("\n"):
        line = raw.strip()
        if line:
            return line
    return ""


def uia_list_item_matches_user_id(text: str, expected_user_id: str) -> bool:
    expected = expected_user_id.strip()
    if not expected:
        return False
    return text.lstrip().startswith(expected)


def short_log_text(text: str, limit: int = 48) -> str:
    clean = text.replace("\r", " ").replace("\n", " ").strip()
    if len(clean) <= limit:
        return clean
    return clean[:limit] + "..."


def is_sns_list_control(window_backend: WindowBackend, node: Any) -> bool:
    return window_backend._automation_id(node) == "sns_list" and window_backend._control_type(node).lower() == "list"


def find_sns_list_control(window_backend: WindowBackend, root: Any, max_depth: int = 8) -> Optional[Any]:
    fast = window_backend.find_sns_list_fast(root)
    if fast is not None:
        return fast

    for node, _ in window_backend.iter_tree(root, max_depth=max_depth):
        if is_sns_list_control(window_backend, node):
            window_backend._sns_list_cache[window_backend._control_identity(root)] = node
            return node

    try:
        for node, _ in window_backend.iter_tree_view(root, max_depth=max_depth, view="raw"):
            if is_sns_list_control(window_backend, node):
                window_backend._sns_list_cache[window_backend._control_identity(root)] = node
                return node
    except Exception:
        pass

    try:
        child_window = getattr(root, "child_window", None)
        if callable(child_window):
            spec = child_window(auto_id="sns_list", control_type="List")
            wrapper = spec.wrapper_object()
            if is_sns_list_control(window_backend, wrapper):
                window_backend._sns_list_cache[window_backend._control_identity(root)] = wrapper
                return wrapper
    except Exception:
        pass

    try:
        descendants = root.descendants(auto_id="sns_list", control_type="List")
    except Exception:
        descendants = []
    for node in descendants:
        if is_sns_list_control(window_backend, node):
            window_backend._sns_list_cache[window_backend._control_identity(root)] = node
            return node
    return None


def find_named_list_control(window_backend: WindowBackend, root: Any, name: str, max_depth: int = 8, view: str = "raw") -> Optional[Any]:
    expected = name.strip()
    for node, _ in window_backend.iter_tree_view(root, max_depth=max_depth, view=view):
        if window_backend._control_type(node).lower() != "list":
            continue
        if window_backend._safe_text(node).strip() == expected:
            return node
    return None


def print_uia_node(window_backend: WindowBackend, node: Any, depth: int = 0) -> None:
    control_type = window_backend._control_type(node)
    automation_id = window_backend._automation_id(node)
    rect = window_backend.rect(node)
    name = window_backend._safe_text(node)
    indent = "  " * depth
    parts = [f"{indent}[{control_type or '?'}]"]
    if name:
        parts.append(f'name="{name}"')
    if automation_id:
        parts.append(f'id="{automation_id}"')
    if rect:
        parts.append(f"rect={rect.describe()}")
    print(" ".join(parts))


def dump_named_list_contents(
    window_backend: WindowBackend,
    root: Any,
    name: str,
    max_depth: int = 7,
    view: str = "raw",
    quiet: bool = False,
    item_limit: int = 1,
    item_index: int = 2,
) -> tuple[bool, bool, int]:
    expected = name.strip()
    if expected == "朋友圈":
        node = find_sns_list_control(window_backend, root, max_depth=max_depth)
    else:
        node = find_named_list_control(window_backend, root, expected, max_depth=max_depth, view=view)
    if node is None:
        return False, False, 0
    if window_backend._control_type(node).lower() != "list":
        return False, False, 0
    if expected and window_backend._safe_text(node).strip() != expected:
        return False, False, 0

    if not quiet:
        print_uia_node(window_backend, node, depth=0)
    found_list_item = False
    seen_list_items = 0
    item_count = 0
    read_limit = None if item_limit <= 0 else item_index + item_limit - 1
    for child in find_list_items_under_control(window_backend, node, limit=read_limit):
        seen_list_items += 1
        if seen_list_items < item_index:
            continue
        item_count += 1
        found_list_item = True
        if not quiet:
            print_uia_node(window_backend, child, depth=1)
        if item_limit > 0 and item_count >= item_limit:
            break
    return True, found_list_item, item_count


def find_list_items_under_control(window_backend: WindowBackend, root: Any, max_depth: int = 3, limit: Optional[int] = None) -> list[Any]:
    items: list[Any] = []
    seen: set[Any] = set()

    def add_if_list_item(node: Any) -> bool:
        marker = window_backend._control_identity(node)
        if marker in seen:
            return False
        if window_backend._control_type(node).lower() != "listitem":
            return False
        seen.add(marker)
        items.append(node)
        return limit is not None and len(items) >= limit

    for node in window_backend.listitem_children(root, limit=limit):
        if add_if_list_item(node):
            return items

    if items and (limit is None or len(items) >= limit):
        return items

    for node in window_backend.children(root):
        if add_if_list_item(node):
            return items

    if items and (limit is None or len(items) >= limit):
        return items

    for node, depth in window_backend.iter_tree(root, max_depth=max_depth):
        if depth == 0:
            continue
        if add_if_list_item(node):
            return items

    if items and (limit is None or len(items) >= limit):
        return items

    try:
        descendants = root.descendants(control_type="ListItem")
    except Exception:
        descendants = []
    for node in descendants:
        if add_if_list_item(node):
            return items
    if items:
        return items

    try:
        descendants = root.descendants()
    except Exception:
        descendants = []
    for node in descendants:
        if add_if_list_item(node):
            return items
    return items


def find_uia_list_items(window_backend: WindowBackend, root: Any, max_depth: int = 12, limit: Optional[int] = None) -> list[Any]:
    sns_list = find_sns_list_control(window_backend, root, max_depth=max_depth)
    if sns_list is None:
        return []
    return find_list_items_under_control(window_backend, sns_list, limit=limit)


def resolve_second_uia_list_item_post(
    window_backend: WindowBackend,
    win: Any,
    expected_user_id: str,
    item_index: int = 1,
    settle_ms: int = 220,
    poll_ms: int = 40,
    include_text: bool = True,
) -> UIAListItemResolution:
    started = time.perf_counter()
    deadline = started + max(0, settle_ms) / 1000
    items: list[Any] = []
    while True:
        sns_list = find_sns_list_control(window_backend, win)
        if sns_list is None:
            items = []
        else:
            items = find_list_items_under_control(window_backend, sns_list, limit=item_index + 1)
        if len(items) > item_index:
            break
        if time.perf_counter() >= deadline:
            if sns_list is None:
                raise UIAListItemUnavailable("UIA 未暴露 sns_list，无法读取朋友圈列表")
            raise UIAListItemUnavailable(
                f"UIA 中 ListItem 数量不足: 当前 {len(items)} 个，无法读取第 {item_index + 1} 个"
            )
        time.sleep(max(0.01, poll_ms / 1000))
    if len(items) <= item_index:
        raise UIAListItemUnavailable(f"UIA 中 ListItem 数量不足: 当前 {len(items)} 个，无法读取第 {item_index + 1} 个")

    item = items[item_index]
    raw_text = window_backend._safe_text(item).strip()
    if not raw_text:
        raise EasyMoneyError(f"第 {item_index + 1} 个 ListItem 的 name 为空")
    prefix = uia_list_item_prefix(raw_text)
    if not uia_list_item_matches_user_id(raw_text, expected_user_id):
        preview = raw_text.replace("\r", "\n").replace("\n", " / ")[:120]
        raise EasyMoneyError(
            f"第 {item_index + 1} 个 ListItem 不匹配 --user={expected_user_id}；"
            f"检测到开头={prefix or '(空)'}；内容预览={preview}"
        )

    rect = window_backend.rect(item)
    if rect is None:
        raise EasyMoneyError(f"第 {item_index + 1} 个 ListItem 无有效坐标")
    elapsed_ms = int((time.perf_counter() - started) * 1000)
    action_point = Point(rect.right - 40, rect.bottom - 10)
    return UIAListItemResolution(
        item_index=item_index,
        body_frame=rect,
        action_point=action_point,
        text=clean_post_text(raw_text) if include_text else "",
        expected_user_id=expected_user_id.strip(),
        detected_prefix=prefix,
        elapsed_ms=elapsed_ms,
    )


def resolve_send_point(action_point: Point, window_rect: Rect, config: CommentConfig) -> tuple[Point, str]:
    if (
        config.fixed_send_action_y_threshold is not None
        and config.fixed_send_window_offset is not None
        and action_point.y - window_rect.top >= config.fixed_send_action_y_threshold
    ):
        point = Point(window_rect.left + config.fixed_send_window_offset.x, window_rect.top + config.fixed_send_window_offset.y)
        return point, "低位固定"
    if config.send_from_action is not None:
        return Point(action_point.x + config.send_from_action.x, action_point.y + config.send_from_action.y), "手动偏移"
    return Point(window_rect.left + window_rect.width * config.send_x_ratio, action_point.y + config.comment_from_action.y + 48), "比例兜底"


def load_easy_money_dotenv() -> dict[str, str]:
    values: dict[str, str] = {}
    for path in [Path.cwd() / ".easyMoney.env", HOME / ".easyMoney.env"]:
        if not path.exists():
            continue
        for raw in path.read_text(encoding="utf-8").splitlines():
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            if line.startswith("export "):
                line = line[len("export ") :].strip()
            if "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip()
            if (value.startswith('"') and value.endswith('"')) or (value.startswith("'") and value.endswith("'")):
                value = value[1:-1]
            if key:
                values[key] = value
    return values


_DOTENV_CACHE: Optional[dict[str, str]] = None


def first_non_empty_env(keys: Iterable[str]) -> Optional[str]:
    global _DOTENV_CACHE
    for key in keys:
        value = os.environ.get(key, "").strip()
        if value:
            return value
    if _DOTENV_CACHE is None:
        _DOTENV_CACHE = load_easy_money_dotenv()
    for key in keys:
        value = (_DOTENV_CACHE.get(key) or "").strip()
        if value:
            return value
    return None


def load_local_llm_config() -> Optional[LocalLLMConfig]:
    raw_endpoint = first_non_empty_env(
        [
            "EASYMONEY_LLM_ENDPOINT",
            "WECHAT_LLM_ENDPOINT",
            "OMLX_ENDPOINT",
            "OLMX_ENDPOINT",
            "OLLAMA_HOST",
            "DOUBAO_ENDPOINT",
            "ARK_ENDPOINT",
            "VOLCENGINE_LLM_ENDPOINT",
        ]
    )
    inferred = None
    if raw_endpoint:
        if ":8000" in raw_endpoint or "/admin/chat" in raw_endpoint or "/v1/" in raw_endpoint:
            inferred = "openai"
        elif "volces.com" in raw_endpoint or "volcengine.com" in raw_endpoint:
            inferred = "doubao"
    provider = (
        first_non_empty_env(
            [
                "EASYMONEY_LLM_PROVIDER",
                "WECHAT_LLM_PROVIDER",
                "OMLX_PROVIDER",
                "OLMX_PROVIDER",
                "OLLAMA_PROVIDER",
                "DOUBAO_PROVIDER",
                "ARK_PROVIDER",
                "VOLCENGINE_LLM_PROVIDER",
            ]
        )
        or inferred
        or "ollama"
    ).lower()
    if provider in {"openai-compatible", "openai_compatible"}:
        provider = "openai"
    if provider not in {"ollama", "openai", "doubao"}:
        provider = "ollama"

    default_endpoint = {
        "ollama": "http://127.0.0.1:11434/api/chat",
        "openai": "http://127.0.0.1:11434/v1/chat/completions",
        "doubao": "https://ark.cn-beijing.volces.com/api/v3/responses",
    }[provider]
    endpoint = raw_endpoint or default_endpoint
    if endpoint.endswith("/admin/chat"):
        endpoint = endpoint[: -len("/admin/chat")] + "/v1/chat/completions"
    elif provider == "ollama" and endpoint.rstrip("/") in {"http://127.0.0.1:11434", "http://localhost:11434"}:
        endpoint = endpoint.rstrip("/") + "/api/chat"
    elif provider == "openai" and (
        endpoint.rstrip("/") in {"http://127.0.0.1:11434", "http://localhost:11434", "http://127.0.0.1:8000", "http://localhost:8000"}
        or endpoint.endswith("/v1")
    ):
        endpoint = endpoint.rstrip("/")
        endpoint = endpoint + "/chat/completions" if endpoint.endswith("/v1") else endpoint + "/v1/chat/completions"
    elif provider == "doubao" and (
        endpoint.rstrip("/") in {"https://ark.cn-beijing.volces.com", "https://ark.cn-beijing.volces.com/api/v3"}
        or endpoint.endswith("/api/v3")
    ):
        endpoint = endpoint.rstrip("/")
        endpoint = endpoint + "/responses" if endpoint.endswith("/api/v3") else endpoint + "/api/v3/responses"

    model = first_non_empty_env(
        [
            "EASYMONEY_LLM_MODEL",
            "WECHAT_LLM_MODEL",
            "OMLX_MODEL",
            "OLMX_MODEL",
            "OLLAMA_MODEL",
            "DOUBAO_MODEL",
            "ARK_MODEL",
            "VOLCENGINE_LLM_MODEL",
        ]
    )
    if not model and provider == "doubao":
        model = "doubao-seed-2-0-mini-260215"
    if not model:
        return None
    timeout = float(first_non_empty_env(["EASYMONEY_LLM_TIMEOUT", "WECHAT_LLM_TIMEOUT"]) or "18")
    api_key = first_non_empty_env(
        [
            "EASYMONEY_LLM_API_KEY",
            "WECHAT_LLM_API_KEY",
            "OMLX_API_KEY",
            "OLMX_API_KEY",
            "OPENAI_API_KEY",
            "DOUBAO_API_KEY",
            "ARK_API_KEY",
            "VOLCENGINE_API_KEY",
        ]
    )
    return LocalLLMConfig(provider=provider, endpoint=endpoint, model=model, api_key=api_key, timeout_seconds=max(5.0, timeout))


def generic_llm_system_prompt() -> str:
    return "你是一个可靠的中文助手。回答要直接、简洁。"


def doubao_question_solve_system_prompt() -> str:
    return "你在帮助用户根据朋友圈正文回答剧本杀/活动相关问题。只输出最终答案；不知道就输出“不知道”。"


def build_generic_llm_user_prompt(prompt: str, context: str = "") -> str:
    if context.strip():
        return f"上下文：\n{context.strip()}\n\n问题：\n{prompt.strip()}"
    return prompt.strip()


def build_doubao_question_prompt(post_text: str) -> str:
    return f"请根据下面朋友圈正文回答问题，只输出答案：\n\n{post_text.strip()}"


def parse_openai_compatible_response(root: dict[str, Any]) -> Optional[str]:
    try:
        content = root["choices"][0]["message"]["content"]
        if isinstance(content, str):
            return content
        if isinstance(content, list):
            parts = []
            for item in content:
                if isinstance(item, dict):
                    parts.append(str(item.get("text") or item.get("content") or ""))
            return "".join(parts).strip()
    except Exception:
        return None
    return None


def parse_responses_api_response(root: dict[str, Any]) -> Optional[str]:
    text = root.get("output_text")
    if isinstance(text, str) and text.strip():
        return text
    parts: list[str] = []
    for output in root.get("output") or []:
        if not isinstance(output, dict):
            continue
        for item in output.get("content") or []:
            if not isinstance(item, dict):
                continue
            value = item.get("text") or item.get("value")
            if isinstance(value, str):
                parts.append(value)
    return "".join(parts).strip() or None


def parse_ollama_response(root: dict[str, Any]) -> Optional[str]:
    message = root.get("message")
    if isinstance(message, dict) and isinstance(message.get("content"), str):
        return message["content"]
    if isinstance(root.get("response"), str):
        return root["response"]
    return None


def clean_llm_answer(text: str) -> str:
    answer = text.strip()
    answer = re.sub(r"^\s*(答案|回答)\s*[:：]\s*", "", answer)
    return answer.strip().strip('"').strip("'")


def request_llm_answer(
    config: LocalLLMConfig,
    system_prompt: str,
    user_prompt: str,
    image_data_urls: Optional[list[str]] = None,
) -> Optional[str]:
    requests = require_module("requests")
    images = [url for url in (image_data_urls or []) if url.strip()]
    headers = {"Content-Type": "application/json"}
    if config.api_key:
        headers["Authorization"] = f"Bearer {config.api_key}"

    if config.provider == "ollama":
        body: dict[str, Any] = {
            "model": config.model,
            "stream": False,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
        }
    elif config.provider == "openai":
        content: Any = user_prompt
        if images:
            content = [{"type": "text", "text": user_prompt}]
            content.extend({"type": "image_url", "image_url": {"url": url}} for url in images)
        body = {
            "model": config.model,
            "stream": False,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": content},
            ],
            "temperature": 0.0,
            "max_tokens": 256,
        }
    else:
        if "/responses" in config.endpoint:
            user_content: list[dict[str, Any]] = [{"type": "input_text", "text": user_prompt}]
            user_content[0:0] = [{"type": "input_image", "image_url": url, "detail": "auto"} for url in images]
            body = {
                "model": config.model,
                "stream": False,
                "max_output_tokens": 256,
                "thinking": {"type": "disabled"},
                "input": [
                    {"type": "message", "role": "system", "content": system_prompt},
                    {"type": "message", "role": "user", "content": user_content},
                ],
            }
        else:
            content = [{"type": "text", "text": user_prompt}]
            content.extend({"type": "image_url", "image_url": {"url": url}} for url in images)
            body = {
                "model": config.model,
                "stream": False,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": content},
                ],
                "temperature": 0.0,
                "max_tokens": 256,
            }

    try:
        response = requests.post(config.endpoint, headers=headers, json=body, timeout=config.timeout_seconds)
    except Exception as exc:
        print(f"  LLM 请求失败: {exc}")
        return None
    if not 200 <= response.status_code <= 299:
        print(f"  LLM 请求失败: HTTP {response.status_code}")
        print(f"  响应: {response.text[:1000]}")
        return None
    try:
        root = response.json()
    except Exception:
        print(f"  LLM 响应不是 JSON: {response.text[:1000]}")
        return None
    if config.provider == "ollama":
        answer = parse_ollama_response(root)
    elif "/responses" in config.endpoint:
        answer = parse_responses_api_response(root)
    else:
        answer = parse_openai_compatible_response(root)
    return clean_llm_answer(answer) if answer else None


def ask_local_llm(prompt: str, context: str = "") -> Optional[str]:
    config = load_local_llm_config()
    if not config:
        print("LLM 配置缺失：请设置 EASYMONEY_LLM_MODEL 或 DOUBAO/ARK_MODEL")
        return None
    return request_llm_answer(config, generic_llm_system_prompt(), build_generic_llm_user_prompt(prompt, context))


def ask_doubao_to_solve_post(post_text: str, image_data_urls: Optional[list[str]] = None) -> Optional[SolvedQuestion]:
    config = load_local_llm_config()
    if not config:
        print("豆包/LLM 配置缺失：请检查 .easyMoney.env")
        return None
    answer = request_llm_answer(config, doubao_question_solve_system_prompt(), build_doubao_question_prompt(post_text), image_data_urls=image_data_urls)
    if not answer:
        return None
    return SolvedQuestion(answer=answer, evidence="LLM", confidence=0.62, source=config.provider)


def image_to_data_url(image: Any, max_side: int = 1280, quality: int = 78) -> str:
    from io import BytesIO

    img = image.convert("RGB")
    if max(img.size) > max_side:
        scale = max_side / max(img.size)
        img = img.resize((max(1, int(img.width * scale)), max(1, int(img.height * scale))))
    buf = BytesIO()
    img.save(buf, format="JPEG", quality=quality, optimize=True)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode("ascii")


def capture_post_image(post: MomentPostResolution, window_rect: Rect) -> Any:
    capture = CaptureBackend()
    region = post.body_frame.expanded(12, 24).clamp_to(window_rect)
    try:
        return capture.screenshot(region)
    finally:
        capture.close()


def capture_yolo_image_data_urls(post: MomentPostResolution, window_rect: Rect) -> list[str]:
    model_path = expand_path(first_non_empty_env(["EASYMONEY_YOLO_MODEL", "EASYMONEY_DOUBAO_IMAGE_MODEL", "DOUBAO_IMAGE_REGION_MODEL"]))
    if model_path is None or not model_path.exists():
        raise EasyMoneyError("--LLM --vision 需要配置 EASYMONEY_YOLO_MODEL 指向 .pt 模型")
    ultralytics = require_module("ultralytics")
    cv2 = require_module("cv2", "opencv-python")
    np = require_module("numpy")
    image = capture_post_image(post, window_rect)
    arr = cv2.cvtColor(np.array(image.convert("RGB")), cv2.COLOR_RGB2BGR)
    conf = float(first_non_empty_env(["EASYMONEY_YOLO_CONF", "YOLO_CONF"]) or "0.25")
    model = ultralytics.YOLO(str(model_path))
    result = model.predict(arr, imgsz=640, conf=conf, verbose=False)[0]
    boxes = getattr(result, "boxes", None)
    if boxes is None or len(boxes) == 0:
        raise EasyMoneyError("YOLO 未检测到图片区域")
    best = max(boxes, key=lambda box: float(box.conf[0]) if box.conf is not None else 0.0)
    x1, y1, x2, y2 = [int(v) for v in best.xyxy[0].tolist()]
    x1, y1 = max(0, x1), max(0, y1)
    x2, y2 = min(image.width, x2), min(image.height, y2)
    cropped = image.crop((x1, y1, x2, y2)) if x2 > x1 and y2 > y1 else image
    return [image_to_data_url(cropped)]


def print_usage() -> None:
    print(
        f"""
{APP_NAME} {APP_VERSION}

用法:
  python easy_money_win.py run [--interval N] [--pos x,y] [--index N] [--title 文本] [--id id]
  python easy_money_win.py locate [--test-click]
  python easy_money_win.py capture-info [--backend auto|dxgi|mss]
  python easy_money_win.py uia-dump [--max-depth N] [--item-index N] [--item-limit N] [--full] [--watch] [--count N] [--timing-only]
  python easy_money_win.py ax-dump [--max-depth N] [--item-index N] [--item-limit N] [--full] [--watch] [--count N] [--timing-only]
  python easy_money_win.py comment-locate
  python easy_money_win.py comment-fixed-send-locate
  python easy_money_win.py post-image-locate
  python easy_money_win.py post-image-x-locate
  python easy_money_win.py comment [--text 文本] [--solve-question|--doubao|--LLM [--vision]] --user <用户名前缀> [--debug]
  python easy_money_win.py llm ask "<问题>" [上下文]
  python easy_money_win.py doubao ask "<朋友圈正文>"
"""
    )


def parse_option_value(args: list[str], index: int, name: str) -> tuple[str, int]:
    if index + 1 >= len(args) or args[index + 1].startswith("--"):
        raise EasyMoneyError(f"{name} 需要提供值")
    return args[index + 1], index + 1


def cmd_capture_info(args: list[str]) -> int:
    backend_name: Optional[str] = None
    i = 0
    while i < len(args):
        if args[i] == "--backend":
            backend_name, i = parse_option_value(args, i, "--backend")
        i += 1
    backend = WindowBackend()
    win_rect = backend.moments_window_rect()
    region = refresh_observation_region(win_rect)
    capture = CaptureBackend(backend_name)
    started = time.perf_counter()
    frame = capture.grab(region)
    elapsed_ms = int((time.perf_counter() - started) * 1000)
    print(f"截图后端: {capture.backend}")
    print(f"测试区域: {region.describe()}")
    print(f"帧尺寸: {frame.width}x{frame.height}")
    print(f"截图耗时: {elapsed_ms}ms")
    capture.close()
    return 0


def cmd_uia_dump(args: list[str]) -> int:
    max_depth = 7
    buttons_only = False
    full = False
    settle_ms = 1500
    view = "raw"
    list_name = "朋友圈"
    watch = False
    interval_seconds = 0.2
    count = 0
    timing_only = False
    item_limit = 1
    item_index = 2
    i = 0
    while i < len(args):
        if args[i] == "--max-depth":
            value, i = parse_option_value(args, i, "--max-depth")
            max_depth = max(1, min(int(value), 30))
        elif args[i] == "--buttons-only":
            buttons_only = True
            full = True
        elif args[i] == "--full":
            full = True
        elif args[i] == "--settle-ms":
            value, i = parse_option_value(args, i, "--settle-ms")
            settle_ms = max(0, int(value))
        elif args[i] == "--view":
            value, i = parse_option_value(args, i, "--view")
            view = value.strip().lower()
            if view not in {"raw", "default", "control", "content"}:
                raise EasyMoneyError("不支持的 --view，可用: raw|default|control|content")
        elif args[i] == "--list-name":
            list_name, i = parse_option_value(args, i, "--list-name")
        elif args[i] == "--item-limit":
            value, i = parse_option_value(args, i, "--item-limit")
            item_limit = max(0, int(value))
        elif args[i] == "--item-index":
            value, i = parse_option_value(args, i, "--item-index")
            item_index = max(1, int(value))
        elif args[i] in {"--watch", "--loop"}:
            watch = True
        elif args[i] == "--interval":
            value, i = parse_option_value(args, i, "--interval")
            interval_seconds = max(0.0, float(value))
        elif args[i] in {"--count", "--iterations"}:
            value, i = parse_option_value(args, i, args[i])
            count = max(0, int(value))
        elif args[i] == "--timing-only":
            timing_only = True
        i += 1
    backend = WindowBackend()
    win = backend.moments_window()

    def dump_once(current_win: Any, quiet: bool) -> tuple[Any, bool, bool, int, int]:
        started = time.perf_counter()
        found_sns = False
        found_item = False
        item_count = 0
        if full:
            _, found_sns, found_item = backend.dump_tree(current_win, max_depth=max_depth, buttons_only=buttons_only, view=view)
        else:
            deadline = time.perf_counter() + settle_ms / 1000
            while time.perf_counter() < deadline:
                found_sns, found_item, item_count = dump_named_list_contents(
                    backend,
                    current_win,
                    list_name,
                    max_depth=max_depth,
                    view=view,
                    quiet=quiet,
                    item_limit=item_limit,
                    item_index=item_index,
                )
                if found_sns:
                    break
                time.sleep(0.03 if watch else 0.12)
                try:
                    current_win = backend.moments_window()
                except Exception:
                    pass
            if not found_sns and not quiet:
                print(f'未找到 name="{list_name}" 的 List；可加 --full 查看完整 UIA 树。')
        elapsed_ms = int((time.perf_counter() - started) * 1000)
        return current_win, found_sns, found_item, item_count, elapsed_ms

    if watch:
        iteration = 0
        try:
            while count <= 0 or iteration < count:
                iteration += 1
                if not timing_only:
                    print(f"\n--- uia-dump watch #{iteration} ---")
                win, found_sns_list, found_list_item, item_count, elapsed_ms = dump_once(win, quiet=timing_only)
                print(
                    f"[watch] #{iteration} total={elapsed_ms}ms "
                    f"found_list={str(found_sns_list).lower()} items={item_count}"
                )
                if interval_seconds > 0 and (count <= 0 or iteration < count):
                    time.sleep(interval_seconds)
        except KeyboardInterrupt:
            print("\n已停止 uia-dump watch")
        return 0

    win, found_sns_list, found_list_item, _, _ = dump_once(win, quiet=False)
    if not buttons_only and (not found_sns_list or not found_list_item):
        print(f'\n提示: 当前未读取到 name="{list_name}" 的 ListItem。')
        print("      可先用 --full 查看完整树，或稍等、重新激活朋友圈窗口后再 dump。")
    return 0


def click_button_by_title(
    backend: WindowBackend,
    root: Any,
    input_backend: InputBackend,
    title: str,
    automation_id: Optional[str] = None,
) -> Optional[Point]:
    for btn in backend.find_buttons(root):
        name = backend._safe_text(btn)
        button_id = backend._automation_id(btn)
        if title and title not in name:
            continue
        if automation_id and automation_id != button_id:
            continue
        rect = backend.rect(btn)
        if not backend.click_control(btn, input_backend, prefer_coordinate=True):
            return None
        return rect.center if rect is not None else None
    return None


def refresh_point_from_saved_offset(backend: WindowBackend) -> Optional[Point]:
    offset = load_point(CONFIG_REFRESH)
    if offset is None:
        return None
    win_rect = backend.moments_window_rect()
    return Point(win_rect.left + offset.x, win_rect.top + offset.y)


def cmd_locate(args: list[str]) -> int:
    test_click = False
    i = 0
    while i < len(args):
        if args[i] in {"--mouse", "--manual"}:
            pass
        elif args[i] == "--test-click":
            test_click = True
        i += 1

    backend = WindowBackend()
    input_backend = InputBackend()
    win = backend.moments_window()
    backend.activate(win)
    win_rect = backend.rect(win)
    if win_rect is None:
        raise EasyMoneyError("无法读取朋友圈窗口位置")

    print("手动标定刷新按钮位置：请将鼠标移到朋友圈顶部“刷新”按钮中心。")
    countdown()
    point = input_backend.position()
    offset = Point(point.x - win_rect.left, point.y - win_rect.top)
    save_point(CONFIG_REFRESH, offset)
    print(f"已定位刷新按钮: 鼠标位置 ({int(point.x)}, {int(point.y)})")
    print(f"窗口相对偏移: dx={int(offset.x)}, dy={int(offset.y)}")
    print(f"配置已保存: {CONFIG_REFRESH}")
    if test_click:
        input_backend.click(point, interval=0.0)
        print("已执行测试点击")
    return 0


def cmd_run(args: list[str]) -> int:
    target_pos: Optional[Point] = None
    target_index: Optional[int] = None
    target_title: Optional[str] = None
    target_id: Optional[str] = None
    interval = 15.0
    i = 0
    while i < len(args):
        if args[i] == "--pos":
            value, i = parse_option_value(args, i, "--pos")
            target_pos = parse_point_text(value)
        elif args[i] == "--index":
            value, i = parse_option_value(args, i, "--index")
            target_index = int(value)
        elif args[i] == "--title":
            target_title, i = parse_option_value(args, i, "--title")
        elif args[i] == "--id":
            target_id, i = parse_option_value(args, i, "--id")
        elif args[i] == "--interval":
            value, i = parse_option_value(args, i, "--interval")
            interval = max(1.0, float(value))
        i += 1
    input_backend = InputBackend()
    backend = WindowBackend()
    if not any([target_pos, target_index is not None, target_title, target_id]):
        target_pos = refresh_point_from_saved_offset(backend)
        if target_pos is None:
            raise EasyMoneyError("未找到刷新按钮坐标配置，请先运行 locate")
    print(f"自动刷新模式启动，间隔 {interval:g} 秒。按 Ctrl+C 退出。")
    while True:
        now = time.strftime("%H:%M:%S")
        try:
            if target_pos:
                input_backend.click(target_pos)
                print(f"[{now}] 已点击坐标 ({int(target_pos.x)}, {int(target_pos.y)})")
            elif target_index is not None or target_title or target_id:
                win = backend.moments_window()
                buttons = backend.find_buttons(win)
                chosen = None
                if target_index is not None and 0 <= target_index < len(buttons):
                    chosen = buttons[target_index]
                else:
                    for btn in buttons:
                        name = backend._safe_text(btn)
                        automation_id = backend._automation_id(btn)
                        if target_title and target_title not in name:
                            continue
                        if target_id and target_id != automation_id:
                            continue
                        chosen = btn
                        break
                if chosen is None:
                    print(f"[{now}] 未找到目标按钮，跳过")
                elif backend.click_control(chosen, input_backend, prefer_coordinate=True):
                    rect = backend.rect(chosen)
                    detail = f" @({int(rect.center.x)}, {int(rect.center.y)})" if rect else ""
                    print(f"[{now}] 已点击按钮 {backend._safe_text(chosen) or target_index}{detail}")
                else:
                    print(f"[{now}] 点击按钮失败")
        except EasyMoneyError as exc:
            print(f"[{now}] {exc}")
        time.sleep(interval)


def cmd_comment_locate(args: list[str]) -> int:
    backend = WindowBackend()
    win_rect = backend.moments_window_rect()
    input_backend = InputBackend()
    capture = CaptureBackend()
    print("标定评论相关位置（分 3 步）")
    print("步骤 1/3: 请将鼠标移到目标动态右下角操作按钮上。")
    countdown()
    action_pos = input_backend.position()
    template_region = Rect(action_pos.x - 25, action_pos.y - 25, action_pos.x + 25, action_pos.y + 25).clamp_to(win_rect)
    capture.save(capture.screenshot(template_region), ACTION_TEMPLATE)
    print(f"  操作按钮位置: ({int(action_pos.x)}, {int(action_pos.y)})")
    print(f"  操作按钮模板已保存: {ACTION_TEMPLATE}")
    input_backend.click(action_pos)
    time.sleep(0.45)
    print("步骤 2/3: 菜单弹出后，请将鼠标移到“评论”选项上。")
    countdown()
    comment_pos = input_backend.position()
    comment_offset = Point(comment_pos.x - action_pos.x, comment_pos.y - action_pos.y)
    input_backend.click(comment_pos)
    time.sleep(0.45)
    print("步骤 3/3: 输入框弹出后，请将鼠标移到“发送”按钮上。")
    countdown()
    send_pos = input_backend.position()
    send_offset = Point(send_pos.x - action_pos.x, send_pos.y - action_pos.y)
    send_x_ratio = (send_pos.x - win_rect.left) / max(win_rect.width, 1)
    config = CommentConfig(comment_from_action=comment_offset, send_x_ratio=send_x_ratio, send_from_action=send_offset)
    save_comment_config(config)
    print(f"  评论偏移: dx={int(comment_offset.x)}, dy={int(comment_offset.y)}")
    print(f"  发送偏移: dx={int(send_offset.x)}, dy={int(send_offset.y)}")
    print(f"  发送 X 比例: {send_x_ratio:.3f}")
    print(f"标定完成，配置已保存: {CONFIG_COMMENT}")
    return 0


def cmd_comment_fixed_send_locate(args: list[str]) -> int:
    config = load_comment_config()
    if not config:
        raise EasyMoneyError("未找到评论配置，请先运行 comment-locate")
    backend = WindowBackend()
    win_rect = backend.moments_window_rect()
    input_backend = InputBackend()
    print("标定低位操作按钮的固定发送位置（分 2 步）")
    print("步骤 1/2: 请将鼠标移到低位动态右下角操作按钮上。")
    countdown()
    action_pos = input_backend.position()
    action_y_threshold = action_pos.y - win_rect.top
    input_backend.click(action_pos)
    time.sleep(0.35)
    input_backend.click(Point(action_pos.x + config.comment_from_action.x, action_pos.y + config.comment_from_action.y))
    time.sleep(0.45)
    print("步骤 2/2: 请将鼠标移到低位场景下实际发送按钮位置。")
    countdown()
    send_pos = input_backend.position()
    fixed_offset = Point(send_pos.x - win_rect.left, send_pos.y - win_rect.top)
    config.fixed_send_action_y_threshold = action_y_threshold
    config.fixed_send_window_offset = fixed_offset
    save_comment_config(config)
    print(f"  阈值: 窗口内 Y >= {int(action_y_threshold)}")
    print(f"  固定发送偏移: x={int(fixed_offset.x)}, y={int(fixed_offset.y)}")
    return 0


def cmd_post_image_locate(args: list[str], x_only: bool = False) -> int:
    backend = WindowBackend()
    win_rect = backend.moments_window_rect()
    input_backend = InputBackend()
    print("请将鼠标移到目标动态图片上，3 秒后记录。")
    countdown()
    pos = input_backend.position()
    if x_only:
        save_float(CONFIG_POST_IMAGE_TAP_X_OFFSET, pos.x - win_rect.left)
        print(f"图片横坐标偏移已保存: {int(pos.x - win_rect.left)}")
    else:
        offset = Point(pos.x - win_rect.left, pos.y - win_rect.top)
        save_point(CONFIG_POST_IMAGE_TAP_OFFSET, offset)
        print(f"图片轻点偏移已保存: ({int(offset.x)}, {int(offset.y)})")
    return 0


def parse_comment_options(args: list[str]) -> CommentOptions:
    options = CommentOptions(
        open_comment_mode=(os.environ.get("EASYMONEY_OPEN_COMMENT_MODE") or "keys").strip().lower(),
        submit_mode=(os.environ.get("EASYMONEY_SUBMIT_MODE") or "click").strip().lower(),
    )
    user_filter = False
    user_name: Optional[str] = None

    i = 0
    while i < len(args):
        arg = args[i]
        if arg == "--text":
            options.comment_text, i = parse_option_value(args, i, "--text")
        elif arg == "--user":
            user_filter = True
            if i + 1 < len(args) and not args[i + 1].startswith("--"):
                user_name = args[i + 1]
                i += 1
        elif arg in {"--solve-question", "--slove-question"}:
            options.solve_question = True
        elif arg == "--doubao":
            options.solve_question = True
        elif arg == "--LLM":
            options.use_llm = True
            options.solve_question = True
        elif arg == "--vision":
            options.use_vision = True
        elif arg == "--debug":
            options.debug = True
        elif arg == "--save-post-image":
            options.save_post_image = True
        elif arg == "--output":
            value, i = parse_option_value(args, i, "--output")
            options.save_path = expand_path(value)
        elif arg == "--click-post-image":
            options.click_post_image = True
        elif arg in {"--test-image-crop", "--debug-image-crop"}:
            options.test_image_crop = True
        elif arg == "--rounds":
            value, i = parse_option_value(args, i, "--rounds")
            options.rounds = max(1, int(value))
        elif arg == "--submit-keys":
            value, i = parse_option_value(args, i, "--submit-keys")
            options.submit_comment_keys_override = parse_key_sequence_text(value, "--submit-keys")
            options.submit_mode = "keys"
        elif arg == "--submit-enter":
            options.submit_comment_keys_override = ("enter",)
            options.submit_mode = "keys"
        elif arg == "--submit-click":
            options.submit_mode = "click"
        elif arg == "--submit-mode":
            options.submit_mode, i = parse_option_value(args, i, "--submit-mode")
            options.submit_mode = options.submit_mode.strip().lower()
        elif arg == "--open-click":
            options.open_comment_mode = "click"
        elif arg == "--open-keys":
            options.open_comment_mode = "keys"
        elif arg == "--open-mode":
            options.open_comment_mode, i = parse_option_value(args, i, "--open-mode")
            options.open_comment_mode = options.open_comment_mode.strip().lower()
        elif arg in {"--ocr-comment", "--fast", "--stream-capture", "--yolo-debug", "--save-yolo-images"}:
            print(f"提示: Windows v1 暂不完整支持 {arg}，已忽略或降级")
        else:
            raise EasyMoneyError(f"未知 comment 参数: {arg}")
        i += 1

    if not user_filter:
        raise EasyMoneyError("comment 命令必须显式指定 --user <用户名前缀>")
    if options.use_vision and not options.use_llm:
        raise EasyMoneyError("--vision 需要与 --LLM 一起使用")
    if not any([options.comment_text, options.solve_question, options.save_post_image, options.click_post_image, options.test_image_crop]):
        raise EasyMoneyError('请指定 --text "评论内容"，或使用 --solve-question / --doubao / --LLM / --save-post-image')

    options.requested_user = (user_name or "").strip()
    if not options.requested_user:
        raise EasyMoneyError("comment --user 需要提供用户名前缀，用于匹配朋友圈第 2 个 ListItem 的开头")
    return options


def comment_requires_config(options: CommentOptions) -> bool:
    return not any([options.save_post_image, options.click_post_image, options.test_image_crop])


def normalize_comment_mode(mode: str, option_name: str) -> str:
    normalized = mode.strip().lower()
    if normalized in {"mouse", "coordinate"}:
        return "click"
    if normalized == "keyboard":
        return "keys"
    if normalized not in {"keys", "click"}:
        raise EasyMoneyError(f"{option_name} 只支持 keys 或 click")
    return normalized


def resolve_comment_target_post(
    backend: WindowBackend,
    input_backend: InputBackend,
    win: Any,
    requested_user: str,
    window_rect: Rect,
    rounds: int,
) -> tuple[MomentPostResolution, Rect]:
    post: Optional[MomentPostResolution] = None
    last_error: Optional[Exception] = None
    refresh_offset = load_point(CONFIG_REFRESH)
    refresh_capture: Optional[CaptureBackend] = None
    try:
        for round_index in range(1, rounds + 1):
            try:
                print(f"[{current_timestamp_ms()}] UIA用户匹配: 第 {round_index}/{rounds} 轮")
                current_window_rect = backend.rect(win)
                if current_window_rect is None:
                    raise WindowPositionUnavailable("无法读取朋友圈窗口位置")
                window_rect = current_window_rect
                list_item = resolve_second_uia_list_item_post(
                    backend,
                    win,
                    requested_user,
                    item_index=1,
                    settle_ms=int(os.environ.get("EASYMONEY_UIA_USER_SETTLE_MS", "220")),
                    include_text=True,
                )
                post = MomentPostResolution(
                    body_frame=list_item.body_frame,
                    action_point=list_item.action_point,
                    text=list_item.text,
                    source=(
                        f"UIA:ListItem #{list_item.item_index + 1} "
                        f"prefix={list_item.detected_prefix or '(空)'} total={list_item.elapsed_ms}ms"
                    ),
                )
                print(
                    "  UIA用户匹配成功: "
                    f"user={requested_user} "
                    f"item=#{list_item.item_index + 1} "
                    f"prefix={list_item.detected_prefix or '(空)'} "
                    f"frame={list_item.body_frame.describe()} "
                    f"耗时={list_item.elapsed_ms}ms"
                )
                return post, window_rect
            except WindowPositionUnavailable:
                raise
            except Exception as exc:
                last_error = exc
                if round_index >= rounds:
                    print(f"  UIA用户匹配失败: {exc}")
                    break
                if refresh_offset is None:
                    raise EasyMoneyError("UIA用户匹配失败且未找到刷新按钮坐标配置，请先运行 locate") from exc
                print(f"  UIA用户匹配失败，执行刷新后继续: {exc}")
                refresh_region = refresh_observation_region(window_rect)
                try:
                    if refresh_capture is None:
                        refresh_capture = CaptureBackend()
                    baseline_fingerprint = quick_capture_fingerprint(refresh_capture, refresh_region)
                except EasyMoneyError:
                    baseline_fingerprint = None
                refresh_button_center = Point(window_rect.left + refresh_offset.x, window_rect.top + refresh_offset.y)
                input_backend.click(refresh_button_center)
                try:
                    if refresh_capture is None:
                        refresh_capture = CaptureBackend()
                    wait_changed = wait_for_region_refresh(refresh_capture, refresh_region, baseline_fingerprint)
                except EasyMoneyError:
                    wait_changed = False
                    time.sleep(COMMENT_REFRESH_WAIT_SECONDS)
                print(
                    f"  已点击 locate 保存的刷新坐标: ({int(refresh_button_center.x)}, {int(refresh_button_center.y)})，"
                    f"{'检测到刷新变化' if wait_changed else f'等待 {int(COMMENT_REFRESH_WAIT_SECONDS * 1000)}ms'}"
                )
    finally:
        if refresh_capture is not None:
            refresh_capture.close()

    raise EasyMoneyError(f"{last_error or 'UIA用户匹配失败'}；已尝试 {rounds} 轮，可用 --rounds N 调整")


def refresh_comment_window_rect(backend: WindowBackend, win: Any, window_rect: Rect) -> Rect:
    time.sleep(float(os.environ.get("EASYMONEY_UIA_AFTER_CAPTURE_DELAY", "0.02")))
    fresh_rect = backend.rect(win)
    return fresh_rect if fresh_rect is not None else window_rect


def print_resolved_comment_post(requested_user: str, post: MomentPostResolution) -> None:
    print(f"已匹配用户: {requested_user}")
    print(f"动态定位: {post.source} frame={post.body_frame.describe()}")
    if post.text:
        print("正文内容开始")
        print(post.text)
        print("正文内容结束")


def save_comment_post_image(post: MomentPostResolution, window_rect: Rect, save_path: Optional[Path]) -> int:
    image = capture_post_image(post, window_rect)
    output = save_path or DEBUG_DIR / f"wechat_post_image_{time.strftime('%Y%m%d_%H%M%S')}.png"
    ensure_parent(output)
    image.save(output)
    print(f"动态图片/区域已保存: {output}")
    return 0


def click_comment_post_image(
    input_backend: InputBackend,
    post: MomentPostResolution,
    window_rect: Rect,
    debug: bool = False,
) -> int:
    tap_offset = load_point(CONFIG_POST_IMAGE_TAP_OFFSET)
    tap_x = load_float(CONFIG_POST_IMAGE_TAP_X_OFFSET)
    if tap_offset:
        point = Point(window_rect.left + tap_offset.x, window_rect.top + tap_offset.y)
    else:
        x = window_rect.left + tap_x if tap_x is not None else post.body_frame.center.x
        point = Point(x, post.body_frame.center.y)
    if debug:
        input_backend.move_to(point)
        print(f"DEBUG: 鼠标已移动到图片点击点 ({int(point.x)}, {int(point.y)})")
        return 0
    input_backend.click(point)
    print(f"已点击动态图片区域 ({int(point.x)}, {int(point.y)})")
    return 0


def resolve_comment_text(options: CommentOptions, post: MomentPostResolution, window_rect: Rect) -> str:
    final_text = (options.comment_text or "").strip()
    if not options.solve_question:
        return final_text

    context = post.text.strip()
    if not context:
        raise EasyMoneyError("需要自动答题但未能读取朋友圈正文")
    image_urls: list[str] = []
    if options.use_vision:
        image_urls = capture_yolo_image_data_urls(post, window_rect)
        print(f"已附带 YOLO 图片: {len(image_urls)} 张")
    solved = ask_doubao_to_solve_post(context, image_data_urls=image_urls)
    if solved:
        final_text = solved.answer
        print(f"LLM 命中: {solved.answer}")
    if not final_text and options.comment_text:
        final_text = options.comment_text.strip()
        print("自动答题未命中，回退到 --text")
    if not final_text:
        raise EasyMoneyError("未能生成评论内容，请补充 --text 作为回退")
    return final_text


def build_comment_send_plan(
    options: CommentOptions,
    post: MomentPostResolution,
    window_rect: Rect,
    config: CommentConfig,
    final_text: str,
) -> CommentSendPlan:
    open_comment_keys = parse_key_sequence_text(os.environ.get("EASYMONEY_OPEN_COMMENT_KEYS", "tab,enter"), "EASYMONEY_OPEN_COMMENT_KEYS")
    submit_comment_keys = options.submit_comment_keys_override or parse_key_sequence_text(
        os.environ.get("EASYMONEY_SUBMIT_KEYS", "tab,tab,tab,enter"),
        "EASYMONEY_SUBMIT_KEYS",
    )
    open_comment_mode = normalize_comment_mode(options.open_comment_mode, "--open-mode")
    submit_mode = normalize_comment_mode(options.submit_mode, "--submit-mode")
    comment_point = Point(
        post.action_point.x + config.comment_from_action.x,
        post.action_point.y + config.comment_from_action.y,
    )
    comment_open_method = (
        f"点击评论菜单[dx={int(config.comment_from_action.x)},dy={int(config.comment_from_action.y)}]"
        if open_comment_mode == "click"
        else format_key_sequence(open_comment_keys)
    )
    send_point, send_point_method = resolve_send_point(post.action_point, window_rect, config)
    submit_method = f"点击发送按钮[{send_point_method}]" if submit_mode == "click" else format_key_sequence(submit_comment_keys)
    return CommentSendPlan(
        text=final_text,
        action_point=post.action_point,
        comment_point=comment_point,
        send_point=send_point,
        send_point_method=send_point_method,
        open_comment_mode=open_comment_mode,
        submit_mode=submit_mode,
        open_comment_keys=open_comment_keys,
        submit_comment_keys=submit_comment_keys,
        comment_open_method=comment_open_method,
        submit_method=submit_method,
    )


def print_comment_debug(plan: CommentSendPlan, input_backend: InputBackend) -> None:
    input_backend.move_to(plan.action_point)
    print(f"DEBUG: 操作按钮点 ({int(plan.action_point.x)}, {int(plan.action_point.y)})")
    print(f"DEBUG: 打开评论方式: {plan.comment_open_method}")
    print(f"DEBUG: 评论菜单点 ({int(plan.comment_point.x)}, {int(plan.comment_point.y)})")
    print(f"DEBUG: 发送方式: {plan.submit_method}")
    print(f"DEBUG: 发送点参考 [{plan.send_point_method}] ({int(plan.send_point.x)}, {int(plan.send_point.y)})")
    print(f"DEBUG: 评论内容: {plan.text}")


def execute_comment_send_plan(plan: CommentSendPlan, input_backend: InputBackend) -> CommentSendResult:
    if plan.open_comment_mode == "keys":
        input_backend.prepare_key_sequence(plan.open_comment_keys)
    if plan.submit_mode == "keys":
        input_backend.prepare_key_sequence(plan.submit_comment_keys)

    send_flow_start = time.perf_counter()
    step_start = time.perf_counter()
    input_backend.click(plan.action_point, interval=0.0)
    action_click_ms = int((time.perf_counter() - step_start) * 1000)

    step_start = time.perf_counter()
    if plan.open_comment_mode == "click":
        open_click_delay_ms = float(os.environ.get("EASYMONEY_OPEN_CLICK_DELAY_MS", "0"))
        if open_click_delay_ms > 0:
            precise_delay(open_click_delay_ms / 1000.0)
        input_backend.click(plan.comment_point, interval=0.0)
    else:
        input_backend.press_sequence_atomic(plan.open_comment_keys)
    open_comment_ms = int((time.perf_counter() - step_start) * 1000)

    can_type_text_directly = getattr(input_backend, "can_type_text_directly", input_backend.can_type_directly)
    step_start = time.perf_counter()
    if can_type_text_directly(plan.text):
        text_input_method = input_backend.type_text_directly(plan.text)
    else:
        text_input_method = input_backend.paste_text(
            plan.text,
            restore_clipboard=False,
            before_paste_delay=0.0,
            after_paste_delay=0.012,
        )
    text_input_ms = int((time.perf_counter() - step_start) * 1000)

    send_step_label = "发送点击" if plan.submit_mode == "click" else "发送快捷键"
    step_start = time.perf_counter()
    if plan.submit_mode == "click":
        input_backend.click(plan.send_point, interval=0.0)
    else:
        input_backend.press_sequence_atomic(plan.submit_comment_keys)
    send_submit_ms = int((time.perf_counter() - step_start) * 1000)
    total_send_ms = int((time.perf_counter() - send_flow_start) * 1000)
    return CommentSendResult(
        text_input_method=text_input_method,
        action_click_ms=action_click_ms,
        open_comment_ms=open_comment_ms,
        text_input_ms=text_input_ms,
        send_submit_ms=send_submit_ms,
        total_send_ms=total_send_ms,
        send_step_label=send_step_label,
    )


def print_comment_send_result(plan: CommentSendPlan, result: CommentSendResult) -> None:
    print(
        f"已执行评论发送: {result.text_input_method} | 打开评论={plan.comment_open_method} | "
        f"发送方式={plan.submit_method} | 发送点参考=({int(plan.send_point.x)}, {int(plan.send_point.y)})"
    )
    print(
        f"发送流程耗时: 总计={result.total_send_ms}ms | "
        f"点操作={result.action_click_ms}ms | 打开评论={result.open_comment_ms}ms | "
        f"输入={result.text_input_ms}ms | {result.send_step_label}={result.send_submit_ms}ms"
    )


def cmd_comment(args: list[str]) -> int:
    options = parse_comment_options(args)
    config = load_comment_config()
    if not config and comment_requires_config(options):
        raise EasyMoneyError("未找到评论配置，请先运行 comment-locate")

    backend = WindowBackend()
    input_backend = InputBackend()
    win = backend.moments_window()
    backend.activate(win)
    window_rect = backend.rect(win)
    if window_rect is None:
        raise WindowPositionUnavailable("无法读取朋友圈窗口位置")

    post, window_rect = resolve_comment_target_post(
        backend,
        input_backend,
        win,
        options.requested_user,
        window_rect,
        options.rounds,
    )
    window_rect = refresh_comment_window_rect(backend, win, window_rect)
    print_resolved_comment_post(options.requested_user, post)

    if options.save_post_image or options.test_image_crop:
        return save_comment_post_image(post, window_rect, options.save_path)
    if options.click_post_image:
        return click_comment_post_image(input_backend, post, window_rect, debug=options.debug)

    final_text = resolve_comment_text(options, post, window_rect)
    if config is None:
        raise EasyMoneyError("未找到评论配置，请先运行 comment-locate")
    plan = build_comment_send_plan(options, post, window_rect, config, final_text)
    if options.debug:
        print_comment_debug(plan, input_backend)
        return 0

    result = execute_comment_send_plan(plan, input_backend)
    print_comment_send_result(plan, result)
    return 0


def cmd_llm(args: list[str]) -> int:
    if len(args) < 2 or args[0] != "ask":
        raise EasyMoneyError('用法: llm ask "<问题>" [上下文]')
    answer = ask_local_llm(args[1], args[2] if len(args) >= 3 else "")
    if not answer:
        return 1
    print(f"回答: {answer}")
    return 0


def cmd_doubao(args: list[str]) -> int:
    if len(args) < 2 or args[0] != "ask":
        raise EasyMoneyError('用法: doubao ask "<朋友圈正文>"')
    solved = ask_doubao_to_solve_post(args[1])
    if not solved:
        print("未能回答")
        return 1
    print(f"答案: {solved.answer}")
    if solved.evidence:
        print(f"证据: {solved.evidence}")
    print(f"置信度: {solved.confidence:.2f}")
    return 0


def main(argv: Optional[list[str]] = None) -> int:
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            try:
                reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass
    enable_dpi_awareness()
    args = list(sys.argv[1:] if argv is None else argv)
    if not args or args[0] in {"-h", "--help", "help"}:
        print_usage()
        return 0
    mode = args[0]
    rest = args[1:]
    dispatch: dict[str, Callable[[list[str]], int]] = {
        "uia-dump": cmd_uia_dump,
        "ax-dump": cmd_uia_dump,
        "capture-info": cmd_capture_info,
        "locate": cmd_locate,
        "run": cmd_run,
        "comment-locate": cmd_comment_locate,
        "comment-fixed-send-locate": cmd_comment_fixed_send_locate,
        "post-image-locate": lambda a: cmd_post_image_locate(a, x_only=False),
        "post-image-x-locate": lambda a: cmd_post_image_locate(a, x_only=True),
        "comment": cmd_comment,
        "llm": cmd_llm,
        "doubao": cmd_doubao,
    }
    handler = dispatch.get(mode)
    if handler is None:
        print_usage()
        raise EasyMoneyError(f"未知命令: {mode}")
    return handler(rest)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\n已停止")
        raise SystemExit(130)
    except EasyMoneyError as exc:
        print(f"错误: {exc}")
        raise SystemExit(1)
