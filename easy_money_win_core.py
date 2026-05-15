#!/usr/bin/env python3

from __future__ import annotations

import base64
import ctypes
import importlib
import logging
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
    inline_image_count: Optional[int] = None


@dataclass
class UIAListItemResolution:
    item_index: int
    body_frame: Rect
    action_point: Point
    text: str
    expected_user_id: str
    detected_prefix: str
    elapsed_ms: int
    inline_image_count: Optional[int] = None


@dataclass
class CommentOptions:
    comment_text: Optional[str] = None
    requested_user: str = ""
    solve_question: bool = False
    use_llm: bool = False
    use_vision: bool = False
    save_vision_image: bool = False
    vision_save_path: Optional[Path] = None
    debug: bool = False
    save_post_image: bool = False
    save_path: Optional[Path] = None
    click_post_image: bool = False
    test_image_crop: bool = False
    rounds: int = 30
    submit_mode: str = "click"
    timing_detail: bool = False


@dataclass
class CommentSendPlan:
    text: str
    action_point: Point
    send_point: Point
    send_point_method: str
    submit_mode: str
    open_comment_keys: tuple[str, ...]
    submit_comment_keys: tuple[str, ...]
    submit_method: str
    timing_detail: bool = False


@dataclass
class CommentSendResult:
    text_input_method: str
    action_click_ms: int
    open_comment_ms: int
    text_input_ms: int
    send_submit_ms: int
    total_send_ms: int
    send_step_label: str
    input_timings: tuple[tuple[str, int], ...] = ()


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
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # 按显示器感知 DPI。
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


def print_ts(message: str) -> None:
    print(f"[{current_timestamp_ms()}] {message}")


def configure_timestamped_logging() -> None:
    class SuppressDxcamSingletonWarning(logging.Filter):
        def filter(self, record: logging.LogRecord) -> bool:
            return "DXCamera instance already exists" not in record.getMessage()

    logging.basicConfig(
        level=logging.WARNING,
        format="[%(asctime)s.%(msecs)03d] %(message)s",
        datefmt="%H:%M:%S",
        force=True,
    )
    logging.getLogger("dxcam").addFilter(SuppressDxcamSingletonWarning())


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


def chinese_image_count_value(raw: str) -> Optional[int]:
    text = raw.strip()
    if not text:
        return None
    try:
        return int(text)
    except ValueError:
        pass
    values = {
        "一": 1,
        "二": 2,
        "两": 2,
        "三": 3,
        "四": 4,
        "五": 5,
        "六": 6,
        "七": 7,
        "八": 8,
        "九": 9,
    }
    if text == "十":
        return 10
    if text.startswith("十") and len(text) > 1:
        ones = values.get(text[-1])
        return 10 + ones if ones is not None else None
    if text.endswith("十") and len(text) > 1:
        tens = values.get(text[0])
        return tens * 10 if tens is not None else None
    if "十" in text:
        first, _, last = text.partition("十")
        tens = values.get(first[:1])
        ones = values.get(last[:1])
        if tens is not None and ones is not None:
            return tens * 10 + ones
    if len(text) == 1:
        return values.get(text)
    return None


def extract_inline_image_count(text: str) -> Optional[int]:
    match = re.search(r"包含\s*([1-9])\s*张图片", text)
    if not match:
        return None
    return int(match.group(1))
