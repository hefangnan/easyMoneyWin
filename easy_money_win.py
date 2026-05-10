#!/usr/bin/env python3
"""Windows Python port of easyMoney.swift.

This is a practical first Windows version, not a one-to-one Swift rewrite.
The UI layer uses Windows UI Automation plus coordinate fallbacks; KB and LLM
commands are kept usable without loading UI automation dependencies.
"""

from __future__ import annotations

import base64
import ctypes
import hashlib
import importlib
import json
import os
import re
import sqlite3
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Iterable, Optional


APP_NAME = "easyMoney Windows"
APP_VERSION = "0.1.0"
SCHEMA_VERSION = 24


class EasyMoneyError(RuntimeError):
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


@dataclass
class CommentConfig:
    comment_from_action: Point
    send_x_ratio: float = 0.8
    send_from_action: Optional[Point] = None
    fixed_send_action_y_threshold: Optional[float] = None
    fixed_send_window_offset: Optional[Point] = None


@dataclass
class UserTemplateEntry:
    template: str
    matchMode: Optional[str] = None
    threshold: Optional[float] = None


@dataclass
class UserTemplateConfig:
    default: Optional[str]
    users: dict[str, UserTemplateEntry]


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
    avatar_center: Point
    body_frame: Rect
    action_point: Point
    text: str
    source: str


HOME = Path.home()
EASYMONEY_DIR = HOME / ".easyMoney"
CONFIG_REFRESH_OFFSET = HOME / ".wechat_refresh_offset"
CONFIG_COMMENT = HOME / ".wechat_comment_config"
CONFIG_AVATAR_OFFSET = HOME / ".wechat_avatar_offset"
CONFIG_POST_IMAGE_TAP_OFFSET = HOME / ".wechat_post_image_tap_offset"
CONFIG_POST_IMAGE_TAP_X_OFFSET = HOME / ".wechat_post_image_tap_x_offset"
CONFIG_USER_TEMPLATES = HOME / ".wechat_user_templates.json"
CONFIG_KB = HOME / ".wechat_kb.sqlite"
CONFIG_PREFIX_CACHE = EASYMONEY_DIR / "doubaotext-prefix-cache.json"
ACTION_TEMPLATE = HOME / ".wechat_action_tpl.png"
LEGACY_AVATAR_TEMPLATE = HOME / ".wechat_avatar_tpl.png"
USER_PHOTO_DIR = Path(os.environ.get("EASYMONEY_USER_PHOTO_DIR", str(EASYMONEY_DIR / "userPhoto")))
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


def atomic_write_json(path: Path, data: Any) -> None:
    ensure_parent(path)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8")
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


def normalized_avatar_match_mode(raw: Optional[str]) -> Optional[str]:
    if raw is None:
        return None
    value = raw.strip().lower().replace("-", "_")
    if value in {"center_square", "square", "anchor", "anchored", "center"}:
        return "center_square"
    if value in {"wide", "classic", "legacy", "left", "left_wide", "full"}:
        return "wide"
    return None


def sanitize_user_template_name(raw: str) -> str:
    cleaned = re.sub(r'[/:\\\0\r\n]+', "_", raw.strip())
    return cleaned or "user"


def named_template_path(user_name: str) -> Path:
    return USER_PHOTO_DIR / f"{sanitize_user_template_name(user_name)}.png"


def load_user_template_config() -> UserTemplateConfig:
    if not CONFIG_USER_TEMPLATES.exists():
        return UserTemplateConfig(default=None, users={})
    try:
        root = json.loads(CONFIG_USER_TEMPLATES.read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"警告: 用户模板配置解析失败，将使用空配置: {exc}")
        return UserTemplateConfig(default=None, users={})
    users: dict[str, UserTemplateEntry] = {}
    for name, item in (root.get("users") or {}).items():
        if isinstance(item, dict) and item.get("template"):
            users[str(name)] = UserTemplateEntry(
                template=str(item["template"]),
                matchMode=item.get("matchMode"),
                threshold=float(item["threshold"]) if item.get("threshold") is not None else None,
            )
    return UserTemplateConfig(default=root.get("default"), users=users)


def save_user_template_config(config: UserTemplateConfig) -> None:
    root = {
        "default": config.default,
        "users": {
            name: {
                "template": entry.template,
                "matchMode": entry.matchMode,
                "threshold": entry.threshold,
            }
            for name, entry in config.users.items()
        },
    }
    atomic_write_json(CONFIG_USER_TEMPLATES, root)


def discover_user_photo_templates() -> dict[str, UserTemplateEntry]:
    if not USER_PHOTO_DIR.exists():
        return {}
    allowed = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}
    found: dict[str, UserTemplateEntry] = {}
    for path in USER_PHOTO_DIR.iterdir():
        if path.is_file() and path.suffix.lower() in allowed:
            found[path.stem.strip()] = UserTemplateEntry(template=str(path), matchMode=None, threshold=None)
    return found


def resolve_user_template(user_name: Optional[str]) -> tuple[Optional[str], Optional[UserTemplateEntry], Optional[Path]]:
    config = load_user_template_config()
    merged = discover_user_photo_templates()
    merged.update(config.users)
    name = user_name or config.default
    if name:
        entry = merged.get(name)
        if entry:
            path = expand_path(entry.template)
            return name, entry, path
        return name, None, None
    if config.default and config.default in merged:
        entry = merged[config.default]
        return config.default, entry, expand_path(entry.template)
    return None, None, None


class InputBackend:
    def __init__(self) -> None:
        self.pyautogui = require_module("pyautogui")
        self.pyperclip = require_module("pyperclip")
        self.pyautogui.PAUSE = 0.02

    def position(self) -> Point:
        x, y = self.pyautogui.position()
        return Point(float(x), float(y))

    def click(self, point: Point, clicks: int = 1, interval: float = 0.04) -> None:
        x, y = point.rounded()
        self.pyautogui.click(x=x, y=y, clicks=clicks, interval=interval, button="left")

    def move_to(self, point: Point) -> None:
        x, y = point.rounded()
        self.pyautogui.moveTo(x=x, y=y)

    def press(self, key: str) -> None:
        self.pyautogui.press(key)

    def hotkey(self, *keys: str) -> None:
        self.pyautogui.hotkey(*keys)

    def paste_text(self, text: str, restore_clipboard: bool = True) -> str:
        old_text: Optional[str]
        try:
            old_text = self.pyperclip.paste()
        except Exception:
            old_text = None
        self.pyperclip.copy(text)
        time.sleep(0.05)
        self.hotkey("ctrl", "v")
        time.sleep(0.12)
        if restore_clipboard and old_text is not None:
            try:
                self.pyperclip.copy(old_text)
            except Exception:
                pass
        return "剪贴板粘贴"


class CaptureBackend:
    def __init__(self) -> None:
        self.mss_mod = require_module("mss")
        self.Image = require_module("PIL.Image", "Pillow")

    def screenshot(self, rect: Rect):
        if rect.width <= 0 or rect.height <= 0:
            raise EasyMoneyError(f"截图区域无效: {rect.describe()}")
        with self.mss_mod.mss() as sct:
            shot = sct.grab(rect.to_mss())
            return self.Image.frombytes("RGB", shot.size, shot.rgb)

    def save(self, image: Any, path: Path) -> Path:
        ensure_parent(path)
        image.save(path)
        return path


class WindowBackend:
    def __init__(self) -> None:
        self.pywinauto = require_module("pywinauto")
        self.Desktop = getattr(self.pywinauto, "Desktop")
        self.desktop = self.Desktop(backend="uia")

    @staticmethod
    def _safe_text(control: Any) -> str:
        for attr in ("window_text",):
            try:
                return getattr(control, attr)() or ""
            except Exception:
                pass
        try:
            return control.element_info.name or ""
        except Exception:
            return ""

    @staticmethod
    def _control_type(control: Any) -> str:
        try:
            return str(control.element_info.control_type or "")
        except Exception:
            return ""

    @staticmethod
    def _automation_id(control: Any) -> str:
        try:
            return str(control.element_info.automation_id or "")
        except Exception:
            return ""

    @staticmethod
    def rect(control: Any) -> Optional[Rect]:
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
        except Exception as exc:
            raise EasyMoneyError(f"读取窗口列表失败: {exc}") from exc

    def moments_window(self) -> Any:
        candidates = []
        for win in self.windows():
            title = self._safe_text(win).strip()
            if not title:
                continue
            if "朋友圈" in title:
                return win
            if "微信" in title or "WeChat" in title or "wechat" in title.lower():
                candidates.append(win)
        if candidates:
            print("警告: 未找到标题为“朋友圈”的窗口，使用疑似微信窗口；如定位异常，请先打开朋友圈窗口。")
            return candidates[0]
        raise EasyMoneyError("未找到微信/朋友圈窗口，请先打开微信桌面版并进入朋友圈")

    def moments_window_rect(self) -> Rect:
        win = self.moments_window()
        rect = self.rect(win)
        if rect is None:
            raise EasyMoneyError("无法读取朋友圈窗口位置")
        return rect

    def activate(self, control: Any) -> None:
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
            return []

    def iter_tree(self, control: Any, max_depth: int = 10) -> Iterable[tuple[Any, int]]:
        stack: list[tuple[Any, int]] = [(control, 0)]
        seen: set[int] = set()
        while stack:
            item, depth = stack.pop()
            marker = id(item)
            if marker in seen or depth > max_depth:
                continue
            seen.add(marker)
            yield item, depth
            kids = self.children(item)
            for child in reversed(kids):
                stack.append((child, depth + 1))

    def dump_tree(self, control: Any, max_depth: int = 10, buttons_only: bool = False) -> None:
        count = 0
        for node, depth in self.iter_tree(control, max_depth=max_depth):
            control_type = self._control_type(node)
            if buttons_only and control_type.lower() != "button":
                continue
            rect = self.rect(node)
            name = self._safe_text(node)
            automation_id = self._automation_id(node)
            indent = "  " * depth
            parts = [f"{indent}[{control_type or '?'}]"]
            if name:
                parts.append(f'name="{name}"')
            if automation_id:
                parts.append(f'id="{automation_id}"')
            if rect:
                parts.append(f"rect={rect.describe()}")
            print(" ".join(parts))
            count += 1
        if buttons_only:
            print(f"\n按钮数量: {count}")

    def find_buttons(self, control: Any, max_depth: int = 12) -> list[Any]:
        return [node for node, _ in self.iter_tree(control, max_depth=max_depth) if self._control_type(node).lower() == "button"]

    def click_control(self, control: Any, input_backend: Optional[InputBackend] = None) -> bool:
        try:
            control.invoke()
            return True
        except Exception:
            pass
        rect = self.rect(control)
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


def save_avatar_template_at(point: Point, window_rect: Rect, path: Path, side: int = 80) -> Path:
    capture = CaptureBackend()
    half = side / 2
    region = Rect(point.x - half, point.y - half, point.x + half, point.y + half).clamp_to(window_rect)
    image = capture.screenshot(region)
    capture.save(image, path)
    return path


def template_match(screen_image: Any, template_path: Path, threshold: float = 0.72) -> Optional[tuple[Point, float]]:
    if not template_path.exists():
        return None
    cv2 = require_module("cv2", "opencv-python")
    np = require_module("numpy")
    try:
        tpl_pil = require_module("PIL.Image", "Pillow").open(template_path).convert("RGB")
    except Exception as exc:
        raise EasyMoneyError(f"头像模板读取失败: {template_path} ({exc})") from exc
    screen = cv2.cvtColor(np.array(screen_image.convert("RGB")), cv2.COLOR_RGB2GRAY)
    tpl = cv2.cvtColor(np.array(tpl_pil), cv2.COLOR_RGB2GRAY)
    if tpl.shape[0] > screen.shape[0] or tpl.shape[1] > screen.shape[1]:
        return None
    result = cv2.matchTemplate(screen, tpl, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(result)
    if max_val < threshold:
        return None
    x = max_loc[0] + tpl.shape[1] / 2
    y = max_loc[1] + tpl.shape[0] / 2
    return Point(float(x), float(y)), float(max_val)


def find_avatar_in_window(window_rect: Rect, template_path: Path, threshold: float, match_mode: str) -> tuple[Point, float]:
    capture = CaptureBackend()
    if match_mode == "wide":
        region = Rect(window_rect.left, window_rect.top, window_rect.left + window_rect.width * 0.30, window_rect.bottom)
    else:
        region = Rect(window_rect.left, window_rect.top, window_rect.left + window_rect.width * 0.28, window_rect.top + window_rect.height * 0.78)
    image = capture.screenshot(region)
    match = template_match(image, template_path, threshold=threshold)
    if match is None:
        raise EasyMoneyError(f"未匹配到目标头像，模板={template_path} 阈值={threshold:.2f}")
    local_center, score = match
    return Point(region.left + local_center.x, region.top + local_center.y), score


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


def resolve_moment_post(window_backend: WindowBackend, win: Any, window_rect: Rect, avatar_center: Point, require_text: bool) -> MomentPostResolution:
    candidates: list[tuple[float, Rect, str, str]] = []
    for node, depth in window_backend.iter_tree(win, max_depth=8):
        rect = window_backend.rect(node)
        if rect is None:
            continue
        if not rect.intersects_y(avatar_center.y, tolerance=26):
            continue
        if rect.width < 80 or rect.height < 12:
            continue
        if rect.right < window_rect.left + window_rect.width * 0.20:
            continue
        name = clean_post_text(window_backend._safe_text(node))
        control_type = window_backend._control_type(node)
        if not name and require_text:
            continue
        if name and len(name) < 2:
            continue
        if rect.left < window_rect.left or rect.right > window_rect.right + 8:
            continue
        distance = abs(rect.center.y - avatar_center.y)
        area_penalty = min(rect.width * rect.height / max(window_rect.width * window_rect.height, 1), 1.0)
        text_bonus = -20 if name else 0
        score = distance + area_penalty * 25 + text_bonus
        candidates.append((score, rect, name, control_type))

    if candidates:
        candidates.sort(key=lambda item: item[0])
        _, body_frame, text, control_type = candidates[0]
        action_point = Point(body_frame.right - 40, body_frame.bottom - 10)
        return MomentPostResolution(
            avatar_center=avatar_center,
            body_frame=body_frame,
            action_point=action_point,
            text=text,
            source=f"UIA:{control_type or 'unknown'}",
        )

    if require_text:
        raise EasyMoneyError("已匹配头像，但 UIA 未能读取该动态正文；可先用 --text，或检查微信窗口/缩放")

    fallback = Rect(
        window_rect.left + window_rect.width * 0.20,
        max(window_rect.top, avatar_center.y - 48),
        window_rect.right - 16,
        min(window_rect.bottom, avatar_center.y + 140),
    )
    return MomentPostResolution(
        avatar_center=avatar_center,
        body_frame=fallback,
        action_point=Point(fallback.right - 40, fallback.bottom - 10),
        text="",
        source="fallback-geometry",
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


def sqlite_row_dict(cursor: sqlite3.Cursor, row: sqlite3.Row) -> dict[str, Any]:
    return {desc[0]: row[idx] for idx, desc in enumerate(cursor.description or [])}


def open_knowledge_db(create: bool = True) -> sqlite3.Connection:
    ensure_parent(CONFIG_KB)
    conn = sqlite3.connect(str(CONFIG_KB))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA busy_timeout=1500")
    conn.execute("PRAGMA foreign_keys=ON")
    if create:
        setup_knowledge_schema(conn)
    return conn


def setup_knowledge_schema(conn: sqlite3.Connection) -> None:
    statements = [
        "PRAGMA journal_mode=WAL",
        """
        CREATE TABLE IF NOT EXISTS stores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            canonical_name TEXT NOT NULL,
            normalized_name TEXT NOT NULL,
            notes TEXT NOT NULL DEFAULT '',
            last_seen_at INTEGER DEFAULT (strftime('%s','now')),
            created_at INTEGER DEFAULT (strftime('%s','now')),
            UNIQUE(normalized_name)
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS scripts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            normalized_title TEXT NOT NULL DEFAULT '',
            host_name TEXT NOT NULL DEFAULT '',
            variant TEXT NOT NULL DEFAULT '',
            store_id INTEGER REFERENCES stores(id),
            common_script_id INTEGER REFERENCES scripts(id),
            created_at INTEGER DEFAULT (strftime('%s','now')),
            UNIQUE(title, host_name, variant)
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS posts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            raw_text TEXT NOT NULL,
            source_type TEXT NOT NULL DEFAULT 'auto',
            store_id INTEGER REFERENCES stores(id),
            post_date INTEGER,
            content_signature TEXT NOT NULL DEFAULT '',
            created_at INTEGER DEFAULT (strftime('%s','now'))
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS post_scripts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            post_id INTEGER NOT NULL REFERENCES posts(id),
            script_id INTEGER NOT NULL REFERENCES scripts(id),
            raw_line TEXT,
            created_at INTEGER DEFAULT (strftime('%s','now')),
            UNIQUE(post_id, script_id)
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS role_mappings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            script_id INTEGER NOT NULL REFERENCES scripts(id),
            actor TEXT NOT NULL,
            npc TEXT NOT NULL,
            raw_line TEXT,
            source_confidence REAL DEFAULT 0.8,
            last_seen_at INTEGER DEFAULT (strftime('%s','now')),
            created_at INTEGER DEFAULT (strftime('%s','now')),
            UNIQUE(script_id, actor, npc)
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS script_characters (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            script_id INTEGER NOT NULL REFERENCES scripts(id),
            character_name TEXT NOT NULL,
            alias TEXT NOT NULL DEFAULT '',
            gender_tag TEXT NOT NULL DEFAULT '',
            raw_text TEXT,
            source_confidence REAL DEFAULT 0.8,
            last_seen_at INTEGER DEFAULT (strftime('%s','now')),
            created_at INTEGER DEFAULT (strftime('%s','now')),
            UNIQUE(script_id, character_name)
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS script_roles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            script_id INTEGER NOT NULL REFERENCES scripts(id),
            canonical_name TEXT NOT NULL,
            gender_tag TEXT NOT NULL DEFAULT '',
            role_kind TEXT NOT NULL DEFAULT 'unknown',
            is_player INTEGER NOT NULL DEFAULT 0,
            is_non_player INTEGER NOT NULL DEFAULT 0,
            is_npc INTEGER NOT NULL DEFAULT 0,
            is_dm_role INTEGER NOT NULL DEFAULT 0,
            is_companion_npc INTEGER NOT NULL DEFAULT 0,
            raw_text TEXT NOT NULL DEFAULT '',
            summary TEXT NOT NULL DEFAULT '',
            tags TEXT NOT NULL DEFAULT '',
            source_confidence REAL DEFAULT 0.8,
            last_seen_at INTEGER DEFAULT (strftime('%s','now')),
            created_at INTEGER DEFAULT (strftime('%s','now')),
            UNIQUE(script_id, canonical_name)
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS script_profiles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            script_id INTEGER NOT NULL REFERENCES scripts(id),
            category TEXT NOT NULL DEFAULT '',
            author TEXT NOT NULL DEFAULT '',
            difficulty TEXT NOT NULL DEFAULT '',
            player_config_text TEXT NOT NULL DEFAULT '',
            male_count INTEGER,
            female_count INTEGER,
            min_players INTEGER,
            max_players INTEGER,
            duration_text TEXT NOT NULL DEFAULT '',
            duration_min_hours REAL,
            duration_max_hours REAL,
            summary TEXT NOT NULL DEFAULT '',
            publisher TEXT NOT NULL DEFAULT '',
            age_rating TEXT NOT NULL DEFAULT '',
            release_type TEXT NOT NULL DEFAULT '',
            cross_gender_allowed INTEGER,
            last_seen_at INTEGER DEFAULT (strftime('%s','now')),
            created_at INTEGER DEFAULT (strftime('%s','now')),
            UNIQUE(script_id)
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS facts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            script_id INTEGER NOT NULL REFERENCES scripts(id),
            key TEXT NOT NULL,
            value TEXT NOT NULL,
            raw_line TEXT,
            source_confidence REAL DEFAULT 0.8,
            last_seen_at INTEGER DEFAULT (strftime('%s','now')),
            created_at INTEGER DEFAULT (strftime('%s','now'))
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS store_offerings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            script_id INTEGER NOT NULL REFERENCES scripts(id),
            store_id INTEGER NOT NULL REFERENCES stores(id),
            effective_label TEXT NOT NULL DEFAULT '',
            effective_from INTEGER,
            effective_to INTEGER,
            is_active INTEGER NOT NULL DEFAULT 1,
            price_text TEXT NOT NULL DEFAULT '',
            price_value INTEGER,
            dm_count INTEGER,
            npc_count INTEGER,
            config_text TEXT NOT NULL DEFAULT '',
            notes TEXT NOT NULL DEFAULT '',
            source_confidence REAL DEFAULT 0.9,
            last_seen_at INTEGER DEFAULT (strftime('%s','now')),
            created_at INTEGER DEFAULT (strftime('%s','now')),
            UNIQUE(script_id, store_id, effective_label)
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS store_castings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            script_id INTEGER NOT NULL REFERENCES scripts(id),
            store_id INTEGER NOT NULL REFERENCES stores(id),
            actor_id INTEGER,
            character_name TEXT NOT NULL,
            actor_name TEXT NOT NULL,
            role_kind TEXT NOT NULL DEFAULT 'unknown',
            card_variant TEXT NOT NULL DEFAULT '',
            tags TEXT NOT NULL DEFAULT '',
            effective_label TEXT NOT NULL DEFAULT '',
            effective_from INTEGER,
            effective_to INTEGER,
            is_active INTEGER NOT NULL DEFAULT 1,
            raw_line TEXT,
            source_confidence REAL DEFAULT 0.9,
            last_seen_at INTEGER DEFAULT (strftime('%s','now')),
            created_at INTEGER DEFAULT (strftime('%s','now'))
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS qa_pairs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            script_id INTEGER REFERENCES scripts(id),
            normalized_question TEXT NOT NULL,
            question TEXT NOT NULL,
            answer TEXT NOT NULL,
            evidence TEXT,
            source_confidence REAL DEFAULT 0.9,
            hit_count INTEGER DEFAULT 0,
            last_seen_at INTEGER DEFAULT (strftime('%s','now')),
            created_at INTEGER DEFAULT (strftime('%s','now'))
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS script_aliases (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            script_id INTEGER NOT NULL REFERENCES scripts(id),
            alias TEXT NOT NULL,
            source_confidence REAL DEFAULT 0.7,
            last_seen_at INTEGER DEFAULT (strftime('%s','now')),
            created_at INTEGER DEFAULT (strftime('%s','now')),
            UNIQUE(script_id, alias)
        )
        """,
        """
        CREATE VIRTUAL TABLE IF NOT EXISTS kb_fts USING fts5(
            content,
            source_table,
            source_id UNINDEXED,
            tokenize='unicode61 remove_diacritics 2'
        )
        """,
        "CREATE INDEX IF NOT EXISTS idx_scripts_title ON scripts(title)",
        "CREATE INDEX IF NOT EXISTS idx_qa_pairs_question ON qa_pairs(normalized_question)",
        "PRAGMA user_version=24",
    ]
    for sql in statements:
        try:
            conn.execute(sql)
        except sqlite3.Error as exc:
            if "fts5" in str(exc).lower():
                print(f"警告: 当前 SQLite 不支持 FTS5，搜索会降级: {exc}")
            else:
                raise
    conn.commit()


def normalize_text(text: str) -> str:
    text = text.lower()
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"[，。、“”‘’！!？?：:；;（）()\[\]【】《》<>\"']", "", text)
    return text


def db_table_exists(conn: sqlite3.Connection, table: str) -> bool:
    row = conn.execute("SELECT 1 FROM sqlite_master WHERE type IN ('table','view') AND name=? LIMIT 1", (table,)).fetchone()
    return row is not None


def resolve_store_id(conn: sqlite3.Connection, store_name: Optional[str]) -> Optional[int]:
    if not store_name or not db_table_exists(conn, "stores"):
        return None
    norm = normalize_text(store_name)
    row = conn.execute(
        "SELECT id FROM stores WHERE normalized_name=? OR canonical_name LIKE ? ORDER BY last_seen_at DESC LIMIT 1",
        (norm, f"%{store_name}%"),
    ).fetchone()
    return int(row["id"]) if row else None


def guess_script_ids(conn: sqlite3.Connection, question: str, context: str = "") -> list[int]:
    if not db_table_exists(conn, "scripts"):
        return []
    combined = f"{question}\n{context}"
    normalized = normalize_text(combined)
    rows = conn.execute("SELECT id, title, normalized_title FROM scripts ORDER BY length(title) DESC LIMIT 4000").fetchall()
    hits: list[int] = []
    for row in rows:
        title = row["title"] or ""
        nt = row["normalized_title"] or normalize_text(title)
        if title and title in combined or nt and nt in normalized:
            hits.append(int(row["id"]))
            if len(hits) >= 8:
                break
    quoted = re.findall(r"[《「“\"]([^》」”\"]{2,40})[》」”\"]", combined)
    for title in quoted:
        row = conn.execute("SELECT id FROM scripts WHERE title LIKE ? ORDER BY created_at DESC LIMIT 1", (f"%{title}%",)).fetchone()
        if row and int(row["id"]) not in hits:
            hits.append(int(row["id"]))
    return hits


def sql_scope(script_ids: list[int]) -> tuple[str, list[Any]]:
    if not script_ids:
        return "", []
    placeholders = ",".join("?" for _ in script_ids)
    return f" AND script_id IN ({placeholders}) ", list(script_ids)


def solve_exact_qa(conn: sqlite3.Connection, question: str, script_ids: list[int]) -> Optional[SolvedQuestion]:
    if not db_table_exists(conn, "qa_pairs"):
        return None
    norm = normalize_text(question)
    scope, params = sql_scope(script_ids)
    rows = conn.execute(
        f"""
        SELECT id, answer, evidence, source_confidence
        FROM qa_pairs
        WHERE (normalized_question=? OR question LIKE ?) {scope}
        ORDER BY source_confidence DESC, hit_count DESC, last_seen_at DESC
        LIMIT 1
        """,
        [norm, f"%{question.strip()}%"] + params,
    ).fetchall()
    if not rows:
        return None
    row = rows[0]
    conn.execute("UPDATE qa_pairs SET hit_count = COALESCE(hit_count,0) + 1, last_seen_at=strftime('%s','now') WHERE id=?", (row["id"],))
    conn.commit()
    return SolvedQuestion(answer=row["answer"], evidence=row["evidence"] or "qa_pairs", confidence=float(row["source_confidence"] or 0.9), source="qa_pairs")


def extract_actor_or_role(question: str, patterns: list[str]) -> Optional[str]:
    for pattern in patterns:
        match = re.search(pattern, question)
        if match:
            value = match.group(1).strip(" ？?，,。.")
            if 1 <= len(value) <= 20:
                return value
    return None


def solve_actor_mapping(conn: sqlite3.Connection, question: str, script_ids: list[int], store_id: Optional[int]) -> Optional[SolvedQuestion]:
    actor = extract_actor_or_role(question, [r"(.{1,20}?)(?:演谁|演的是谁|饰演谁|扮演谁)", r"(?:演员|dm|npc)?(.{1,20}?)(?:对应|对应的)(?:角色|npc)"])
    reverse = extract_actor_or_role(question, [r"谁(?:演|饰演|扮演)(.{1,20})", r"(.{1,20})是谁演的"])
    if actor and db_table_exists(conn, "role_mappings"):
        scope, params = sql_scope(script_ids)
        rows = conn.execute(
            f"""
            SELECT actor, npc, raw_line, source_confidence
            FROM role_mappings
            WHERE actor LIKE ? {scope}
            ORDER BY source_confidence DESC, last_seen_at DESC
            LIMIT 5
            """,
            [f"%{actor}%"] + params,
        ).fetchall()
        if rows:
            answer = "、".join(dict.fromkeys(row["npc"] for row in rows if row["npc"]))
            evidence = rows[0]["raw_line"] or f"{rows[0]['actor']} -> {rows[0]['npc']}"
            return SolvedQuestion(answer=answer, evidence=evidence, confidence=float(rows[0]["source_confidence"] or 0.8), source="role_mappings")
    if actor and db_table_exists(conn, "store_castings"):
        scope, params = sql_scope(script_ids)
        store_sql = " AND store_id=? " if store_id else ""
        rows = conn.execute(
            f"""
            SELECT actor_name, character_name, raw_line, source_confidence
            FROM store_castings
            WHERE actor_name LIKE ? {scope} {store_sql}
            ORDER BY source_confidence DESC, last_seen_at DESC
            LIMIT 5
            """,
            [f"%{actor}%"] + params + ([store_id] if store_id else []),
        ).fetchall()
        if rows:
            answer = "、".join(dict.fromkeys(row["character_name"] for row in rows if row["character_name"]))
            evidence = rows[0]["raw_line"] or f"{rows[0]['actor_name']} -> {rows[0]['character_name']}"
            return SolvedQuestion(answer=answer, evidence=evidence, confidence=float(rows[0]["source_confidence"] or 0.8), source="store_castings")
    if reverse and db_table_exists(conn, "store_castings"):
        scope, params = sql_scope(script_ids)
        store_sql = " AND store_id=? " if store_id else ""
        rows = conn.execute(
            f"""
            SELECT actor_name, character_name, raw_line, source_confidence
            FROM store_castings
            WHERE character_name LIKE ? {scope} {store_sql}
            ORDER BY source_confidence DESC, last_seen_at DESC
            LIMIT 5
            """,
            [f"%{reverse}%"] + params + ([store_id] if store_id else []),
        ).fetchall()
        if rows:
            answer = "、".join(dict.fromkeys(row["actor_name"] for row in rows if row["actor_name"]))
            evidence = rows[0]["raw_line"] or f"{rows[0]['character_name']} <- {rows[0]['actor_name']}"
            return SolvedQuestion(answer=answer, evidence=evidence, confidence=float(rows[0]["source_confidence"] or 0.8), source="store_castings")
    return None


def solve_profile_lookup(conn: sqlite3.Connection, question: str, script_ids: list[int]) -> Optional[SolvedQuestion]:
    if not script_ids or not db_table_exists(conn, "script_profiles"):
        return None
    key_map = [
        ("作者", "author"),
        ("分类", "category"),
        ("难度", "difficulty"),
        ("时长", "duration_text"),
        ("简介", "summary"),
        ("发行", "publisher"),
        ("适龄", "age_rating"),
        ("反串", "cross_gender_allowed"),
        ("人数", "player_config_text"),
        ("几人", "player_config_text"),
    ]
    column = None
    label = None
    for needle, col in key_map:
        if needle in question:
            label = needle
            column = col
            break
    if not column:
        return None
    placeholders = ",".join("?" for _ in script_ids)
    row = conn.execute(
        f"""
        SELECT s.title, p.*
        FROM script_profiles p
        JOIN scripts s ON s.id = p.script_id
        WHERE p.script_id IN ({placeholders})
        ORDER BY p.last_seen_at DESC
        LIMIT 1
        """,
        script_ids,
    ).fetchone()
    if not row:
        return None
    value = row[column]
    if column == "cross_gender_allowed":
        value = "可反串" if value == 1 else ("不可反串" if value == 0 else "")
    if value is None or str(value).strip() == "":
        return None
    return SolvedQuestion(answer=str(value), evidence=f"{row['title']} 的 {label}", confidence=0.78, source="script_profiles")


def solve_enumeration(conn: sqlite3.Connection, question: str, script_ids: list[int]) -> Optional[SolvedQuestion]:
    if not script_ids or not db_table_exists(conn, "script_roles"):
        return None
    if not any(key in question for key in ["哪些", "所有", "角色", "npc", "NPC", "DM", "dm"]):
        return None
    filters = []
    if "NPC" in question or "npc" in question or "陪伴" in question:
        filters.append("(is_npc=1 OR is_companion_npc=1 OR role_kind LIKE '%npc%')")
    if "DM" in question or "dm" in question:
        filters.append("(is_dm_role=1 OR role_kind='dm')")
    where_kind = " AND (" + " OR ".join(filters) + ")" if filters else ""
    placeholders = ",".join("?" for _ in script_ids)
    rows = conn.execute(
        f"""
        SELECT canonical_name, raw_text, source_confidence
        FROM script_roles
        WHERE script_id IN ({placeholders}) {where_kind}
        ORDER BY role_kind, canonical_name
        LIMIT 30
        """,
        script_ids,
    ).fetchall()
    if not rows:
        return None
    names = [row["canonical_name"] for row in rows if row["canonical_name"]]
    return SolvedQuestion(answer="、".join(dict.fromkeys(names)), evidence=f"script_roles 命中 {len(rows)} 条", confidence=0.72, source="script_roles")


def solve_count(conn: sqlite3.Connection, question: str, script_ids: list[int], store_id: Optional[int]) -> Optional[SolvedQuestion]:
    if not script_ids or not any(key in question for key in ["几个", "几位", "多少", "数量", "人数"]):
        return None
    if store_id and db_table_exists(conn, "store_offerings"):
        placeholders = ",".join("?" for _ in script_ids)
        row = conn.execute(
            f"""
            SELECT dm_count, npc_count, config_text
            FROM store_offerings
            WHERE script_id IN ({placeholders}) AND store_id=? AND is_active=1
            ORDER BY source_confidence DESC, last_seen_at DESC
            LIMIT 1
            """,
            script_ids + [store_id],
        ).fetchone()
        if row:
            if ("NPC" in question or "npc" in question) and row["npc_count"] is not None:
                return SolvedQuestion(answer=str(row["npc_count"]), evidence=row["config_text"] or "store_offerings.npc_count", confidence=0.82, source="store_offerings")
            if ("DM" in question or "dm" in question) and row["dm_count"] is not None:
                return SolvedQuestion(answer=str(row["dm_count"]), evidence=row["config_text"] or "store_offerings.dm_count", confidence=0.82, source="store_offerings")
    if db_table_exists(conn, "script_roles"):
        placeholders = ",".join("?" for _ in script_ids)
        if "NPC" in question or "npc" in question:
            row = conn.execute(
                f"SELECT COUNT(*) AS c FROM script_roles WHERE script_id IN ({placeholders}) AND (is_npc=1 OR is_companion_npc=1 OR role_kind LIKE '%npc%')",
                script_ids,
            ).fetchone()
            if row and row["c"]:
                return SolvedQuestion(answer=str(row["c"]), evidence="script_roles NPC count", confidence=0.7, source="script_roles")
        if "DM" in question or "dm" in question:
            row = conn.execute(
                f"SELECT COUNT(*) AS c FROM script_roles WHERE script_id IN ({placeholders}) AND (is_dm_role=1 OR role_kind='dm')",
                script_ids,
            ).fetchone()
            if row and row["c"]:
                return SolvedQuestion(answer=str(row["c"]), evidence="script_roles DM count", confidence=0.7, source="script_roles")
    return None


def solve_fts_fallback(conn: sqlite3.Connection, question: str) -> Optional[SolvedQuestion]:
    if not db_table_exists(conn, "kb_fts"):
        return None
    query = re.sub(r"[\"'():*^~-]+", " ", question).strip()
    if not query:
        return None
    try:
        rows = conn.execute(
            """
            SELECT content, source_table, source_id
            FROM kb_fts
            WHERE kb_fts MATCH ?
            LIMIT 5
            """,
            (query,),
        ).fetchall()
    except sqlite3.Error:
        rows = []
    if not rows:
        return None
    content = rows[0]["content"] or ""
    answer = content.strip().splitlines()[0][:80]
    return SolvedQuestion(answer=answer, evidence=f"FTS: {rows[0]['source_table']}#{rows[0]['source_id']}", confidence=0.45, source="kb_fts")


def solve_question_from_context(question: str, preferred_store: Optional[str] = None, context: str = "") -> Optional[SolvedQuestion]:
    with open_knowledge_db(create=True) as conn:
        store_id = resolve_store_id(conn, preferred_store)
        script_ids = guess_script_ids(conn, question, context)
        for solver in [
            lambda: solve_exact_qa(conn, question, script_ids),
            lambda: solve_actor_mapping(conn, question, script_ids, store_id),
            lambda: solve_profile_lookup(conn, question, script_ids),
            lambda: solve_count(conn, question, script_ids, store_id),
            lambda: solve_enumeration(conn, question, script_ids),
            lambda: solve_fts_fallback(conn, question),
        ]:
            solved = solver()
            if solved and solved.answer.strip():
                return solved
    return None


def rebuild_knowledge_index() -> None:
    with open_knowledge_db(create=True) as conn:
        if not db_table_exists(conn, "kb_fts"):
            print("当前 SQLite 不支持 FTS5，无法重建索引")
            return
        conn.execute("DELETE FROM kb_fts")
        sources = [
            ("qa_pairs", "id", "question || '\n' || answer || '\n' || COALESCE(evidence,'')"),
            ("scripts", "id", "title || '\n' || COALESCE(host_name,'') || '\n' || COALESCE(variant,'')"),
            ("script_aliases", "id", "alias"),
            ("facts", "id", "key || '\n' || value || '\n' || COALESCE(raw_line,'')"),
            ("script_roles", "id", "canonical_name || '\n' || COALESCE(raw_text,'') || '\n' || COALESCE(summary,'')"),
            ("role_mappings", "id", "actor || '\n' || npc || '\n' || COALESCE(raw_line,'')"),
            ("store_offerings", "id", "price_text || '\n' || config_text || '\n' || notes"),
            ("store_castings", "id", "character_name || '\n' || actor_name || '\n' || COALESCE(raw_line,'')"),
            ("posts", "id", "raw_text"),
        ]
        inserted = 0
        for table, id_col, expr in sources:
            if not db_table_exists(conn, table):
                continue
            rows = conn.execute(f"SELECT {id_col} AS id, {expr} AS content FROM {table}").fetchall()
            for row in rows:
                content = (row["content"] or "").strip()
                if not content:
                    continue
                conn.execute("INSERT INTO kb_fts(content, source_table, source_id) VALUES (?, ?, ?)", (content, table, row["id"]))
                inserted += 1
        conn.commit()
        print(f"知识库索引已重建: {inserted} 条")


def print_knowledge_stats() -> None:
    with open_knowledge_db(create=True) as conn:
        print(f"知识库: {CONFIG_KB}")
        print(f"schema version: {conn.execute('PRAGMA user_version').fetchone()[0]}")
        for table in [
            "stores",
            "scripts",
            "posts",
            "qa_pairs",
            "facts",
            "script_roles",
            "role_mappings",
            "store_offerings",
            "store_castings",
            "kb_fts",
        ]:
            if db_table_exists(conn, table):
                try:
                    count = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
                    print(f"  {table}: {count}")
                except sqlite3.Error as exc:
                    print(f"  {table}: 读取失败 {exc}")


def search_knowledge(keyword: str) -> None:
    with open_knowledge_db(create=True) as conn:
        rows: list[sqlite3.Row] = []
        if db_table_exists(conn, "kb_fts"):
            try:
                rows = conn.execute(
                    "SELECT content, source_table, source_id FROM kb_fts WHERE kb_fts MATCH ? LIMIT 20",
                    (re.sub(r"[\"'():*^~-]+", " ", keyword).strip(),),
                ).fetchall()
            except sqlite3.Error:
                rows = []
        if not rows:
            like = f"%{keyword}%"
            parts: list[tuple[str, int, str]] = []
            for table, id_col, col in [
                ("qa_pairs", "id", "question || '\n' || answer"),
                ("scripts", "id", "title"),
                ("facts", "id", "key || '\n' || value || '\n' || COALESCE(raw_line,'')"),
                ("posts", "id", "raw_text"),
            ]:
                if not db_table_exists(conn, table):
                    continue
                for row in conn.execute(f"SELECT {id_col} AS id, {col} AS content FROM {table} WHERE {col} LIKE ? LIMIT 8", (like,)):
                    parts.append((table, row["id"], row["content"] or ""))
            for table, row_id, content in parts[:20]:
                print(f"[{table}#{row_id}] {content[:160].replace(chr(10), ' / ')}")
            return
        for row in rows:
            print(f"[{row['source_table']}#{row['source_id']}] {(row['content'] or '')[:160].replace(chr(10), ' / ')}")


def print_parsed_question(question: str, preferred_store: Optional[str]) -> None:
    with open_knowledge_db(create=True) as conn:
        script_ids = guess_script_ids(conn, question)
        store_id = resolve_store_id(conn, preferred_store)
        print(f"问题: {question}")
        print(f"normalized: {normalize_text(question)}")
        print(f"store: {preferred_store or '-'} -> {store_id or '-'}")
        if script_ids:
            placeholders = ",".join("?" for _ in script_ids)
            rows = conn.execute(f"SELECT id, title FROM scripts WHERE id IN ({placeholders})", script_ids).fetchall()
            print("候选剧本:")
            for row in rows:
                print(f"  - #{row['id']} {row['title']}")
        else:
            print("候选剧本: -")
        intents = []
        if any(key in question for key in ["演谁", "谁演", "饰演", "扮演"]):
            intents.append("actor_mapping")
        if any(key in question for key in ["几个", "几位", "多少", "数量"]):
            intents.append("count")
        if any(key in question for key in ["哪些", "所有", "角色"]):
            intents.append("enumeration")
        if any(key in question for key in ["作者", "分类", "难度", "时长", "简介", "发行", "反串"]):
            intents.append("profile_lookup")
        print(f"意图: {', '.join(intents) if intents else 'unknown'}")


def learn_knowledge_from_clipboard() -> None:
    pyperclip = require_module("pyperclip")
    text = (pyperclip.paste() or "").strip()
    if not text:
        print("剪贴板为空")
        return
    signature = hashlib.sha256(text.encode("utf-8")).hexdigest()
    with open_knowledge_db(create=True) as conn:
        row = conn.execute("SELECT id FROM posts WHERE content_signature=? LIMIT 1", (signature,)).fetchone()
        if row:
            print(f"剪贴板内容已存在 posts#{row['id']}")
            return
        cur = conn.execute(
            "INSERT INTO posts(raw_text, source_type, content_signature) VALUES (?, 'clipboard', ?)",
            (text, signature),
        )
        conn.commit()
        print(f"已写入 posts#{cur.lastrowid}")
    rebuild_knowledge_index()


def print_knowledge_history(title: str) -> None:
    with open_knowledge_db(create=True) as conn:
        rows = conn.execute("SELECT id, title, host_name, variant, created_at FROM scripts WHERE title LIKE ? ORDER BY created_at DESC LIMIT 20", (f"%{title}%",)).fetchall()
        if not rows:
            print(f"未找到剧本: {title}")
            return
        for row in rows:
            print(f"script#{row['id']} {row['title']} host={row['host_name'] or '-'} variant={row['variant'] or '-'}")
            for fact in conn.execute("SELECT key, value, raw_line FROM facts WHERE script_id=? ORDER BY last_seen_at DESC LIMIT 8", (row["id"],)):
                print(f"  fact: {fact['key']} = {fact['value']}")


def set_character_gender(script_title: str, character_name: str, gender: str) -> bool:
    if gender not in {"男", "女"}:
        return False
    with open_knowledge_db(create=True) as conn:
        script = conn.execute("SELECT id FROM scripts WHERE title LIKE ? ORDER BY created_at DESC LIMIT 1", (f"%{script_title}%",)).fetchone()
        if not script:
            return False
        script_id = script["id"]
        changed = 0
        if db_table_exists(conn, "script_characters"):
            cur = conn.execute(
                "UPDATE script_characters SET gender_tag=?, last_seen_at=strftime('%s','now') WHERE script_id=? AND character_name LIKE ?",
                (gender, script_id, f"%{character_name}%"),
            )
            changed += cur.rowcount
        if db_table_exists(conn, "script_roles"):
            cur = conn.execute(
                "UPDATE script_roles SET gender_tag=?, last_seen_at=strftime('%s','now') WHERE script_id=? AND canonical_name LIKE ?",
                (gender, script_id, f"%{character_name}%"),
            )
            changed += cur.rowcount
        conn.commit()
        return changed > 0


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
    return capture.screenshot(region)


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
  python easy_money_win.py locate
  python easy_money_win.py run [--interval N] [--pos x,y] [--index N] [--title 文本] [--id id]
  python easy_money_win.py uia-dump [--max-depth N] [--buttons-only]
  python easy_money_win.py ax-dump [--max-depth N] [--buttons-only]
  python easy_money_win.py comment-locate
  python easy_money_win.py comment-fixed-send-locate
  python easy_money_win.py avatar-template-locate [--name 用户] [--set-default]
  python easy_money_win.py avatar-center-locate
  python easy_money_win.py post-image-locate
  python easy_money_win.py post-image-x-locate
  python easy_money_win.py avatar-locate
  python easy_money_win.py user add <name> --template <path> [--match-mode center_square|wide] [--threshold 0.72]
  python easy_money_win.py user remove <name>
  python easy_money_win.py user list
  python easy_money_win.py user default <name>
  python easy_money_win.py comment [--text 文本] [--solve-question|--doubao|--LLM [--vision]] [--noLocal] [--store 商家] --user [name] [--debug]
  python easy_money_win.py llm ask "<问题>" [上下文]
  python easy_money_win.py doubao ask "<朋友圈正文>"
  python easy_money_win.py kb stats|search|ask|parse|rebuild|learn|history|set-gender ...
"""
    )


def parse_option_value(args: list[str], index: int, name: str) -> tuple[str, int]:
    if index + 1 >= len(args) or args[index + 1].startswith("--"):
        raise EasyMoneyError(f"{name} 需要提供值")
    return args[index + 1], index + 1


def cmd_uia_dump(args: list[str]) -> int:
    max_depth = 10
    buttons_only = False
    i = 0
    while i < len(args):
        if args[i] == "--max-depth":
            value, i = parse_option_value(args, i, "--max-depth")
            max_depth = max(1, min(int(value), 30))
        elif args[i] == "--buttons-only":
            buttons_only = True
        i += 1
    backend = WindowBackend()
    win = backend.moments_window()
    backend.dump_tree(win, max_depth=max_depth, buttons_only=buttons_only)
    return 0


def cmd_locate(args: list[str]) -> int:
    no_click = "--no-click" in args
    backend = WindowBackend()
    win_rect = backend.moments_window_rect()
    input_backend = InputBackend()
    print(f"已找到朋友圈窗口: {win_rect.describe()}")
    print("请将鼠标移动到刷新按钮上，3 秒后记录。")
    countdown()
    pos = input_backend.position()
    offset = Point(pos.x - win_rect.left, pos.y - win_rect.top)
    save_point(CONFIG_REFRESH_OFFSET, offset)
    print(f"鼠标位置: ({int(pos.x)}, {int(pos.y)})")
    print(f"相对窗口偏移: ({int(offset.x)}, {int(offset.y)})")
    print(f"已保存到: {CONFIG_REFRESH_OFFSET}")
    if not no_click:
        print("测试点击刷新按钮...")
        input_backend.click(pos)
    print("之后可运行: python easy_money_win.py run --interval 15")
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
    saved_offset = load_point(CONFIG_REFRESH_OFFSET)
    if not any([target_pos, target_index is not None, target_title, target_id]) and saved_offset is None:
        raise EasyMoneyError(f"未找到刷新按钮偏移，请先运行 locate: {CONFIG_REFRESH_OFFSET}")
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
                elif backend.click_control(chosen, input_backend):
                    print(f"[{now}] 已点击按钮 {backend._safe_text(chosen) or target_index}")
                else:
                    print(f"[{now}] 点击按钮失败")
            else:
                win_rect = backend.moments_window_rect()
                assert saved_offset is not None
                point = Point(win_rect.left + saved_offset.x, win_rect.top + saved_offset.y)
                input_backend.click(point)
                print(f"[{now}] 已点击 ({int(point.x)}, {int(point.y)})")
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


def cmd_avatar_template_locate(args: list[str]) -> int:
    name: Optional[str] = None
    set_default = False
    i = 0
    while i < len(args):
        if args[i] == "--name":
            name, i = parse_option_value(args, i, "--name")
        elif args[i] == "--set-default":
            set_default = True
        i += 1
    backend = WindowBackend()
    win_rect = backend.moments_window_rect()
    input_backend = InputBackend()
    user_label = f"（用户: {name}）" if name else ""
    print(f"标定目标用户头像模板{user_label}")
    print("请将鼠标移到目标用户头像中心。")
    countdown()
    pos = input_backend.position()
    if name:
        path = named_template_path(name)
        save_avatar_template_at(pos, win_rect, path)
        config = load_user_template_config()
        old = config.users.get(name)
        config.users[name] = UserTemplateEntry(template=str(path), matchMode=(old.matchMode if old else "center_square"), threshold=(old.threshold if old else None))
        if set_default or config.default is None:
            config.default = name
        save_user_template_config(config)
        print(f"头像模板已保存: {path}")
        if config.default == name:
            print("已设为默认模板")
    else:
        save_avatar_template_at(pos, win_rect, LEGACY_AVATAR_TEMPLATE)
        print(f"头像模板已保存: {LEGACY_AVATAR_TEMPLATE}")
    return 0


def cmd_avatar_center_locate(args: list[str]) -> int:
    backend = WindowBackend()
    win_rect = backend.moments_window_rect()
    input_backend = InputBackend()
    print("请将鼠标移到目标用户头像中心，3 秒后记录。")
    countdown()
    pos = input_backend.position()
    offset = Point(pos.x - win_rect.left, pos.y - win_rect.top)
    save_point(CONFIG_AVATAR_OFFSET, offset)
    print(f"头像中心偏移已保存: ({int(offset.x)}, {int(offset.y)}) -> {CONFIG_AVATAR_OFFSET}")
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


def cmd_avatar_locate(args: list[str]) -> int:
    backend = WindowBackend()
    win_rect = backend.moments_window_rect()
    input_backend = InputBackend()
    print("兼容模式：一次性标定头像中心和 legacy 模板。请将鼠标移到头像中心。")
    countdown()
    pos = input_backend.position()
    save_point(CONFIG_AVATAR_OFFSET, Point(pos.x - win_rect.left, pos.y - win_rect.top))
    save_avatar_template_at(pos, win_rect, LEGACY_AVATAR_TEMPLATE)
    print(f"头像偏移已保存: {CONFIG_AVATAR_OFFSET}")
    print(f"头像模板已保存: {LEGACY_AVATAR_TEMPLATE}")
    return 0


def cmd_avatar_debug(args: list[str]) -> int:
    backend = WindowBackend()
    win_rect = backend.moments_window_rect()
    capture = CaptureBackend()
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    path = DEBUG_DIR / f"avatar_debug_window_{stamp}.png"
    capture.save(capture.screenshot(win_rect), path)
    print(f"已保存窗口截图: {path}")
    left = Rect(win_rect.left, win_rect.top, win_rect.left + win_rect.width * 0.28, win_rect.top + win_rect.height * 0.78)
    path = DEBUG_DIR / f"avatar_debug_left_{stamp}.png"
    capture.save(capture.screenshot(left), path)
    print(f"已保存头像搜索区: {path}")
    return 0


def cmd_user(args: list[str]) -> int:
    if not args:
        raise EasyMoneyError("用法: user <add|remove|list|default> ...")
    sub = args[0]
    if sub == "add":
        if len(args) < 2:
            raise EasyMoneyError("用法: user add <name> --template <path> [--match-mode center_square|wide] [--threshold 0.72]")
        name = args[1]
        template: Optional[str] = None
        match_mode: Optional[str] = None
        threshold: Optional[float] = None
        i = 2
        while i < len(args):
            if args[i] == "--template":
                template, i = parse_option_value(args, i, "--template")
            elif args[i] == "--match-mode":
                value, i = parse_option_value(args, i, "--match-mode")
                match_mode = normalized_avatar_match_mode(value)
                if not match_mode:
                    raise EasyMoneyError("不支持的 match-mode，可用: center_square|wide")
            elif args[i] == "--threshold":
                value, i = parse_option_value(args, i, "--threshold")
                threshold = float(value)
            i += 1
        path = expand_path(template)
        if path is None or not path.exists():
            raise EasyMoneyError(f"模板文件不存在: {template}")
        config = load_user_template_config()
        config.users[name] = UserTemplateEntry(template=str(path), matchMode=match_mode, threshold=threshold)
        if config.default is None:
            config.default = name
        save_user_template_config(config)
        print(f"已保存用户模板: {name} -> {path}")
        return 0
    if sub == "remove":
        if len(args) < 2:
            raise EasyMoneyError("用法: user remove <name>")
        config = load_user_template_config()
        if args[1] not in config.users:
            raise EasyMoneyError(f"未找到用户模板: {args[1]}")
        del config.users[args[1]]
        if config.default == args[1]:
            config.default = sorted(config.users.keys())[0] if config.users else None
        save_user_template_config(config)
        print(f"已删除用户模板: {args[1]}")
        return 0
    if sub == "list":
        config = load_user_template_config()
        merged = discover_user_photo_templates()
        merged.update(config.users)
        if not merged:
            print("当前没有可用的用户模板")
            return 0
        print("用户模板列表:")
        for name in sorted(merged):
            entry = merged[name]
            marker = " (default)" if config.default == name else ""
            print(f"  - {name}{marker}")
            print(f"      template: {entry.template}")
            print(f"      matchMode: {entry.matchMode or 'center_square'}, threshold: {entry.threshold if entry.threshold is not None else '-'}")
        return 0
    if sub == "default":
        if len(args) < 2:
            raise EasyMoneyError("用法: user default <name>")
        name = args[1]
        config = load_user_template_config()
        if name not in config.users:
            auto = discover_user_photo_templates().get(name)
            if auto:
                config.users[name] = auto
            else:
                raise EasyMoneyError(f"未找到用户模板: {name}")
        config.default = name
        save_user_template_config(config)
        print(f"已设置默认模板: {name}")
        return 0
    raise EasyMoneyError(f"未知 user 子命令: {sub}")


def cmd_comment(args: list[str]) -> int:
    comment_text: Optional[str] = None
    user_filter = False
    user_name: Optional[str] = None
    match_mode_override: Optional[str] = None
    solve_question = False
    use_doubao = False
    use_llm = False
    use_vision = False
    no_local = False
    preferred_store: Optional[str] = None
    debug = False
    save_post_image = False
    save_post_image_raw = False
    save_path: Optional[Path] = None
    click_post_image = False
    test_image_crop = False
    rounds = 1

    i = 0
    while i < len(args):
        arg = args[i]
        if arg == "--text":
            comment_text, i = parse_option_value(args, i, "--text")
        elif arg == "--user":
            user_filter = True
            if i + 1 < len(args) and not args[i + 1].startswith("--"):
                user_name = args[i + 1]
                i += 1
        elif arg in {"--user-square"}:
            user_filter = True
            match_mode_override = "center_square"
        elif arg in {"--user-wide", "--avatar-wide"}:
            user_filter = True
            match_mode_override = "wide"
        elif arg == "--match-mode":
            value, i = parse_option_value(args, i, "--match-mode")
            match_mode_override = normalized_avatar_match_mode(value)
            if not match_mode_override:
                raise EasyMoneyError("不支持的 match-mode，可用: center_square|wide")
        elif arg in {"--solve-question", "--slove-question"}:
            solve_question = True
        elif arg == "--doubao":
            use_doubao = True
            solve_question = True
        elif arg == "--LLM":
            use_llm = True
            solve_question = True
        elif arg == "--vision":
            use_vision = True
        elif arg == "--noLocal":
            no_local = True
        elif arg == "--store":
            preferred_store, i = parse_option_value(args, i, "--store")
        elif arg == "--debug":
            debug = True
        elif arg == "--save-post-image":
            save_post_image = True
        elif arg == "--raw":
            save_post_image_raw = True
        elif arg == "--output":
            value, i = parse_option_value(args, i, "--output")
            save_path = expand_path(value)
        elif arg == "--click-post-image":
            click_post_image = True
        elif arg in {"--test-image-crop", "--debug-image-crop"}:
            test_image_crop = True
        elif arg == "--rounds":
            value, i = parse_option_value(args, i, "--rounds")
            rounds = max(1, int(value))
        elif arg in {"--ocr-comment", "--fast", "--stream-capture", "--yolo-debug", "--save-yolo-images"}:
            print(f"提示: Windows v1 暂不完整支持 {arg}，已忽略或降级")
        i += 1

    if not user_filter:
        raise EasyMoneyError("comment 命令必须显式指定 --user [name]")
    if use_vision and not use_llm:
        raise EasyMoneyError("--vision 需要与 --LLM 一起使用")
    if not any([comment_text, solve_question, save_post_image, click_post_image, test_image_crop]):
        raise EasyMoneyError('请指定 --text "评论内容"，或使用 --solve-question / --doubao / --LLM / --save-post-image')

    config = load_comment_config()
    if not config and not any([save_post_image, click_post_image, test_image_crop]):
        raise EasyMoneyError("未找到评论配置，请先运行 comment-locate")

    resolved_name, entry, template_path = resolve_user_template(user_name)
    if template_path is None or not template_path.exists():
        available = ", ".join(sorted(load_user_template_config().users.keys()))
        raise EasyMoneyError(f"未找到用户头像模板: {user_name or '(default)'}" + (f"；可用: {available}" if available else ""))
    match_mode = match_mode_override or normalized_avatar_match_mode(entry.matchMode if entry else None) or "center_square"
    threshold = entry.threshold if entry and entry.threshold is not None else 0.72

    backend = WindowBackend()
    input_backend = InputBackend()
    win = backend.moments_window()
    backend.activate(win)
    window_rect = backend.rect(win)
    if window_rect is None:
        raise EasyMoneyError("无法读取朋友圈窗口位置")

    avatar_center: Optional[Point] = None
    avatar_score = 0.0
    last_error: Optional[Exception] = None
    for _ in range(rounds):
        try:
            avatar_center, avatar_score = find_avatar_in_window(window_rect, template_path, threshold, match_mode)
            break
        except Exception as exc:
            last_error = exc
            refresh_offset = load_point(CONFIG_REFRESH_OFFSET)
            if refresh_offset:
                input_backend.click(Point(window_rect.left + refresh_offset.x, window_rect.top + refresh_offset.y))
                time.sleep(0.6)
    if avatar_center is None:
        raise EasyMoneyError(str(last_error or "头像匹配失败"))

    need_text = solve_question
    post = resolve_moment_post(backend, win, window_rect, avatar_center, require_text=need_text)
    print(f"已匹配用户: {resolved_name or user_name or 'legacy'} score={avatar_score:.3f}")
    print(f"头像中心: ({int(avatar_center.x)}, {int(avatar_center.y)})")
    print(f"动态定位: {post.source} frame={post.body_frame.describe()}")
    if post.text:
        print("正文内容开始")
        print(post.text)
        print("正文内容结束")

    if save_post_image or test_image_crop:
        image = capture_post_image(post, window_rect)
        output = save_path or DEBUG_DIR / f"wechat_post_image_{time.strftime('%Y%m%d_%H%M%S')}.png"
        CaptureBackend().save(image, output)
        print(f"动态图片/区域已保存: {output}")
        return 0
    if click_post_image:
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

    final_text = (comment_text or "").strip()
    image_urls: list[str] = []
    if solve_question:
        context = post.text.strip()
        if not context:
            raise EasyMoneyError("需要自动答题但未能读取朋友圈正文")
        solved: Optional[SolvedQuestion] = None
        if not no_local and not use_llm:
            solved = solve_question_from_context(context, preferred_store=preferred_store, context=context)
        if solved:
            final_text = solved.answer
            print(f"本地知识库命中: {solved.answer} (source={solved.source}, confidence={solved.confidence:.2f})")
        elif use_doubao or use_llm:
            if use_vision:
                image_urls = capture_yolo_image_data_urls(post, window_rect)
                print(f"已附带 YOLO 图片: {len(image_urls)} 张")
            solved = ask_doubao_to_solve_post(context, image_data_urls=image_urls)
            if solved:
                final_text = solved.answer
                print(f"LLM 命中: {solved.answer}")
        if not final_text and comment_text:
            final_text = comment_text.strip()
            print("自动答题未命中，回退到 --text")
        if not final_text:
            raise EasyMoneyError("未能生成评论内容，请补充 --text 作为回退")

    if config is None:
        raise EasyMoneyError("未找到评论配置，请先运行 comment-locate")
    send_point, send_method = resolve_send_point(post.action_point, window_rect, config)
    if debug:
        input_backend.move_to(post.action_point)
        print(f"DEBUG: 操作按钮点 ({int(post.action_point.x)}, {int(post.action_point.y)})")
        print(f"DEBUG: 发送点 [{send_method}] ({int(send_point.x)}, {int(send_point.y)})")
        print(f"DEBUG: 评论内容: {final_text}")
        return 0

    input_backend.click(post.action_point)
    time.sleep(0.25)
    comment_point = Point(post.action_point.x + config.comment_from_action.x, post.action_point.y + config.comment_from_action.y)
    input_backend.click(comment_point)
    time.sleep(0.25)
    paste_method = input_backend.paste_text(final_text)
    time.sleep(0.12)
    input_backend.click(send_point, clicks=3, interval=0.03)
    print(f"已执行评论发送: {paste_method} | 发送点={send_method} ({int(send_point.x)}, {int(send_point.y)})")
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


def cmd_kb(args: list[str]) -> int:
    if not args:
        raise EasyMoneyError("用法: kb <stats|search|ask|parse|rebuild|learn|history|set-gender> ...")
    sub = args[0]
    if sub == "stats":
        print_knowledge_stats()
        return 0
    if sub == "search":
        if len(args) < 2:
            raise EasyMoneyError('用法: kb search "<关键词>"')
        search_knowledge(" ".join(args[1:]))
        return 0
    if sub == "ask":
        question_parts: list[str] = []
        store: Optional[str] = None
        i = 1
        while i < len(args):
            if args[i] == "--store":
                store, i = parse_option_value(args, i, "--store")
            else:
                question_parts.append(args[i])
            i += 1
        question = " ".join(question_parts).strip()
        if not question:
            raise EasyMoneyError('用法: kb ask "<问题>" [--store 商家名]')
        started = time.perf_counter()
        solved = solve_question_from_context(question, preferred_store=store)
        elapsed = (time.perf_counter() - started) * 1000
        if solved:
            print(f"答案: {solved.answer}")
            if solved.evidence:
                print(f"证据: {solved.evidence}")
            print(f"置信度: {solved.confidence:.2f}")
            print(f"总耗时: {elapsed:.1f}ms")
            return 0
        print(f"总耗时: {elapsed:.1f}ms")
        print("未能回答")
        return 1
    if sub == "parse":
        question_parts = []
        store = None
        i = 1
        while i < len(args):
            if args[i] == "--store":
                store, i = parse_option_value(args, i, "--store")
            else:
                question_parts.append(args[i])
            i += 1
        question = " ".join(question_parts).strip()
        if not question:
            raise EasyMoneyError('用法: kb parse "<问题>" [--store 商家名]')
        print_parsed_question(question, store)
        return 0
    if sub == "rebuild":
        rebuild_knowledge_index()
        return 0
    if sub == "learn":
        learn_knowledge_from_clipboard()
        return 0
    if sub == "history":
        if len(args) < 2:
            raise EasyMoneyError('用法: kb history "<剧本名>"')
        print_knowledge_history(" ".join(args[1:]))
        return 0
    if sub == "set-gender":
        if len(args) < 4:
            raise EasyMoneyError('用法: kb set-gender "<剧本名>" "<角色名>" "<男|女>"')
        if set_character_gender(args[1], args[2], args[3]):
            print(f"已更新角色性别: {args[1]} / {args[2]} -> {args[3]}")
            rebuild_knowledge_index()
            return 0
        print("更新失败：请确认剧本名、角色名存在，且性别为 男 或 女")
        return 1
    if sub in {"audit", "self-test", "regression", "post-regression", "post-ask"}:
        raise EasyMoneyError(f"Windows v1 暂未迁移 kb {sub}，请使用 stats/search/ask/parse/rebuild/learn/history/set-gender")
    raise EasyMoneyError(f"未知 kb 子命令: {sub}")


def main(argv: Optional[list[str]] = None) -> int:
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
        "locate": cmd_locate,
        "run": cmd_run,
        "comment-locate": cmd_comment_locate,
        "comment-fixed-send-locate": cmd_comment_fixed_send_locate,
        "avatar-template-locate": cmd_avatar_template_locate,
        "avatar-center-locate": cmd_avatar_center_locate,
        "post-image-locate": lambda a: cmd_post_image_locate(a, x_only=False),
        "post-image-x-locate": lambda a: cmd_post_image_locate(a, x_only=True),
        "avatar-locate": cmd_avatar_locate,
        "avatar-debug": cmd_avatar_debug,
        "user": cmd_user,
        "comment": cmd_comment,
        "llm": cmd_llm,
        "doubao": cmd_doubao,
        "kb": cmd_kb,
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
