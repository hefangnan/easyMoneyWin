from __future__ import annotations

from easy_money_win_core import *

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
        self._unicode_text_input_cache: dict[str, tuple[tuple[int, Any], ...]] = {}
        self._mouse_click_cache: dict[int, tuple[INPUT, ...]] = {}
        self._mouse_click_array_cache: dict[int, tuple[int, Any]] = {}
        self._virtual_screen_metrics: Optional[tuple[int, int, int, int]] = None
        self._input_timing_enabled = False
        self._input_timing_context = ""
        self._input_timing_events: list[tuple[str, int]] = []
        if self.user32 is not None:
            try:
                self.user32.SendInput.argtypes = (wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int)
                self.user32.SendInput.restype = wintypes.UINT
            except Exception:
                pass

    def set_input_timing_enabled(self, enabled: bool) -> None:
        self._input_timing_enabled = enabled
        self._input_timing_events = []

    def set_input_timing_context(self, context: str) -> None:
        self._input_timing_context = context

    def input_timings(self) -> tuple[tuple[str, int], ...]:
        return tuple(self._input_timing_events)

    def _input_timing_label(self, action: str) -> str:
        return f"{self._input_timing_context} {action}".strip()

    def _record_input_timing(self, action: str, started_ns: int) -> None:
        if self._input_timing_enabled:
            self._input_timing_events.append((self._input_timing_label(action), time.perf_counter_ns() - started_ns))

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
            "left": 0x25,
            "up": 0x26,
            "right": 0x27,
            "down": 0x28,
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
                started_ns = time.perf_counter_ns()
                sent = self.user32.SendInput(count, event_array, ctypes.sizeof(INPUT))
                self._record_input_timing(f"SendInput(keys {format_key_sequence(key_tuple)})", started_ns)
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

    def prepare_mouse_click(self, clicks: int = 1) -> None:
        if clicks <= 0 or not (self.native and self.user32 is not None):
            return
        if clicks in self._mouse_click_array_cache:
            return
        events = self._mouse_click_input_events(clicks)
        array_type = INPUT * len(events)
        self._mouse_click_array_cache[clicks] = (len(events), array_type(*events))

    def send_prepared_mouse_click(self, clicks: int = 1) -> bool:
        if clicks <= 0 or not (self.native and self.user32 is not None):
            return False
        self.prepare_mouse_click(clicks)
        cached = self._mouse_click_array_cache.get(clicks)
        if cached is None:
            return False
        count, event_array = cached
        started_ns = time.perf_counter_ns()
        sent = self.user32.SendInput(count, event_array, ctypes.sizeof(INPUT))
        self._record_input_timing(f"SendInput(mouse click x{clicks})", started_ns)
        return sent == count

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
        started_ns = time.perf_counter_ns()
        sent = self.user32.SendInput(len(events), event_array, ctypes.sizeof(INPUT))
        self._record_input_timing(f"SendInput(events x{len(events)})", started_ns)
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
            started_ns = time.perf_counter_ns()
            self.user32.SetCursorPos(x, y)
            self._record_input_timing(f"SetCursorPos({x},{y})", started_ns)
            if interval <= 0:
                if self.send_prepared_mouse_click(clicks):
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
            started_ns = time.perf_counter_ns()
            sent = self.user32.SendInput(len(events), event_array, ctypes.sizeof(INPUT))
            self._record_input_timing(f"SendInput(keys {format_key_sequence(key_list)})", started_ns)
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

    def prepare_unicode_text_directly(self, text: str) -> None:
        if not (self.native and self.user32 is not None):
            raise EasyMoneyError("当前平台不支持 Unicode 直接输入")
        if not self.can_type_unicode_directly(text):
            raise EasyMoneyError("当前文本不适合 Unicode 直接输入")
        if text in self._unicode_text_input_cache:
            return
        units = self._utf16_units(text)
        chunk_size = max(1, DIRECT_TEXT_ENTRY_CHUNK_UTF16_UNITS)
        chunks: list[tuple[int, Any]] = []
        for offset in range(0, len(units), chunk_size):
            events: list[INPUT] = []
            for unit in units[offset : offset + chunk_size]:
                events.extend(self._unicode_input_events_for_unit(unit))
            array_type = INPUT * len(events)
            chunks.append((len(events), array_type(*events)))
        self._unicode_text_input_cache[text] = tuple(chunks)

    def prepare_text_input(self, text: str) -> None:
        if self.native and self.user32 is not None and self.can_type_unicode_directly(text):
            self.prepare_unicode_text_directly(text)

    def type_unicode_text_directly(self, text: str) -> str:
        if not (self.native and self.user32 is not None):
            raise EasyMoneyError("当前平台不支持 Unicode 直接输入")
        if not self.can_type_unicode_directly(text):
            raise EasyMoneyError("当前文本不适合 Unicode 直接输入")
        self.prepare_unicode_text_directly(text)
        for count, event_array in self._unicode_text_input_cache[text]:
            started_ns = time.perf_counter_ns()
            sent = self.user32.SendInput(count, event_array, ctypes.sizeof(INPUT))
            self._record_input_timing(f"SendInput(unicode units={count // 2})", started_ns)
            if sent != count:
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
    valid_names = {
        "tab",
        "enter",
        "return",
        "esc",
        "escape",
        "ctrl",
        "control",
        "shift",
        "alt",
        "left",
        "up",
        "right",
        "down",
        "space",
    }
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
        "left": "Left",
        "up": "Up",
        "right": "Right",
        "down": "Down",
        "space": "Space",
        " ": "Space",
    }
    return "+".join(labels.get(key.lower(), key.upper() if len(key) == 1 else key) for key in keys)


