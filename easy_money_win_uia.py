from __future__ import annotations

from easy_money_win_core import *
from easy_money_win_input import InputBackend

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
        self._list_item_strategy_cache: dict[Any, str] = {}

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

    def _find_moments_window_by_win32_class(self) -> Optional[Any]:
        if os.name != "nt":
            return None
        try:
            hwnd = ctypes.windll.user32.FindWindowW("mmui::SNSWindow", None)
        except Exception:
            return None
        if not hwnd:
            return None
        try:
            spec = self.desktop.window(handle=hwnd)
            wrapper = spec.wrapper_object()
            if wrapper is not None:
                return wrapper
        except Exception:
            pass
        try:
            automation, _ = self._ensure_automation()
            return automation.ElementFromHandle(ctypes.c_void_p(int(hwnd)))
        except Exception:
            return None

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
        fast = self._find_moments_window_by_win32_class()
        if fast is not None:
            return fast

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
    strategies = (
        "listitem_children",
        "children",
        "iter_tree",
        "descendants_listitem",
        "descendants",
    )
    try:
        cache_key = window_backend._control_identity(root)
    except Exception:
        cache_key = ("object", id(root))
    strategy_cache = getattr(window_backend, "_list_item_strategy_cache", None)

    def add_if_list_item(node: Any) -> bool:
        marker = window_backend._control_identity(node)
        if marker in seen:
            return False
        if window_backend._control_type(node).lower() != "listitem":
            return False
        seen.add(marker)
        items.append(node)
        return limit is not None and len(items) >= limit

    def strategy_nodes(strategy: str) -> Iterable[Any]:
        if strategy == "listitem_children":
            yield from window_backend.listitem_children(root, limit=limit)
        elif strategy == "children":
            yield from window_backend.children(root)
        elif strategy == "iter_tree":
            for node, depth in window_backend.iter_tree(root, max_depth=max_depth):
                if depth != 0:
                    yield node
        elif strategy == "descendants_listitem":
            try:
                yield from root.descendants(control_type="ListItem")
            except Exception:
                return
        elif strategy == "descendants":
            try:
                yield from root.descendants()
            except Exception:
                return

    def scan_strategy(strategy: str) -> tuple[bool, bool]:
        start_count = len(items)
        for node in strategy_nodes(strategy):
            if add_if_list_item(node):
                return True, True
        found = len(items) > start_count
        complete = found and (limit is None or len(items) >= limit)
        return found, complete

    cached_strategy = strategy_cache.get(cache_key) if isinstance(strategy_cache, dict) else None
    if cached_strategy in strategies:
        found, complete = scan_strategy(cached_strategy)
        if complete:
            return items

    for strategy in strategies:
        if strategy == cached_strategy:
            continue
        found, complete = scan_strategy(strategy)
        if found and isinstance(strategy_cache, dict):
            strategy_cache[cache_key] = strategy
        if complete:
            return items

    if items and isinstance(strategy_cache, dict):
        strategy_cache[cache_key] = cached_strategy if cached_strategy in strategies else strategies[0]
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
        inline_image_count=extract_inline_image_count(raw_text) if include_text else None,
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

