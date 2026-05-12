from __future__ import annotations

import argparse
import ctypes
import datetime as _dt
import sys
import threading
import time
from dataclasses import dataclass
from typing import Any, Iterable, Optional

import comtypes
import comtypes.client
from comtypes import COMError, COMObject
from comtypes.automation import VARIANT

from easy_money_win import EasyMoneyError, Rect, WindowBackend


comtypes.client.GetModule("UIAutomationCore.dll")
from comtypes.gen import UIAutomationClient as UIA  # noqa: E402


PROPERTY_IDS = {
    "name": UIA.UIA_NamePropertyId,
    "automation-id": UIA.UIA_AutomationIdPropertyId,
    "control-type": UIA.UIA_ControlTypePropertyId,
    "class-name": UIA.UIA_ClassNamePropertyId,
    "rect": UIA.UIA_BoundingRectanglePropertyId,
    "offscreen": UIA.UIA_IsOffscreenPropertyId,
}

AUTOMATION_EVENT_IDS = {
    "text": UIA.UIA_Text_TextChangedEventId,
    "async": UIA.UIA_AsyncContentLoadedEventId,
    "layout": UIA.UIA_LayoutInvalidatedEventId,
    "window-opened": UIA.UIA_Window_WindowOpenedEventId,
    "window-closed": UIA.UIA_Window_WindowClosedEventId,
    "menu-opened": UIA.UIA_MenuOpenedEventId,
    "menu-closed": UIA.UIA_MenuClosedEventId,
}

if hasattr(UIA, "UIA_NotificationEventId"):
    AUTOMATION_EVENT_IDS["notification"] = UIA.UIA_NotificationEventId

STRUCTURE_CHANGE_NAMES = {
    getattr(UIA, "StructureChangeType_ChildAdded", 0): "ChildAdded",
    getattr(UIA, "StructureChangeType_ChildRemoved", 1): "ChildRemoved",
    getattr(UIA, "StructureChangeType_ChildrenInvalidated", 2): "ChildrenInvalidated",
    getattr(UIA, "StructureChangeType_ChildrenBulkAdded", 3): "ChildrenBulkAdded",
    getattr(UIA, "StructureChangeType_ChildrenBulkRemoved", 4): "ChildrenBulkRemoved",
    getattr(UIA, "StructureChangeType_ChildrenReordered", 5): "ChildrenReordered",
}


def make_reverse_name_map(suffix: str) -> dict[int, str]:
    result: dict[int, str] = {}
    for name in dir(UIA):
        if not name.startswith("UIA_") or not name.endswith(suffix):
            continue
        value = getattr(UIA, name)
        if isinstance(value, int):
            pretty = name.removeprefix("UIA_").removesuffix(suffix)
            result[value] = pretty
    return result


EVENT_NAMES = make_reverse_name_map("EventId")
PROPERTY_NAMES = make_reverse_name_map("PropertyId")
CONTROL_TYPE_NAMES = make_reverse_name_map("ControlTypeId")


@dataclass(frozen=True)
class ElementSnapshot:
    name: str
    automation_id: str
    control_type: str
    class_name: str
    rect: Optional[Rect]

    def describe(self) -> str:
        parts = [self.control_type or "unknown"]
        if self.automation_id:
            parts.append(f"auto_id={self.automation_id!r}")
        if self.class_name:
            parts.append(f"class={self.class_name!r}")
        if self.name:
            parts.append(f"name={shorten(self.name)!r}")
        if self.rect:
            parts.append(f"rect={self.rect.describe()}")
        return " ".join(parts)


def shorten(value: str, limit: int = 80) -> str:
    value = value.replace("\r", " ").replace("\n", " ").strip()
    if len(value) <= limit:
        return value
    return value[: limit - 3] + "..."


def safe_current(element: Any, attr: str, default: Any = "") -> Any:
    try:
        value = getattr(element, attr)
    except Exception:
        return default
    try:
        return value() if callable(value) else value
    except Exception:
        return default


def rect_from_uia(value: Any) -> Optional[Rect]:
    try:
        left = float(getattr(value, "left"))
        top = float(getattr(value, "top"))
        right = float(getattr(value, "right"))
        bottom = float(getattr(value, "bottom"))
    except Exception:
        return None
    if right <= left or bottom <= top:
        return None
    return Rect(left, top, right, bottom)


def snapshot_element(element: Any) -> ElementSnapshot:
    control_type = safe_current(element, "CurrentControlType", 0)
    control_type_name = CONTROL_TYPE_NAMES.get(control_type, str(control_type or ""))
    return ElementSnapshot(
        name=str(safe_current(element, "CurrentName", "") or ""),
        automation_id=str(safe_current(element, "CurrentAutomationId", "") or ""),
        control_type=control_type_name,
        class_name=str(safe_current(element, "CurrentClassName", "") or ""),
        rect=rect_from_uia(safe_current(element, "CurrentBoundingRectangle", None)),
    )


def rect_intersects(lhs: Rect, rhs: Rect) -> bool:
    return lhs.left < rhs.right and lhs.right > rhs.left and lhs.top < rhs.bottom and lhs.bottom > rhs.top


def format_variant(value: Any) -> str:
    if hasattr(value, "value"):
        try:
            value = value.value
        except Exception:
            pass
    rect = rect_from_uia(value)
    if rect is not None:
        return rect.describe()
    if isinstance(value, str):
        return repr(shorten(value))
    return repr(value)


def format_runtime_id(runtime_id: Any) -> str:
    try:
        values = [int(item) for item in runtime_id]
    except Exception:
        return ""
    if not values:
        return ""
    return " runtime_id=" + ".".join(str(item) for item in values[:8])


class EventLogger:
    def __init__(self, root_rect: Rect, max_events: int = 0, filter_to_root: bool = True) -> None:
        self.root_rect = root_rect
        self.max_events = max(0, max_events)
        self.filter_to_root = filter_to_root
        self.count = 0
        self.started = time.perf_counter()
        self.stop_requested = False
        self.lock = threading.Lock()

    def _should_log(self, element: Any) -> tuple[bool, ElementSnapshot]:
        snap = snapshot_element(element)
        if self.filter_to_root and snap.rect is not None and not rect_intersects(snap.rect, self.root_rect):
            return False, snap
        return True, snap

    def log(self, kind: str, element: Any, detail: str = "") -> None:
        should_log, snap = self._should_log(element)
        if not should_log:
            return
        with self.lock:
            self.count += 1
            elapsed_ms = int((time.perf_counter() - self.started) * 1000)
            now = _dt.datetime.now().strftime("%H:%M:%S.%f")[:-3]
            suffix = f" {detail}" if detail else ""
            print(f"{now} +{elapsed_ms:06d}ms #{self.count:04d} {kind}{suffix} | {snap.describe()}", flush=True)
            if self.max_events and self.count >= self.max_events:
                self.stop_requested = True


class AutomationEventHandler(COMObject):
    _com_interfaces_ = [UIA.IUIAutomationEventHandler]

    def __init__(self, logger: EventLogger) -> None:
        super().__init__()
        self.logger = logger

    def HandleAutomationEvent(self, sender: Any, eventId: int) -> int:
        event_name = EVENT_NAMES.get(eventId, str(eventId))
        self.logger.log("EVENT", sender, event_name)
        return 0


class StructureChangedEventHandler(COMObject):
    _com_interfaces_ = [UIA.IUIAutomationStructureChangedEventHandler]

    def __init__(self, logger: EventLogger) -> None:
        super().__init__()
        self.logger = logger

    def HandleStructureChangedEvent(self, sender: Any, changeType: int, runtimeId: Any) -> int:
        change_name = STRUCTURE_CHANGE_NAMES.get(changeType, str(changeType))
        self.logger.log("STRUCTURE", sender, change_name + format_runtime_id(runtimeId))
        return 0


class PropertyChangedEventHandler(COMObject):
    _com_interfaces_ = [UIA.IUIAutomationPropertyChangedEventHandler]

    def __init__(self, logger: EventLogger) -> None:
        super().__init__()
        self.logger = logger

    def HandlePropertyChangedEvent(self, sender: Any, propertyId: int, newValue: Any) -> int:
        prop_name = PROPERTY_NAMES.get(propertyId, str(propertyId))
        self.logger.log("PROPERTY", sender, f"{prop_name}={format_variant(newValue)}")
        return 0


class FocusChangedEventHandler(COMObject):
    _com_interfaces_ = [UIA.IUIAutomationFocusChangedEventHandler]

    def __init__(self, logger: EventLogger) -> None:
        super().__init__()
        self.logger = logger

    def HandleFocusChangedEvent(self, sender: Any) -> int:
        self.logger.log("FOCUS", sender)
        return 0


def window_handle(win: Any) -> int:
    for attr in ("NativeWindowHandle", "CurrentNativeWindowHandle"):
        try:
            value = getattr(win, attr)
            value = value() if callable(value) else value
            if value:
                return int(value)
        except Exception:
            pass
    for attr in ("handle",):
        value = getattr(win, attr, None)
        if value:
            return int(value() if callable(value) else value)
    element_info = getattr(win, "element_info", None)
    value = getattr(element_info, "handle", None)
    if value:
        return int(value)
    raise EasyMoneyError("无法读取朋友圈窗口句柄")


def find_sns_list_element(automation: Any, root: Any) -> Optional[Any]:
    condition = automation.CreatePropertyCondition(UIA.UIA_AutomationIdPropertyId, VARIANT("sns_list"))
    try:
        found = root.FindFirst(UIA.TreeScope_Subtree, condition)
    except COMError:
        return None
    return found or None


def parse_csv_set(raw: str, allowed: Iterable[str], label: str) -> set[str]:
    allowed_set = set(allowed)
    values = {item.strip().lower() for item in raw.split(",") if item.strip()}
    unknown = sorted(values - allowed_set)
    if unknown:
        raise EasyMoneyError(f"不支持的 {label}: {', '.join(unknown)}；可用: {', '.join(sorted(allowed_set))}")
    return values


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="监听微信朋友圈窗口的 UIAutomation 事件。")
    parser.add_argument("--seconds", type=float, default=30.0, help="监听秒数；0 表示一直监听直到 Ctrl+C。默认 30。")
    parser.add_argument("--root", choices=["window", "sns-list"], default="window", help="监听根节点。默认 window。")
    parser.add_argument("--scope", choices=["element", "children", "descendants", "subtree"], default="subtree", help="UIA TreeScope。默认 subtree。")
    parser.add_argument("--events", default="structure,property,text,focus,async,layout,window-opened,window-closed", help="逗号分隔事件类型。")
    parser.add_argument("--properties", default="name,automation-id,control-type,class-name,rect,offscreen", help="逗号分隔属性监听项。")
    parser.add_argument("--max-events", type=int, default=0, help="收到 N 条事件后自动退出；0 表示不限。")
    parser.add_argument("--no-root-filter", action="store_true", help="不按朋友圈窗口矩形过滤全局 focus 事件。")
    return parser.parse_args(argv)


def scope_value(name: str) -> int:
    return {
        "element": UIA.TreeScope_Element,
        "children": UIA.TreeScope_Children,
        "descendants": UIA.TreeScope_Descendants,
        "subtree": UIA.TreeScope_Subtree,
    }[name]


def register_handlers(args: argparse.Namespace, automation: Any, root_element: Any, logger: EventLogger) -> list[Any]:
    handlers: list[Any] = []
    selected_events = parse_csv_set(args.events, {"structure", "property", "focus", *AUTOMATION_EVENT_IDS.keys()}, "events")
    selected_properties = parse_csv_set(args.properties, PROPERTY_IDS.keys(), "properties")
    scope = scope_value(args.scope)

    if "structure" in selected_events:
        handler = StructureChangedEventHandler(logger)
        automation.AddStructureChangedEventHandler(root_element, scope, None, handler)
        handlers.append(("structure", handler))

    if "property" in selected_events and selected_properties:
        handler = PropertyChangedEventHandler(logger)
        property_values = [PROPERTY_IDS[name] for name in sorted(selected_properties)]
        property_array = (ctypes.c_int * len(property_values))(*property_values)
        automation.AddPropertyChangedEventHandlerNativeArray(
            root_element,
            scope,
            None,
            handler,
            property_array,
            len(property_values),
        )
        handlers.append(("property", handler))

    automation_handler: Optional[AutomationEventHandler] = None
    for event_name, event_id in AUTOMATION_EVENT_IDS.items():
        if event_name not in selected_events:
            continue
        if automation_handler is None:
            automation_handler = AutomationEventHandler(logger)
        automation.AddAutomationEventHandler(event_id, root_element, scope, None, automation_handler)
        handlers.append(("automation", event_id, automation_handler))

    if "focus" in selected_events:
        handler = FocusChangedEventHandler(logger)
        automation.AddFocusChangedEventHandler(None, handler)
        handlers.append(("focus", handler))

    return handlers


def unregister_handlers(automation: Any, root_element: Any, handlers: list[Any]) -> None:
    for item in reversed(handlers):
        try:
            kind = item[0]
            if kind == "structure":
                automation.RemoveStructureChangedEventHandler(root_element, item[1])
            elif kind == "property":
                automation.RemovePropertyChangedEventHandler(root_element, item[1])
            elif kind == "automation":
                automation.RemoveAutomationEventHandler(item[1], root_element, item[2])
            elif kind == "focus":
                automation.RemoveFocusChangedEventHandler(item[1])
        except Exception:
            pass


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    backend = WindowBackend()
    win = backend.moments_window()
    backend.activate(win)
    win_rect = backend.moments_window_rect()
    hwnd = window_handle(win)

    automation = comtypes.client.CreateObject(UIA.CUIAutomation, interface=UIA.IUIAutomation)
    root_element = automation.ElementFromHandle(ctypes.c_void_p(hwnd))
    root_label = "window"
    if args.root == "sns-list":
        sns_list = find_sns_list_element(automation, root_element)
        if sns_list is None:
            raise EasyMoneyError("当前 UIA 树未暴露 sns_list，无法以 sns-list 为根监听；请改用 --root window")
        root_element = sns_list
        root_label = "sns-list"

    root_snapshot = snapshot_element(root_element)
    logger = EventLogger(win_rect, max_events=args.max_events, filter_to_root=not args.no_root_filter)
    print(f"UIA 事件监听启动: root={root_label} scope={args.scope} hwnd={hwnd} window={win_rect.describe()}")
    print(f"根节点: {root_snapshot.describe()}")
    print("操作微信朋友圈或点击刷新按钮，事件会实时输出；按 Ctrl+C 退出。", flush=True)

    handlers = register_handlers(args, automation, root_element, logger)
    if not handlers:
        raise EasyMoneyError("没有注册任何 UIA 事件，请检查 --events 参数")

    deadline = None if args.seconds <= 0 else time.perf_counter() + args.seconds
    heartbeat_at = time.perf_counter() + 5
    try:
        while not logger.stop_requested:
            if deadline is not None and time.perf_counter() >= deadline:
                break
            comtypes.client.PumpEvents(0.05)
            if time.perf_counter() >= heartbeat_at:
                print(f"[alive] 已监听 {int(time.perf_counter() - logger.started)}s，事件数={logger.count}", flush=True)
                heartbeat_at = time.perf_counter() + 5
    except KeyboardInterrupt:
        print("\n收到 Ctrl+C，停止监听。")
    finally:
        unregister_handlers(automation, root_element, handlers)
    print(f"UIA 事件监听结束，事件数={logger.count}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main(sys.argv[1:]))
    except EasyMoneyError as exc:
        print(f"错误: {exc}", file=sys.stderr)
        raise SystemExit(1)
