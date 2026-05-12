from __future__ import annotations

import argparse
import ctypes
import datetime as _dt
import sys
import threading
import time
from dataclasses import dataclass
from typing import Any, Optional

import comtypes.client

from easy_money_win import (
    EasyMoneyError,
    Rect,
    WindowBackend,
    clean_post_text,
    find_sns_list_control,
    find_uia_list_items,
)
from uia_event_watch import (
    UIA,
    find_sns_list_element,
    register_handlers,
    shorten,
    snapshot_element,
    unregister_handlers,
    window_handle,
)


@dataclass(frozen=True)
class ListItemSnapshot:
    index: int
    control_type: str
    automation_id: str
    rect: Optional[Rect]
    text: str
    self_text: str
    raw_text: str
    node_count: int

    def signature_part(self) -> str:
        rect = self.rect.describe() if self.rect else "no-rect"
        return f"{self.index}|{self.control_type}|{self.automation_id}|{rect}|{self.text}"


class SnsRefreshTrigger:
    def __init__(self, max_events: int = 0) -> None:
        self.lock = threading.Lock()
        self.pending = True
        self.event_count = 0
        self.max_events = max(0, max_events)
        self.stop_requested = False
        self.last_event_at = 0.0
        self.reasons: list[str] = ["initial"]

    def log(self, kind: str, element: Any, detail: str = "") -> None:
        try:
            snap = snapshot_element(element)
            source = snap.control_type or "unknown"
            if snap.automation_id:
                source += f"/{snap.automation_id}"
        except Exception:
            source = "unknown"
        reason = f"{kind}:{detail or source}"
        with self.lock:
            self.pending = True
            self.event_count += 1
            self.last_event_at = time.perf_counter()
            self.reasons.append(reason)
            if len(self.reasons) > 8:
                self.reasons = self.reasons[-8:]
            if self.max_events and self.event_count >= self.max_events:
                self.stop_requested = True

    def mark_poll(self) -> None:
        with self.lock:
            self.pending = True
            self.last_event_at = time.perf_counter()
            self.reasons.append("poll")
            if len(self.reasons) > 8:
                self.reasons = self.reasons[-8:]

    def ready(self, debounce_seconds: float) -> bool:
        with self.lock:
            if not self.pending:
                return False
            if self.last_event_at and time.perf_counter() - self.last_event_at < debounce_seconds:
                return False
            return True

    def consume_reasons(self) -> list[str]:
        with self.lock:
            self.pending = False
            reasons = self.reasons
            self.reasons = []
            return reasons


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="实时列出微信朋友圈 sns_list 下的 UIA ListItem 内容。")
    parser.add_argument("--seconds", type=float, default=0.0, help="监听秒数；0 表示一直监听直到 Ctrl+C。默认 0。")
    parser.add_argument("--limit", type=int, default=12, help="最多显示多少个 ListItem。默认 12。")
    parser.add_argument("--max-text", type=int, default=160, help="每条内容最多显示多少字符。默认 160。")
    parser.add_argument("--full-text", action="store_true", help="显示完整多行文本，不截断。")
    parser.add_argument("--show-empty", action="store_true", help="显示文本为空的 ListItem。")
    parser.add_argument("--text-depth", type=int, default=5, help="递归读取 ListItem 子控件文本的深度。默认 5。")
    parser.add_argument("--event-root", choices=["window", "sns-list"], default="window", help="UIA 事件监听根节点。默认 window，能覆盖 sns_list 被替换的情况。")
    parser.add_argument("--scope", choices=["element", "children", "descendants", "subtree"], default="subtree", help="UIA TreeScope。默认 subtree。")
    parser.add_argument("--events", default="structure,property,text,async,layout", help="用于触发刷新快照的 UIA 事件，逗号分隔。")
    parser.add_argument("--properties", default="name,automation-id,control-type,rect,offscreen", help="用于触发刷新快照的属性，逗号分隔。")
    parser.add_argument("--debounce-ms", type=int, default=120, help="事件触发后等待多少毫秒再重读列表。默认 120。")
    parser.add_argument("--poll-seconds", type=float, default=0.0, help="额外轮询间隔；0 表示只靠 UIA 事件触发。默认 0。")
    parser.add_argument("--print-unchanged", action="store_true", help="即使列表内容没变，也打印事件触发后的快照。")
    parser.add_argument("--clear", action="store_true", help="每次刷新前清屏，只保留最新快照。")
    parser.add_argument("--max-events", type=int, default=0, help="收到 N 个 UIA 事件后退出；0 表示不限。")
    return parser.parse_args(argv)


def format_text(text: str, max_text: int, full_text: bool) -> str:
    if full_text:
        return text or "(空)"
    return shorten(text or "(空)", max_text)


def collect_list_item_text(backend: WindowBackend, item: Any, max_depth: int) -> tuple[str, str, int]:
    max_depth = max(0, min(max_depth, 20))
    stack: list[tuple[Any, int]] = [(item, 0)]
    seen_nodes: set[int] = set()
    seen_lines: set[str] = set()
    lines: list[str] = []
    self_text = backend._safe_text(item).strip()
    node_count = 0

    while stack:
        node, depth = stack.pop()
        marker = id(node)
        if marker in seen_nodes or depth > max_depth:
            continue
        seen_nodes.add(marker)
        node_count += 1
        text = backend._safe_text(node).strip()
        for raw_line in text.replace("\r", "\n").split("\n"):
            line = raw_line.strip()
            if not line or line in seen_lines:
                continue
            seen_lines.add(line)
            lines.append(line)
        if depth >= max_depth:
            continue
        try:
            children = backend.children(node)
        except Exception:
            children = []
        for child in reversed(children):
            stack.append((child, depth + 1))

    return "\n".join(lines), self_text, node_count


def snapshot_sns_items(
    backend: WindowBackend,
    win: Any,
    max_items: int,
    show_empty: bool,
    text_depth: int,
) -> tuple[Optional[Rect], list[ListItemSnapshot]]:
    sns_list = find_sns_list_control(backend, win, max_depth=20)
    sns_rect = backend.rect(sns_list) if sns_list is not None else None
    if sns_list is None:
        return None, []
    limit = max(1, max_items) if max_items > 0 else None
    items = find_uia_list_items(backend, win, max_depth=20, limit=limit)
    snapshots: list[ListItemSnapshot] = []
    for index, item in enumerate(items):
        raw_text, self_text, node_count = collect_list_item_text(backend, item, text_depth)
        text = clean_post_text(raw_text)
        if not show_empty and not text:
            continue
        snapshots.append(
            ListItemSnapshot(
                index=index,
                control_type=backend._control_type(item),
                automation_id=backend._automation_id(item),
                rect=backend.rect(item),
                text=text,
                self_text=self_text,
                raw_text=raw_text,
                node_count=node_count,
            )
        )
    return sns_rect, snapshots


def list_signature(items: list[ListItemSnapshot]) -> str:
    return "\n".join(item.signature_part() for item in items)


def print_snapshot(
    *,
    sns_rect: Optional[Rect],
    items: list[ListItemSnapshot],
    snapshot_index: int,
    reasons: list[str],
    event_count: int,
    max_text: int,
    full_text: bool,
    clear: bool,
) -> str:
    signature = list_signature(items)
    if clear:
        print("\x1b[2J\x1b[H", end="")
    now = _dt.datetime.now().strftime("%H:%M:%S.%f")[:-3]
    reason_text = ", ".join(reasons[-4:]) if reasons else "manual"
    rect_text = sns_rect.describe() if sns_rect else "未找到"
    print(f"\n=== sns_list 快照 #{snapshot_index} {now} | UIA事件累计={event_count} | 触发={reason_text} ===", flush=True)
    print(f"sns_list: {rect_text} | ListItem={len(items)}", flush=True)
    if not items:
        print("  (当前没有读取到 ListItem；可能 UIA 只暴露了渲染壳，或朋友圈窗口尚未稳定。)", flush=True)
        return signature
    for item in items:
        rect_text = item.rect.describe() if item.rect else "no-rect"
        print(f"  #{item.index:02d} [{item.control_type or '?'}] {rect_text} nodes={item.node_count}", flush=True)
        text = format_text(item.text, max_text=max_text, full_text=full_text)
        if not item.text and item.self_text:
            print(f"      self={shorten(item.self_text, max_text)}", flush=True)
        if full_text:
            for line in text.splitlines() or ["(空)"]:
                print(f"      {line}", flush=True)
        else:
            print(f"      {text}", flush=True)
    return signature


def wait_for_sns_list(automation: Any, root_element: Any, timeout_seconds: float = 3.0) -> Optional[Any]:
    deadline = time.perf_counter() + max(0.0, timeout_seconds)
    while True:
        sns_list = find_sns_list_element(automation, root_element)
        if sns_list is not None:
            return sns_list
        if time.perf_counter() >= deadline:
            return None
        comtypes.client.PumpEvents(0.05)
        time.sleep(0.05)


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    backend = WindowBackend()
    win = backend.moments_window()
    backend.activate(win)
    win_rect = backend.rect(win)
    if win_rect is None:
        raise EasyMoneyError("无法读取朋友圈窗口位置")
    hwnd = window_handle(win)

    automation = comtypes.client.CreateObject(UIA.CUIAutomation, interface=UIA.IUIAutomation)
    root_element = automation.ElementFromHandle(ctypes.c_void_p(hwnd))
    sns_element = wait_for_sns_list(automation, root_element, timeout_seconds=3.0)
    if sns_element is None:
        print("警告: 启动时 UIA 未暴露 sns_list；仍会监听窗口事件，并在后续快照中继续尝试读取。", flush=True)

    event_root = root_element
    event_root_label = "window"
    if args.event_root == "sns-list":
        if sns_element is None:
            raise EasyMoneyError("当前 UIA 树未暴露 sns_list，无法以 sns-list 为事件根；请改用 --event-root window")
        event_root = sns_element
        event_root_label = "sns-list"

    print(f"UIA sns_list 实时列表启动: event_root={event_root_label} hwnd={hwnd} window={win_rect.describe()}")
    print(f"事件根节点: {snapshot_element(event_root).describe()}")
    print("输出的是 sns_list 下 ListItem 的当前快照；UIA 事件到达后会自动重新读取。按 Ctrl+C 退出。", flush=True)

    trigger = SnsRefreshTrigger(max_events=args.max_events)
    event_args = argparse.Namespace(events=args.events, properties=args.properties, scope=args.scope)
    handlers = register_handlers(event_args, automation, event_root, trigger)
    if not handlers:
        raise EasyMoneyError("没有注册任何 UIA 事件，请检查 --events 参数")

    deadline = None if args.seconds <= 0 else time.perf_counter() + args.seconds
    next_poll = time.perf_counter() + args.poll_seconds if args.poll_seconds > 0 else None
    last_signature: Optional[str] = None
    snapshot_index = 0
    debounce_seconds = max(0.0, args.debounce_ms / 1000)
    try:
        while not trigger.stop_requested:
            now = time.perf_counter()
            if deadline is not None and now >= deadline:
                break
            if next_poll is not None and now >= next_poll:
                trigger.mark_poll()
                next_poll = now + args.poll_seconds
            comtypes.client.PumpEvents(0.03)
            if not trigger.ready(debounce_seconds):
                continue
            reasons = trigger.consume_reasons()
            sns_rect, items = snapshot_sns_items(backend, win, args.limit, args.show_empty, args.text_depth)
            signature = list_signature(items)
            if args.print_unchanged or signature != last_signature:
                snapshot_index += 1
                last_signature = print_snapshot(
                    sns_rect=sns_rect,
                    items=items,
                    snapshot_index=snapshot_index,
                    reasons=reasons,
                    event_count=trigger.event_count,
                    max_text=args.max_text,
                    full_text=args.full_text,
                    clear=args.clear,
                )
    except KeyboardInterrupt:
        print("\n收到 Ctrl+C，停止监听。")
    finally:
        unregister_handlers(automation, event_root, handlers)
    print("UIA sns_list 实时列表结束。")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main(sys.argv[1:]))
    except EasyMoneyError as exc:
        print(f"错误: {exc}", file=sys.stderr)
        raise SystemExit(1)
