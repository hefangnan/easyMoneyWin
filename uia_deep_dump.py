from __future__ import annotations

import argparse
import ctypes
import sys
from pathlib import Path
from typing import Any, Iterable, Optional

import comtypes.client
import win32con
import win32gui

from easy_money_win import EasyMoneyError, Rect, WindowBackend
from uia_event_watch import UIA, snapshot_element, window_handle


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="深度 dump 当前微信/朋友圈窗口暴露的 UI 信息。")
    parser.add_argument("--view", choices=["raw", "control", "content", "all"], default="all", help="UIA 视图。默认 all。")
    parser.add_argument("--max-depth", type=int, default=20, help="最大递归深度。默认 20。")
    parser.add_argument("--win32", action="store_true", help="同时 dump Win32 子窗口树。")
    parser.add_argument("--output", type=Path, help="保存到文件，同时也打印到屏幕。")
    return parser.parse_args(argv)


class Tee:
    def __init__(self, path: Optional[Path]) -> None:
        self.file = None
        if path is not None:
            path.parent.mkdir(parents=True, exist_ok=True)
            self.file = path.open("w", encoding="utf-8")

    def write(self, text: str = "") -> None:
        print(text)
        if self.file is not None:
            self.file.write(text + "\n")

    def close(self) -> None:
        if self.file is not None:
            self.file.close()


def walker_for_view(automation: Any, view: str) -> Any:
    if view == "raw":
        return automation.RawViewWalker
    if view == "control":
        return automation.ControlViewWalker
    if view == "content":
        return automation.ContentViewWalker
    raise EasyMoneyError(f"未知 UIA view: {view}")


def iter_uia_tree(walker: Any, root: Any, max_depth: int) -> Iterable[tuple[Any, int]]:
    stack: list[tuple[Any, int]] = [(root, 0)]
    seen: set[int] = set()
    while stack:
        element, depth = stack.pop()
        marker = id(element)
        if marker in seen or depth > max_depth:
            continue
        seen.add(marker)
        yield element, depth
        if depth >= max_depth:
            continue
        children: list[Any] = []
        try:
            child = walker.GetFirstChildElement(element)
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


def dump_uia_view(out: Tee, automation: Any, root: Any, view: str, max_depth: int) -> None:
    walker = walker_for_view(automation, view)
    out.write(f"\n=== UIA {view.upper()} VIEW ===")
    count = 0
    found_sns_list = False
    found_list_item = False
    for element, depth in iter_uia_tree(walker, root, max_depth):
        snap = snapshot_element(element)
        indent = "  " * depth
        if snap.automation_id == "sns_list":
            found_sns_list = True
        if "ListItem" in snap.control_type:
            found_list_item = True
        out.write(f"{indent}[{snap.describe()}]")
        count += 1
    out.write(f"UIA {view} 节点数: {count}; sns_list={found_sns_list}; ListItem={found_list_item}")


def rect_from_win32(hwnd: int) -> Optional[Rect]:
    try:
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    except Exception:
        return None
    if right <= left or bottom <= top:
        return None
    return Rect(float(left), float(top), float(right), float(bottom))


def window_text(hwnd: int) -> str:
    try:
        return win32gui.GetWindowText(hwnd) or ""
    except Exception:
        return ""


def class_name(hwnd: int) -> str:
    try:
        return win32gui.GetClassName(hwnd) or ""
    except Exception:
        return ""


def win32_style_flags(hwnd: int) -> str:
    try:
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
    except Exception:
        return ""
    flags: list[str] = []
    for name, value in [
        ("visible", win32con.WS_VISIBLE),
        ("child", win32con.WS_CHILD),
        ("disabled", win32con.WS_DISABLED),
        ("popup", win32con.WS_POPUP),
    ]:
        if style & value:
            flags.append(name)
    return ",".join(flags)


def dump_win32_tree(out: Tee, hwnd: int, max_depth: int) -> None:
    out.write("\n=== WIN32 HWND TREE ===")
    count = 0

    def visit(node: int, depth: int) -> None:
        nonlocal count
        if depth > max_depth:
            return
        rect = rect_from_win32(node)
        rect_text = rect.describe() if rect else "no-rect"
        out.write(
            f"{'  ' * depth}[hwnd={node}] class={class_name(node)!r} "
            f"text={window_text(node)!r} style={win32_style_flags(node)!r} rect={rect_text}"
        )
        count += 1
        children: list[int] = []
        try:
            win32gui.EnumChildWindows(node, lambda child, _: children.append(child), None)
        except Exception:
            return
        direct_children = [child for child in children if win32gui.GetParent(child) == node]
        for child in direct_children:
            visit(child, depth + 1)

    visit(hwnd, 0)
    out.write(f"Win32 节点数: {count}")


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    backend = WindowBackend()
    win = backend.moments_window()
    backend.activate(win)
    rect = backend.rect(win)
    if rect is None:
        raise EasyMoneyError("无法读取朋友圈窗口位置")
    hwnd = window_handle(win)
    automation = comtypes.client.CreateObject(UIA.CUIAutomation, interface=UIA.IUIAutomation)
    root = automation.ElementFromHandle(ctypes.c_void_p(hwnd))

    out = Tee(args.output)
    try:
        out.write(f"目标窗口 hwnd={hwnd} rect={rect.describe()} title={backend._safe_text(win)!r}")
        views = ["raw", "control", "content"] if args.view == "all" else [args.view]
        for view in views:
            dump_uia_view(out, automation, root, view, max(0, args.max_depth))
        if args.win32:
            dump_win32_tree(out, hwnd, max(0, args.max_depth))
        if args.output is not None:
            out.write(f"\n已保存: {args.output}")
    finally:
        out.close()
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main(sys.argv[1:]))
    except EasyMoneyError as exc:
        print(f"错误: {exc}", file=sys.stderr)
        raise SystemExit(1)
