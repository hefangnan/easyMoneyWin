#!/usr/bin/env python3
from __future__ import annotations

import argparse
import ctypes
import sys
import time
from typing import Any, Iterable, Optional

from easy_money_win_core import Rect
from easy_money_win_uia import WindowBackend, find_list_items_under_control, find_sns_list_control


IMAGE_NAME = "\u56fe\u7247"
IMAGE_CLASS = "mmui::XMouseEventView"


def shorten(value: str, limit: int = 160) -> str:
    if len(value) <= limit:
        return value
    return value[: max(0, limit - 3)] + "..."


def safe_current(element: Any, attr: str, default: Any = "") -> Any:
    try:
        value = getattr(element, attr)
        return value() if callable(value) else value
    except Exception:
        return default


def runtime_id_text(backend: WindowBackend, element: Any) -> str:
    try:
        identity = backend._control_identity(element)
        return repr(identity)
    except Exception:
        return "?"


def get_legacy_name(backend: WindowBackend, element: Any) -> str:
    try:
        _, UIA = backend._ensure_automation()
        pattern = element.GetCurrentPattern(UIA.UIA_LegacyIAccessiblePatternId)
        if not pattern:
            return ""
        try:
            pattern = pattern.QueryInterface(UIA.IUIAutomationLegacyIAccessiblePattern)
        except Exception:
            pass
        value = safe_current(pattern, "CurrentName", "")
        return str(value or "")
    except Exception:
        return ""


def has_invoke_pattern(backend: WindowBackend, element: Any) -> bool:
    try:
        _, UIA = backend._ensure_automation()
        return bool(element.GetCurrentPattern(UIA.UIA_InvokePatternId))
    except Exception:
        return False


def rect_inside_or_intersects(rect: Optional[Rect], container: Optional[Rect]) -> bool:
    if rect is None or container is None:
        return False
    return (
        rect.left < container.right
        and rect.right > container.left
        and rect.top < container.bottom
        and rect.bottom > container.top
    )


def describe_node(backend: WindowBackend, element: Any, container: Optional[Rect] = None) -> str:
    name = backend._safe_text(element).strip()
    legacy_name = get_legacy_name(backend, element).strip()
    control_type = backend._control_type(element)
    class_name = backend._class_name(element)
    automation_id = backend._automation_id(element)
    rect = backend.rect(element)
    focused = bool(safe_current(element, "CurrentHasKeyboardFocus", False))
    focusable = bool(safe_current(element, "CurrentIsKeyboardFocusable", False))
    inside = rect_inside_or_intersects(rect, container)
    parts = [
        f"type={control_type or '?'}",
        f"name={name!r}",
        f"legacy={legacy_name!r}",
        f"class={class_name!r}",
        f"id={automation_id!r}",
        f"rect={rect.describe() if rect else '?'}",
        f"focusable={str(focusable).lower()}",
        f"focused={str(focused).lower()}",
        f"invoke={str(has_invoke_pattern(backend, element)).lower()}",
        f"in_item={str(inside).lower()}",
        f"runtime={runtime_id_text(backend, element)}",
    ]
    return " ".join(parts)


def is_image_button_candidate(backend: WindowBackend, element: Any, item_rect: Optional[Rect]) -> bool:
    if not is_image_like_button(backend, element):
        return False
    return rect_inside_or_intersects(backend.rect(element), item_rect)


def is_image_like_button(backend: WindowBackend, element: Any) -> bool:
    name = backend._safe_text(element).strip()
    legacy_name = get_legacy_name(backend, element).strip()
    control_type = backend._control_type(element).lower()
    class_name = backend._class_name(element)
    if control_type != "button":
        return False
    return name == IMAGE_NAME or legacy_name == IMAGE_NAME or class_name == IMAGE_CLASS


def iter_view_nodes(backend: WindowBackend, root: Any, view: str, max_depth: int) -> Iterable[tuple[Any, int]]:
    if view == "pywinauto":
        yield from backend.iter_tree(root, max_depth=max_depth)
    else:
        yield from backend.iter_tree_view(root, max_depth=max_depth, view=view)


def main() -> int:
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            try:
                reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass

    parser = argparse.ArgumentParser(description="Probe UIA image-like buttons under a WeChat Moments ListItem.")
    parser.add_argument("--item-index", type=int, default=2, help="1-based Moments ListItem index. Default: 2.")
    parser.add_argument("--max-depth", type=int, default=8, help="Traversal depth under the ListItem. Default: 8.")
    parser.add_argument("--views", default="raw,control,content,pywinauto", help="Comma-separated views to scan.")
    parser.add_argument("--all-buttons", action="store_true", help="Print all buttons inside the ListItem, not only image candidates.")
    parser.add_argument("--window-scan-depth", type=int, default=18, help="Traversal depth for whole-window fallback scan. Default: 18.")
    parser.add_argument("--point", default="", help="Probe ElementFromPoint at x,y before scanning.")
    parser.add_argument("--wait-focus", type=float, default=0.0, help="Wait N seconds before reading focused element.")
    parser.add_argument("--activate-window", action="store_true", help="Activate the Moments window before point probing/scanning.")
    args = parser.parse_args()

    item_index = max(1, args.item_index)
    views = [item.strip().lower() for item in args.views.split(",") if item.strip()]

    backend = WindowBackend()
    automation, _ = backend._ensure_automation()
    win: Any = None

    if args.activate_window:
        win = backend.moments_window()
        backend.activate(win)
        time.sleep(0.15)

    if args.wait_focus > 0:
        print(f"Waiting {args.wait_focus:g}s before reading focused element...")
        time.sleep(args.wait_focus)

    focused = automation.GetFocusedElement()
    focused_hit = False
    point_hit = False
    if focused:
        print("Focused element:")
        print("  " + describe_node(backend, focused))
        focused_hit = is_image_like_button(backend, focused)

    if args.point.strip():
        try:
            x_text, y_text = args.point.split(",", 1)
            point = ctypes.wintypes.POINT(int(float(x_text)), int(float(y_text)))
            hit = automation.ElementFromPoint(point)
        except Exception as exc:
            print(f"ElementFromPoint failed: {exc}")
            hit = None
        if hit:
            print(f"ElementFromPoint {args.point}:")
            print("  " + describe_node(backend, hit))
            point_hit = is_image_like_button(backend, hit)

    if win is None:
        win = backend.moments_window()
    win_rect = backend.rect(win)
    print(f"Window rect: {win_rect.describe() if win_rect else '?'}")

    sns_list = find_sns_list_control(backend, win)
    if sns_list is None:
        print("sns_list not found")
        return 2

    items = find_list_items_under_control(backend, sns_list, limit=item_index)
    if len(items) < item_index:
        print(f"ListItem not enough: got={len(items)} need={item_index}")
        return 2

    item = items[item_index - 1]
    item_rect = backend.rect(item)
    item_text = backend._safe_text(item).replace("\r", " ").replace("\n", " ").strip()
    print(f"Target ListItem #{item_index}: rect={item_rect.describe() if item_rect else '?'} text={shorten(item_text, 160)!r}")

    seen: set[Any] = set()
    matches: list[str] = []
    button_count = 0

    def scan_scope(scope_name: str, root: Any, depth_limit: int) -> None:
        nonlocal button_count
        for view in views:
            try:
                nodes = list(iter_view_nodes(backend, root, view=view, max_depth=max(1, depth_limit)))
            except Exception as exc:
                print(f"[{scope_name}:{view}] traversal failed: {exc}")
                continue
            view_matches = 0
            for node, depth in nodes:
                identity = runtime_id_text(backend, node)
                marker = (scope_name, view, identity)
                if marker in seen:
                    continue
                seen.add(marker)
                control_type = backend._control_type(node).lower()
                if control_type == "button" and rect_inside_or_intersects(backend.rect(node), item_rect):
                    button_count += 1
                    if args.all_buttons:
                        print(f"[{scope_name}:{view}] button depth={depth} {describe_node(backend, node, item_rect)}")
                if is_image_button_candidate(backend, node, item_rect):
                    view_matches += 1
                    matches.append(f"[{scope_name}:{view}] depth={depth} {describe_node(backend, node, item_rect)}")
            print(f"[{scope_name}:{view}] nodes={len(nodes)} image_candidates={view_matches}")

    scan_scope("item", item, args.max_depth)
    scan_scope("window", win, args.window_scan_depth)

    print(f"Buttons inside item: {button_count}")
    print(f"Image candidates: {len(matches)}")
    for line in matches:
        print("  " + line)
    return 0 if matches or focused_hit or point_hit else 1


if __name__ == "__main__":
    raise SystemExit(main())
