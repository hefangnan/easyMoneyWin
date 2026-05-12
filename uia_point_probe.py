from __future__ import annotations

import argparse
import ctypes
import sys
import time
from pathlib import Path
from typing import Any, Optional

import comtypes.client
import win32gui
from comtypes import COMError
from comtypes.automation import VARIANT

from easy_money_win import EasyMoneyError
from uia_deep_dump import Tee, dump_uia_view, iter_uia_tree
from uia_event_watch import CONTROL_TYPE_NAMES, UIA, format_variant, rect_from_uia, shorten, snapshot_element

comtypes.client.GetModule("oleacc.dll")
from comtypes.gen import Accessibility  # noqa: E402


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="读取鼠标所在位置的 UIA 元素，对齐 Inspect 的命中结果。")
    parser.add_argument("--x", type=int, help="屏幕 X 坐标；不传则使用当前鼠标位置。")
    parser.add_argument("--y", type=int, help="屏幕 Y 坐标；不传则使用当前鼠标位置。")
    parser.add_argument("--countdown", type=float, default=3.0, help="未传坐标时，等待几秒让你把鼠标放到目标文字上。默认 3。")
    parser.add_argument("--subtree-depth", type=int, default=6, help="从命中元素向下 dump RawView 的深度。默认 6。")
    parser.add_argument("--ancestor-limit", type=int, default=20, help="向上追溯父链的最大层数。默认 20。")
    parser.add_argument("--text-limit", type=int, default=2000, help="TextPattern/Legacy 文本最大读取长度。默认 2000。")
    parser.add_argument("--output", type=Path, help="保存到文件，同时也打印到屏幕。")
    return parser.parse_args(argv)


def current_or_empty(element: Any, attr: str) -> Any:
    try:
        value = getattr(element, attr)
        return value() if callable(value) else value
    except Exception:
        return None


def current_property_lines(element: Any) -> list[str]:
    attrs = [
        "CurrentName",
        "CurrentAutomationId",
        "CurrentClassName",
        "CurrentControlType",
        "CurrentLocalizedControlType",
        "CurrentFrameworkId",
        "CurrentProcessId",
        "CurrentNativeWindowHandle",
        "CurrentIsControlElement",
        "CurrentIsContentElement",
        "CurrentIsOffscreen",
        "CurrentHelpText",
        "CurrentItemType",
        "CurrentItemStatus",
        "CurrentAriaRole",
        "CurrentAriaProperties",
        "CurrentBoundingRectangle",
    ]
    lines: list[str] = []
    for attr in attrs:
        value = current_or_empty(element, attr)
        if value is None or value == "":
            continue
        if attr == "CurrentControlType":
            value = CONTROL_TYPE_NAMES.get(int(value), str(value))
        elif attr == "CurrentBoundingRectangle":
            rect = rect_from_uia(value)
            value = rect.describe() if rect else format_variant(value)
        lines.append(f"{attr}: {value!r}" if isinstance(value, str) else f"{attr}: {value}")
    return lines


def all_pattern_ids() -> list[tuple[str, int]]:
    pairs: list[tuple[str, int]] = []
    for name in dir(UIA):
        if not name.startswith("UIA_") or not name.endswith("PatternId"):
            continue
        value = getattr(UIA, name)
        if isinstance(value, int):
            pairs.append((name.removeprefix("UIA_").removesuffix("PatternId"), value))
    return sorted(pairs, key=lambda item: item[0])


def get_pattern(element: Any, pattern_id: int, interface: Any) -> Optional[Any]:
    try:
        pattern = element.GetCurrentPattern(pattern_id)
    except (COMError, OSError):
        return None
    except Exception:
        return None
    if not pattern:
        return None
    try:
        return pattern.QueryInterface(interface)
    except Exception:
        return pattern


def supported_pattern_names(element: Any) -> list[str]:
    names: list[str] = []
    for name, pattern_id in all_pattern_ids():
        try:
            pattern = element.GetCurrentPattern(pattern_id)
        except Exception:
            continue
        if pattern:
            names.append(name)
    return names


def pattern_detail_lines(element: Any, point: tuple[int, int], text_limit: int) -> list[str]:
    lines: list[str] = []

    legacy = get_pattern(element, UIA.UIA_LegacyIAccessiblePatternId, UIA.IUIAutomationLegacyIAccessiblePattern)
    if legacy is not None:
        lines.append("LegacyIAccessiblePattern:")
        for attr in [
            "CurrentName",
            "CurrentValue",
            "CurrentDescription",
            "CurrentRole",
            "CurrentState",
            "CurrentDefaultAction",
            "CurrentChildId",
        ]:
            value = current_or_empty(legacy, attr)
            if value is not None and value != "":
                lines.append(f"  {attr}: {shorten(str(value), text_limit)!r}")

    value_pattern = get_pattern(element, UIA.UIA_ValuePatternId, UIA.IUIAutomationValuePattern)
    if value_pattern is not None:
        value = current_or_empty(value_pattern, "CurrentValue")
        lines.append(f"ValuePattern.CurrentValue: {shorten(str(value or ''), text_limit)!r}")

    text_pattern = get_pattern(element, UIA.UIA_TextPatternId, UIA.IUIAutomationTextPattern)
    if text_pattern is not None:
        lines.append("TextPattern:")
        try:
            doc_text = text_pattern.DocumentRange.GetText(text_limit)
            lines.append(f"  DocumentRange: {shorten(doc_text or '', text_limit)!r}")
        except Exception as exc:
            lines.append(f"  DocumentRange读取失败: {exc}")
        try:
            pt = ctypes.wintypes.POINT(point[0], point[1])
            text_range = text_pattern.RangeFromPoint(pt)
            point_text = text_range.GetText(text_limit)
            lines.append(f"  RangeFromPoint: {shorten(point_text or '', text_limit)!r}")
        except Exception as exc:
            lines.append(f"  RangeFromPoint读取失败: {exc}")

    return lines


def msaa_accessible_from_point(point: tuple[int, int]) -> tuple[Optional[Any], Optional[VARIANT]]:
    pacc = ctypes.POINTER(Accessibility.IAccessible)()
    child = VARIANT()
    try:
        ctypes.oledll.oleacc.AccessibleObjectFromPoint(
            ctypes.wintypes.POINT(point[0], point[1]),
            ctypes.byref(pacc),
            ctypes.byref(child),
        )
    except Exception:
        return None, None
    if not pacc:
        return None, None
    return pacc, child


def msaa_value(accessible: Any, attr: str, child: VARIANT) -> Any:
    try:
        return getattr(accessible, attr)(child)
    except Exception:
        return None


def dump_msaa_point(out: Tee, point: tuple[int, int], text_limit: int) -> None:
    out.write("\n=== MSAA AccessibleObjectFromPoint ===")
    accessible, child = msaa_accessible_from_point(point)
    if accessible is None or child is None:
        out.write("(无 MSAA 命中)")
        return
    out.write(f"child={getattr(child, 'value', child)!r}")
    for attr in ["accName", "accValue", "accDescription", "accRole", "accState", "accDefaultAction", "accKeyboardShortcut"]:
        value = msaa_value(accessible, attr, child)
        if value is None or value == "":
            continue
        out.write(f"{attr}: {shorten(str(value), text_limit)!r}")
    try:
        left = ctypes.c_long()
        top = ctypes.c_long()
        width = ctypes.c_long()
        height = ctypes.c_long()
        accessible.accLocation(ctypes.byref(left), ctypes.byref(top), ctypes.byref(width), ctypes.byref(height), child)
        out.write(f"accLocation: ({left.value},{top.value}) {width.value}x{height.value}")
    except Exception as exc:
        out.write(f"accLocation读取失败: {exc}")


def dump_ancestors(out: Tee, automation: Any, element: Any, ancestor_limit: int) -> None:
    walker = automation.RawViewWalker
    chain: list[Any] = []
    current = element
    for _ in range(max(1, ancestor_limit)):
        if not current:
            break
        chain.append(current)
        try:
            current = walker.GetParentElement(current)
        except Exception:
            break
    out.write("\n=== RAW VIEW 父链：命中元素 -> 根 ===")
    for index, node in enumerate(chain):
        out.write(f"#{index:02d} {snapshot_element(node).describe()}")


def dump_hit_subtree(out: Tee, automation: Any, element: Any, max_depth: int) -> None:
    out.write("\n=== 命中元素 RawView 子树 ===")
    walker = automation.RawViewWalker
    count = 0
    for node, depth in iter_uia_tree(walker, element, max(0, max_depth)):
        out.write(f"{'  ' * depth}[{snapshot_element(node).describe()}]")
        count += 1
    out.write(f"命中元素子树节点数: {count}")


def point_from_args(args: argparse.Namespace) -> tuple[int, int]:
    if args.x is not None or args.y is not None:
        if args.x is None or args.y is None:
            raise EasyMoneyError("--x 和 --y 必须同时提供")
        return args.x, args.y
    if args.countdown > 0:
        print(f"请把鼠标放到 Inspect 能看到正文的那行文字上，{args.countdown:g} 秒后读取...")
        time.sleep(args.countdown)
    return win32gui.GetCursorPos()


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    x, y = point_from_args(args)
    automation = comtypes.client.CreateObject(UIA.CUIAutomation, interface=UIA.IUIAutomation)
    element = automation.ElementFromPoint(ctypes.wintypes.POINT(x, y))
    if not element:
        raise EasyMoneyError(f"ElementFromPoint({x}, {y}) 没有返回元素")

    out = Tee(args.output)
    try:
        out.write(f"命中坐标: ({x}, {y})")
        out.write(f"命中元素: {snapshot_element(element).describe()}")
        out.write("\n=== 当前属性 ===")
        for line in current_property_lines(element):
            out.write(line)
        patterns = supported_pattern_names(element)
        out.write("\n=== 支持的 Pattern ===")
        out.write(", ".join(patterns) if patterns else "(无)")
        details = pattern_detail_lines(element, (x, y), max(1, args.text_limit))
        if details:
            out.write("\n=== Pattern 详情 ===")
            for line in details:
                out.write(line)
        dump_msaa_point(out, (x, y), max(1, args.text_limit))
        dump_ancestors(out, automation, element, args.ancestor_limit)
        dump_hit_subtree(out, automation, element, args.subtree_depth)
        out.write("\n=== 从命中元素所属顶层 RawView 再 dump 一遍 ===")
        top = element
        walker = automation.RawViewWalker
        for _ in range(max(1, args.ancestor_limit)):
            try:
                parent = walker.GetParentElement(top)
            except Exception:
                break
            if not parent:
                break
            top = parent
        dump_uia_view(out, automation, top, "raw", args.subtree_depth)
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
