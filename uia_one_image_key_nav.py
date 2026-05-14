from __future__ import annotations

import argparse
import os
import re
import sys
import time
from pathlib import Path

from easy_money_win_core import (
    COMMENT_REFRESH_WAIT_SECONDS,
    CONFIG_REFRESH,
    DEBUG_DIR,
    EasyMoneyError,
    Point,
    Rect,
    UIAListItemUnavailable,
    WindowPositionUnavailable,
    current_timestamp_ms,
    load_point,
)
from easy_money_win_capture import CaptureBackend
from easy_money_win_input import InputBackend
from easy_money_win_uia import WindowBackend, resolve_second_uia_list_item_post


KEY_SEQUENCE = ("down", "tab", "tab", "tab", "enter")
FOCUS_CAPTURE_SEQUENCE = ("down", "tab", "tab", "tab")
ONE_IMAGE_RE = re.compile(r"(?<![0-9一二两三四五六七八九十])(?:包含|含)?\s*(?:1|一)\s*张图片")


def configure_output() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="匹配指定朋友圈 ListItem；若文案表示一张图片，则发送 Down、Tab、Tab、Tab、Enter。"
    )
    parser.add_argument("--user", required=True, help="目标朋友圈文案开头的用户标识。")
    parser.add_argument("--item-index", type=int, default=2, help="读取第几个朋友圈 ListItem，1 表示第一条；默认 2。")
    parser.add_argument("--rounds", type=int, default=30, help="匹配失败或不是一张图片时刷新重试的轮数；默认 30。")
    parser.add_argument("--settle-ms", type=int, default=220, help="等待 UIA 列表稳定的毫秒数；默认 220。")
    parser.add_argument(
        "--refresh-wait-ms",
        type=int,
        default=int(COMMENT_REFRESH_WAIT_SECONDS * 1000),
        help="点击 locate 保存的刷新坐标后等待多久；默认 100。",
    )
    parser.add_argument("--activate-wait-ms", type=int, default=80, help="激活朋友圈窗口后等待多久再匹配；默认 80。")
    parser.add_argument("--key-gap-ms", type=int, default=30, help="每个按键之间的间隔；0 表示一次 SendInput 发出。默认 30。")
    parser.add_argument(
        "--capture-focus",
        action="store_true",
        help="匹配一张图片后只按 Down、Tab、Tab、Tab，截图当前焦点区域并退出，不发送 Enter。",
    )
    parser.add_argument("--focus-wait-ms", type=int, default=120, help="按完 Down+3Tab 后等待多久再读取焦点；默认 120。")
    parser.add_argument("--focus-padding", type=int, default=0, help="焦点矩形截图向外扩多少像素；默认 0。")
    parser.add_argument("--output", type=Path, help="焦点截图输出文件或目录；默认保存到 DEBUG_DIR。")
    parser.add_argument("--dry-run", action="store_true", help="只跳过最终键盘序列；刷新重试仍会点击刷新坐标。")
    parser.add_argument("--no-refresh", action="store_true", help="匹配失败或不是一张图片时不点击刷新。")
    parser.add_argument("--no-activate", action="store_true", help="不主动激活朋友圈窗口。")
    return parser.parse_args(argv)


def is_one_image_post(text: str, inline_image_count: int | None) -> bool:
    if inline_image_count is not None:
        return inline_image_count == 1
    return bool(ONE_IMAGE_RE.search(text))


def is_missing_sns_list_error(exc: Exception) -> bool:
    return "sns_list" in str(exc)


def key_sequence_label(keys: tuple[str, ...] = KEY_SEQUENCE) -> str:
    return " -> ".join(keys)


def refresh_button_center(window_rect: object) -> Point:
    refresh_offset = load_point(CONFIG_REFRESH)
    if refresh_offset is None:
        raise EasyMoneyError("未找到刷新按钮坐标配置，请先运行 locate")
    left = getattr(window_rect, "left")
    top = getattr(window_rect, "top")
    return Point(left + refresh_offset.x, top + refresh_offset.y)


def click_refresh(input_backend: InputBackend, point: Point, wait_ms: int) -> None:
    input_backend.click(point)
    wait_seconds = max(0, wait_ms) / 1000
    if wait_seconds > 0:
        time.sleep(wait_seconds)
    print(f"  已点击 locate 保存的刷新坐标: ({int(point.x)}, {int(point.y)})，等待 {max(0, wait_ms)}ms")


def output_path(raw: Path | None) -> Path:
    filename = f"wechat_focus_after_tabs_{time.strftime('%Y%m%d_%H%M%S')}.png"
    if raw is None:
        return DEBUG_DIR / filename
    path = raw.expanduser()
    if path.suffix.lower() == ".png":
        return path
    return path / filename


def rect_intersects(lhs: Rect, rhs: Rect) -> bool:
    return lhs.left < rhs.right and lhs.right > rhs.left and lhs.top < rhs.bottom and lhs.bottom > rhs.top


def describe_focus_element(backend: WindowBackend, element: object, rect: Rect) -> str:
    name = backend._safe_text(element).strip()
    control_type = backend._control_type(element)
    class_name = backend._class_name(element)
    return f"type={control_type or '?'} name={name or '(空)'} class={class_name or '(空)'} rect={rect.describe()}"


def capture_focus_after_tabs(
    backend: WindowBackend,
    input_backend: InputBackend,
    window_rect: Rect,
    key_gap_ms: int,
    focus_wait_ms: int,
    focus_padding: int,
    raw_output: Path | None,
) -> Path:
    input_backend.press_sequence(FOCUS_CAPTURE_SEQUENCE, gap=max(0, key_gap_ms) / 1000)
    if focus_wait_ms > 0:
        time.sleep(focus_wait_ms / 1000)

    automation, _ = backend._ensure_automation()
    focused = automation.GetFocusedElement()
    if focused is None:
        raise EasyMoneyError("按下 Down+3Tab 后没有读取到当前焦点元素")
    focus_rect = backend.rect(focused)
    if focus_rect is None:
        raise EasyMoneyError("按下 Down+3Tab 后当前焦点元素没有有效矩形")

    expanded = focus_rect.expanded(max(0, focus_padding), max(0, focus_padding))
    crop_rect = expanded.clamp_to(window_rect) if rect_intersects(expanded, window_rect) else expanded
    if crop_rect.width <= 0 or crop_rect.height <= 0:
        raise EasyMoneyError(f"焦点截图区域无效: {crop_rect.describe()}")

    path = output_path(raw_output)
    capture = CaptureBackend()
    try:
        image = capture.screenshot(crop_rect)
        capture.save(image, path)
    finally:
        capture.close()

    print(f"  当前焦点: {describe_focus_element(backend, focused, focus_rect)}")
    print(f"  焦点截图区域: {crop_rect.describe()}")
    print(f"  焦点截图已保存: {path}")
    return path


def main(argv: list[str]) -> int:
    configure_output()
    args = parse_args(argv)
    if args.item_index < 1:
        raise EasyMoneyError("--item-index 必须大于等于 1")
    if args.rounds < 1:
        raise EasyMoneyError("--rounds 必须大于等于 1")

    backend = WindowBackend()
    input_backend = InputBackend()
    input_backend.prepare_key_sequence(KEY_SEQUENCE)
    input_backend.prepare_key_sequence(FOCUS_CAPTURE_SEQUENCE)

    win = backend.moments_window()
    if not args.no_activate:
        backend.activate(win)
        if args.activate_wait_ms > 0:
            time.sleep(args.activate_wait_ms / 1000)

    window_rect = backend.rect(win)
    if window_rect is None:
        raise WindowPositionUnavailable("无法读取朋友圈窗口位置")
    refresh_point = None if args.no_refresh else refresh_button_center(window_rect)

    last_error: Exception | None = None
    missing_sns_list_failures = 0
    for round_index in range(1, args.rounds + 1):
        try:
            print(f"[{current_timestamp_ms()}] UIA一图按键匹配: 第 {round_index}/{args.rounds} 轮")
            result = resolve_second_uia_list_item_post(
                backend,
                win,
                args.user,
                item_index=args.item_index - 1,
                settle_ms=args.settle_ms,
                include_text=True,
            )
            text = result.text.strip()
            print(f"  已匹配第 {args.item_index} 条: {result.detected_prefix}")
            print(f"  UIA 读取耗时: {result.elapsed_ms}ms")
            print(f"  图片数量: {result.inline_image_count if result.inline_image_count is not None else '未解析'}")
            if text:
                print("  文案:")
                for line in text.splitlines():
                    print(f"    {line}")

            if not is_one_image_post(text, result.inline_image_count):
                last_error = EasyMoneyError("目标用户已匹配，但文案未识别为一张图片")
                if round_index >= args.rounds:
                    break
                if refresh_point is None:
                    raise last_error
                print(f"  {last_error}，执行刷新后继续")
                click_refresh(input_backend, refresh_point, args.refresh_wait_ms)
                continue

            if args.capture_focus:
                print(f"  识别到一张图片，先按键: {key_sequence_label(FOCUS_CAPTURE_SEQUENCE)}")
                if args.dry_run:
                    print("  dry-run: 未实际发送 Down+3Tab，也未截图。")
                    return 0
                capture_focus_after_tabs(
                    backend,
                    input_backend,
                    window_rect,
                    args.key_gap_ms,
                    args.focus_wait_ms,
                    args.focus_padding,
                    args.output,
                )
                print("  capture-focus 模式: 未发送 Enter。")
                return 0

            print(f"  识别到一张图片，按键序列: {key_sequence_label()}")
            if args.dry_run:
                print("  dry-run: 未实际发送最终按键。")
                return 0

            key_gap = max(0, args.key_gap_ms) / 1000
            input_backend.press_sequence(KEY_SEQUENCE, gap=key_gap)
            print("  按键已发送。")
            return 0
        except WindowPositionUnavailable:
            raise
        except Exception as exc:
            last_error = exc
            if is_missing_sns_list_error(exc):
                missing_sns_list_failures += 1
                missing_sns_list_limit = max(0, int(os.environ.get("EASYMONEY_UIA_MISSING_SNS_REFRESH_LIMIT", "3")))
                if missing_sns_list_limit and missing_sns_list_failures >= missing_sns_list_limit:
                    print(
                        "  UIA连续未暴露 sns_list，停止刷新重试；"
                        "请确认朋友圈窗口已打开且当前微信版本仍暴露 sns_list"
                    )
                    break
            else:
                missing_sns_list_failures = 0
            if round_index >= args.rounds:
                print(f"  UIA一图按键匹配失败: {exc}")
                break
            if refresh_point is None:
                raise
            print(f"  UIA一图按键匹配失败，执行刷新后继续: {exc}")
            click_refresh(input_backend, refresh_point, args.refresh_wait_ms)

    raise EasyMoneyError(f"{last_error or 'UIA一图按键匹配失败'}；已尝试 {args.rounds} 轮，可用 --rounds N 调整")


if __name__ == "__main__":
    try:
        raise SystemExit(main(sys.argv[1:]))
    except KeyboardInterrupt:
        print()
        print("已停止")
        raise SystemExit(130)
    except UIAListItemUnavailable as exc:
        print(f"UIA 列表不可用: {exc}", file=sys.stderr)
        raise SystemExit(1)
    except EasyMoneyError as exc:
        print(f"错误: {exc}", file=sys.stderr)
        raise SystemExit(1)
