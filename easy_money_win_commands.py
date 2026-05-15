from __future__ import annotations

import concurrent.futures
import dataclasses
import math

from easy_money_win_core import *
from easy_money_win_input import *
from easy_money_win_capture import *
from easy_money_win_uia import *
from easy_money_win_llm import *

def print_usage() -> None:
    print(
        f"""
{APP_NAME} {APP_VERSION}

用法:
  python easy_money_win.py run [--interval N] [--pos x,y] [--index N] [--title 文本] [--id id]
  python easy_money_win.py locate [--test-click]
  python easy_money_win.py capture-info [--backend dxgi]
  python easy_money_win.py uia-dump [--max-depth N] [--item-index N] [--item-limit N] [--full] [--watch] [--count N] [--timing-only]
  python easy_money_win.py ax-dump [--max-depth N] [--item-index N] [--item-limit N] [--full] [--watch] [--count N] [--timing-only]
  python easy_money_win.py comment-locate
  python easy_money_win.py comment-fixed-send-locate
  python easy_money_win.py post-image-locate
  python easy_money_win.py post-image-x-locate
  python easy_money_win.py input-bench --kind click|keys --count N --confirm-send
  python easy_money_win.py comment [--text 文本] [--solve-question|--doubao|--LLM [--vision] [--save-vision-image] [--vision-output 路径]] --user <用户名前缀> [--debug] [--timing-detail]
  python easy_money_win.py llm ask "<问题>" [上下文]
  python easy_money_win.py doubao ask "<朋友圈正文>"  # 兼容入口：本地规则答题
"""
    )


def parse_option_value(args: list[str], index: int, name: str) -> tuple[str, int]:
    if index + 1 >= len(args) or args[index + 1].startswith("--"):
        raise EasyMoneyError(f"{name} 需要提供值")
    return args[index + 1], index + 1


def cmd_capture_info(args: list[str]) -> int:
    backend_name: Optional[str] = None
    i = 0
    while i < len(args):
        if args[i] == "--backend":
            backend_name, i = parse_option_value(args, i, "--backend")
        i += 1
    backend = WindowBackend()
    win_rect = backend.moments_window_rect()
    region = refresh_observation_region(win_rect)
    capture = CaptureBackend(backend_name)
    started = time.perf_counter()
    frame = capture.grab(region)
    elapsed_ms = int((time.perf_counter() - started) * 1000)
    print(f"截图后端: {capture.backend}")
    print(f"测试区域: {region.describe()}")
    print(f"帧尺寸: {frame.width}x{frame.height}")
    print(f"截图耗时: {elapsed_ms}ms")
    capture.close()
    return 0


def percentile_nearest_rank(sorted_values: list[int], percentile: float) -> int:
    if not sorted_values:
        return 0
    rank = max(1, int(math.ceil(percentile / 100.0 * len(sorted_values))))
    return sorted_values[min(rank - 1, len(sorted_values) - 1)]


def format_ns_as_ms(value_ns: int) -> str:
    return f"{value_ns / 1_000_000:.3f}ms"


def parse_point_text(raw: str, option_name: str = "--point") -> Point:
    parts = [part.strip() for part in raw.split(",")]
    if len(parts) != 2:
        raise EasyMoneyError(f"{option_name} 格式应为 x,y")
    return Point(float(parts[0]), float(parts[1]))


def parse_input_bench_options(args: list[str]) -> InputBenchOptions:
    options = InputBenchOptions()
    i = 0
    while i < len(args):
        arg = args[i]
        if arg == "--kind":
            value, i = parse_option_value(args, i, "--kind")
            kind = value.strip().lower()
            if kind not in {"click", "keys"}:
                raise EasyMoneyError("--kind 只支持 click 或 keys")
            options = dataclasses.replace(options, kind=kind)
        elif arg == "--count":
            value, i = parse_option_value(args, i, "--count")
            options = dataclasses.replace(options, count=max(1, int(value)))
        elif arg == "--warmup":
            value, i = parse_option_value(args, i, "--warmup")
            options = dataclasses.replace(options, warmup=max(0, int(value)))
        elif arg == "--keys":
            value, i = parse_option_value(args, i, "--keys")
            options = dataclasses.replace(options, keys=parse_key_sequence_text(value, "--keys"))
        elif arg == "--interval-ms":
            value, i = parse_option_value(args, i, "--interval-ms")
            options = dataclasses.replace(options, interval_ms=max(0, int(value)))
        elif arg == "--countdown-ms":
            value, i = parse_option_value(args, i, "--countdown-ms")
            options = dataclasses.replace(options, countdown_ms=max(0, int(value)))
        elif arg == "--point":
            value, i = parse_option_value(args, i, "--point")
            options = dataclasses.replace(options, point=parse_point_text(value, "--point"))
        elif arg == "--confirm-send":
            options = dataclasses.replace(options, confirm_send=True)
        elif arg == "--priority-boost":
            options = dataclasses.replace(options, priority_boost=True)
        else:
            raise EasyMoneyError(f"未知 input-bench 参数: {arg}")
        i += 1
    return options


def collect_input_bench_stats(values_ns: list[int]) -> dict[str, int]:
    if not values_ns:
        raise EasyMoneyError("没有可统计的输入耗时")
    sorted_values = sorted(values_ns)
    return {
        "min": sorted_values[0],
        "p50": percentile_nearest_rank(sorted_values, 50),
        "p95": percentile_nearest_rank(sorted_values, 95),
        "p99": percentile_nearest_rank(sorted_values, 99),
        "max": sorted_values[-1],
        "avg": int(sum(sorted_values) / len(sorted_values)),
    }


def print_input_bench_stats(values_ns: list[int], failures: int) -> None:
    stats = collect_input_bench_stats(values_ns)
    slow_5ms = sum(1 for value in values_ns if value >= 5_000_000)
    slow_10ms = sum(1 for value in values_ns if value >= 10_000_000)
    top_values = sorted(values_ns, reverse=True)[:5]
    print(
        "SendInput 耗时统计: "
        f"count={len(values_ns)} failures={failures} "
        f"min={format_ns_as_ms(stats['min'])} "
        f"p50={format_ns_as_ms(stats['p50'])} "
        f"p95={format_ns_as_ms(stats['p95'])} "
        f"p99={format_ns_as_ms(stats['p99'])} "
        f"max={format_ns_as_ms(stats['max'])} "
        f"avg={format_ns_as_ms(stats['avg'])}"
    )
    print(f"慢调用: >=5ms {slow_5ms}/{len(values_ns)} | >=10ms {slow_10ms}/{len(values_ns)}")
    print("最慢 5 次: " + ", ".join(format_ns_as_ms(value) for value in top_values))


def cmd_input_bench(args: list[str]) -> int:
    options = parse_input_bench_options(args)
    if not options.confirm_send:
        raise EasyMoneyError("input-bench 会真实发送鼠标/键盘输入；确认后请加 --confirm-send")
    if os.name != "nt":
        raise EasyMoneyError("input-bench 只支持 Windows 原生 SendInput")
    user32 = ctypes.windll.user32
    user32.SendInput.argtypes = (wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int)
    user32.SendInput.restype = wintypes.UINT

    if options.point is not None:
        x, y = options.point.rounded()
        user32.SetCursorPos(x, y)
        print(f"已移动到压测坐标: ({int(options.point.x)}, {int(options.point.y)})")
    if options.kind == "click":
        down = INPUT()
        down.type = INPUT_MOUSE
        down.union.mi = MOUSEINPUT(0, 0, 0, MOUSEEVENTF_LEFTDOWN, 0, 0)
        up = INPUT()
        up.type = INPUT_MOUSE
        up.union.mi = MOUSEINPUT(0, 0, 0, MOUSEEVENTF_LEFTUP, 0, 0)
        events = [down, up]
        label = "mouse click x1"
    else:
        events = []
        for key in options.keys:
            vk = InputBackend._vk(key)
            down = INPUT()
            down.type = INPUT_KEYBOARD
            down.union.ki = KEYBDINPUT(vk, 0, 0, 0, 0)
            up = INPUT()
            up.type = INPUT_KEYBOARD
            up.union.ki = KEYBDINPUT(vk, 0, KEYEVENTF_KEYUP, 0, 0)
            events.extend([down, up])
        label = f"keys {format_key_sequence(options.keys)}"
    array_type = INPUT * len(events)
    event_array = array_type(*events)
    count = len(events)

    total_iterations = options.warmup + options.count
    if options.countdown_ms > 0:
        print(f"将在 {options.countdown_ms}ms 后开始真实发送输入；按 Ctrl+C 可取消")
        precise_delay(options.countdown_ms / 1000.0)
    print(
        f"input-bench: kind={options.kind} label={label} count={options.count} "
        f"warmup={options.warmup} interval={options.interval_ms}ms priority_boost={options.priority_boost}"
    )

    elapsed_values: list[int] = []
    failures = 0
    delay_seconds = options.interval_ms / 1000.0
    priority_boost = CurrentThreadPriorityBoost(options.priority_boost)
    with priority_boost:
        for index in range(total_iterations):
            started_ns = time.perf_counter_ns()
            sent = user32.SendInput(count, event_array, ctypes.sizeof(INPUT))
            elapsed_ns = time.perf_counter_ns() - started_ns
            if index >= options.warmup:
                if sent != count:
                    failures += 1
                elapsed_values.append(elapsed_ns)
            if delay_seconds > 0 and index < total_iterations - 1:
                precise_delay(delay_seconds)

    print(f"输入线程优先级: {priority_boost.describe()}")
    print_input_bench_stats(elapsed_values, failures)
    return 0


def cmd_uia_dump(args: list[str]) -> int:
    max_depth = 7
    buttons_only = False
    full = False
    settle_ms = 1500
    view = "raw"
    list_name = "朋友圈"
    watch = False
    interval_seconds = 0.2
    count = 0
    timing_only = False
    item_limit = 1
    item_index = 2
    i = 0
    while i < len(args):
        if args[i] == "--max-depth":
            value, i = parse_option_value(args, i, "--max-depth")
            max_depth = max(1, min(int(value), 30))
        elif args[i] == "--buttons-only":
            buttons_only = True
            full = True
        elif args[i] == "--full":
            full = True
        elif args[i] == "--settle-ms":
            value, i = parse_option_value(args, i, "--settle-ms")
            settle_ms = max(0, int(value))
        elif args[i] == "--view":
            value, i = parse_option_value(args, i, "--view")
            view = value.strip().lower()
            if view not in {"raw", "default", "control", "content"}:
                raise EasyMoneyError("不支持的 --view，可用: raw|default|control|content")
        elif args[i] == "--list-name":
            list_name, i = parse_option_value(args, i, "--list-name")
        elif args[i] == "--item-limit":
            value, i = parse_option_value(args, i, "--item-limit")
            item_limit = max(0, int(value))
        elif args[i] == "--item-index":
            value, i = parse_option_value(args, i, "--item-index")
            item_index = max(1, int(value))
        elif args[i] in {"--watch", "--loop"}:
            watch = True
        elif args[i] == "--interval":
            value, i = parse_option_value(args, i, "--interval")
            interval_seconds = max(0.0, float(value))
        elif args[i] in {"--count", "--iterations"}:
            value, i = parse_option_value(args, i, args[i])
            count = max(0, int(value))
        elif args[i] == "--timing-only":
            timing_only = True
        i += 1
    backend = WindowBackend()
    win = backend.moments_window()

    def dump_once(current_win: Any, quiet: bool) -> tuple[Any, bool, bool, int, int]:
        started = time.perf_counter()
        found_sns = False
        found_item = False
        item_count = 0
        if full:
            _, found_sns, found_item = backend.dump_tree(current_win, max_depth=max_depth, buttons_only=buttons_only, view=view)
        else:
            deadline = time.perf_counter() + settle_ms / 1000
            while time.perf_counter() < deadline:
                found_sns, found_item, item_count = dump_named_list_contents(
                    backend,
                    current_win,
                    list_name,
                    max_depth=max_depth,
                    view=view,
                    quiet=quiet,
                    item_limit=item_limit,
                    item_index=item_index,
                )
                if found_sns:
                    break
                time.sleep(0.03 if watch else 0.12)
                try:
                    current_win = backend.moments_window()
                except Exception:
                    pass
            if not found_sns and not quiet:
                print(f'未找到 name="{list_name}" 的 List；可加 --full 查看完整 UIA 树。')
        elapsed_ms = int((time.perf_counter() - started) * 1000)
        return current_win, found_sns, found_item, item_count, elapsed_ms

    if watch:
        iteration = 0
        try:
            while count <= 0 or iteration < count:
                iteration += 1
                if not timing_only:
                    print(f"\n--- uia-dump watch #{iteration} ---")
                win, found_sns_list, found_list_item, item_count, elapsed_ms = dump_once(win, quiet=timing_only)
                print(
                    f"[watch] #{iteration} total={elapsed_ms}ms "
                    f"found_list={str(found_sns_list).lower()} items={item_count}"
                )
                if interval_seconds > 0 and (count <= 0 or iteration < count):
                    time.sleep(interval_seconds)
        except KeyboardInterrupt:
            print("\n已停止 uia-dump watch")
        return 0

    win, found_sns_list, found_list_item, _, _ = dump_once(win, quiet=False)
    if not buttons_only and (not found_sns_list or not found_list_item):
        print(f'\n提示: 当前未读取到 name="{list_name}" 的 ListItem。')
        print("      可先用 --full 查看完整树，或稍等、重新激活朋友圈窗口后再 dump。")
    return 0


def refresh_point_from_saved_offset(backend: WindowBackend) -> Optional[Point]:
    offset = load_point(CONFIG_REFRESH)
    if offset is None:
        return None
    win_rect = backend.moments_window_rect()
    return Point(win_rect.left + offset.x, win_rect.top + offset.y)


def cmd_locate(args: list[str]) -> int:
    test_click = False
    i = 0
    while i < len(args):
        if args[i] in {"--mouse", "--manual"}:
            pass
        elif args[i] == "--test-click":
            test_click = True
        i += 1

    backend = WindowBackend()
    input_backend = InputBackend()
    win = backend.moments_window()
    backend.activate(win)
    win_rect = backend.rect(win)
    if win_rect is None:
        raise EasyMoneyError("无法读取朋友圈窗口位置")

    print("手动标定刷新按钮位置：请将鼠标移到朋友圈顶部“刷新”按钮中心。")
    countdown()
    point = input_backend.position()
    offset = Point(point.x - win_rect.left, point.y - win_rect.top)
    save_point(CONFIG_REFRESH, offset)
    print(f"已定位刷新按钮: 鼠标位置 ({int(point.x)}, {int(point.y)})")
    print(f"窗口相对偏移: dx={int(offset.x)}, dy={int(offset.y)}")
    print(f"配置已保存: {CONFIG_REFRESH}")
    if test_click:
        input_backend.click(point, interval=0.0)
        print("已执行测试点击")
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
    if not any([target_pos, target_index is not None, target_title, target_id]):
        target_pos = refresh_point_from_saved_offset(backend)
        if target_pos is None:
            raise EasyMoneyError("未找到刷新按钮坐标配置，请先运行 locate")
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
                elif backend.click_control(chosen, input_backend, prefer_coordinate=True):
                    rect = backend.rect(chosen)
                    detail = f" @({int(rect.center.x)}, {int(rect.center.y)})" if rect else ""
                    print(f"[{now}] 已点击按钮 {backend._safe_text(chosen) or target_index}{detail}")
                else:
                    print(f"[{now}] 点击按钮失败")
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


def parse_comment_options(args: list[str]) -> CommentOptions:
    timing_detail_env = (os.environ.get("EASYMONEY_TIMING_DETAIL") or os.environ.get("EASYMONEY_INPUT_TIMING") or "").strip().lower()
    options = CommentOptions(
        submit_mode=(os.environ.get("EASYMONEY_SUBMIT_MODE") or "click").strip().lower(),
        timing_detail=timing_detail_env in {"1", "true", "yes", "on"},
    )
    user_filter = False
    user_name: Optional[str] = None

    i = 0
    while i < len(args):
        arg = args[i]
        if arg == "--text":
            options.comment_text, i = parse_option_value(args, i, "--text")
        elif arg == "--user":
            user_filter = True
            if i + 1 < len(args) and not args[i + 1].startswith("--"):
                user_name = args[i + 1]
                i += 1
        elif arg == "--solve-question":
            options.solve_question = True
        elif arg == "--doubao":
            options.solve_question = True
        elif arg == "--LLM":
            options.use_llm = True
            options.solve_question = True
        elif arg == "--vision":
            options.use_vision = True
        elif arg == "--save-vision-image":
            options.save_vision_image = True
        elif arg == "--vision-output":
            value, i = parse_option_value(args, i, "--vision-output")
            options.save_vision_image = True
            options.vision_save_path = expand_path(value)
        elif arg == "--debug":
            options.debug = True
        elif arg == "--save-post-image":
            options.save_post_image = True
        elif arg == "--output":
            value, i = parse_option_value(args, i, "--output")
            options.save_path = expand_path(value)
        elif arg == "--click-post-image":
            options.click_post_image = True
        elif arg in {"--test-image-crop", "--debug-image-crop"}:
            options.test_image_crop = True
        elif arg == "--rounds":
            value, i = parse_option_value(args, i, "--rounds")
            options.rounds = max(1, int(value))
        elif arg == "--submit-click":
            options.submit_mode = "click"
        elif arg == "--submit-mode":
            options.submit_mode, i = parse_option_value(args, i, "--submit-mode")
            options.submit_mode = options.submit_mode.strip().lower()
        elif arg in {"--timing-detail", "--input-timing", "--trace-input"}:
            options.timing_detail = True
        elif arg in {"--ocr-comment", "--stream-capture"}:
            print(f"提示: Windows v1 暂不完整支持 {arg}，已忽略或降级")
        else:
            raise EasyMoneyError(f"未知 comment 参数: {arg}")
        i += 1

    if not user_filter:
        raise EasyMoneyError("comment 命令必须显式指定 --user <用户名前缀>")
    if options.use_vision and not options.use_llm:
        raise EasyMoneyError("--vision 需要与 --LLM 一起使用")
    if options.save_vision_image and not options.use_vision:
        raise EasyMoneyError("--save-vision-image 需要与 --LLM --vision 一起使用")
    if not any([options.comment_text, options.solve_question, options.save_post_image, options.click_post_image, options.test_image_crop]):
        raise EasyMoneyError('请指定 --text "评论内容"，或使用 --solve-question / --doubao / --LLM / --save-post-image')

    options.requested_user = (user_name or "").strip()
    if not options.requested_user:
        raise EasyMoneyError("comment --user 需要提供用户名前缀，用于匹配朋友圈第 2 个 ListItem 的开头")
    return options


def comment_requires_config(options: CommentOptions) -> bool:
    return not any([options.save_post_image, options.click_post_image, options.test_image_crop])


def normalize_comment_mode(mode: str, option_name: str) -> str:
    normalized = mode.strip().lower()
    if normalized in {"mouse", "coordinate"}:
        return "click"
    if normalized == "keyboard":
        return "keys"
    if normalized not in {"keys", "click"}:
        raise EasyMoneyError(f"{option_name} 只支持 keys 或 click")
    return normalized


def uia_worker_thread_enabled() -> bool:
    value = os.environ.get("EASYMONEY_UIA_WORKER_THREAD", "1").strip().lower()
    return value not in {"0", "false", "no", "off"}


def comment_input_priority_boost_enabled() -> bool:
    value = os.environ.get("EASYMONEY_INPUT_PRIORITY_BOOST", "1").strip().lower()
    return value not in {"0", "false", "no", "off"}


class CurrentThreadPriorityBoost:
    def __init__(self, enabled: bool, priority: int = THREAD_PRIORITY_HIGHEST) -> None:
        self.enabled = enabled
        self.priority = priority
        self.previous_priority: Optional[int] = None
        self.applied = False
        self.error = ""
        self._kernel32: Any = None
        self._thread_handle: Any = None

    def __enter__(self) -> "CurrentThreadPriorityBoost":
        if not self.enabled:
            return self
        if os.name != "nt":
            self.error = "非 Windows"
            return self
        try:
            kernel32 = ctypes.windll.kernel32
            kernel32.GetCurrentThread.restype = wintypes.HANDLE
            kernel32.GetThreadPriority.argtypes = (wintypes.HANDLE,)
            kernel32.GetThreadPriority.restype = ctypes.c_int
            kernel32.SetThreadPriority.argtypes = (wintypes.HANDLE, ctypes.c_int)
            kernel32.SetThreadPriority.restype = wintypes.BOOL

            thread = kernel32.GetCurrentThread()
            previous = kernel32.GetThreadPriority(thread)
            if previous == THREAD_PRIORITY_ERROR_RETURN:
                self.error = "GetThreadPriority 失败"
                return self
            self._kernel32 = kernel32
            self._thread_handle = thread
            self.previous_priority = previous
            if kernel32.SetThreadPriority(thread, self.priority):
                self.applied = True
            else:
                self.error = "SetThreadPriority 失败"
        except Exception as exc:
            self.error = f"{exc.__class__.__name__}: {exc}"
        return self

    def __exit__(self, exc_type: object, exc: object, tb: object) -> None:
        if not self.applied or self.previous_priority is None or self._kernel32 is None or self._thread_handle is None:
            return
        try:
            self._kernel32.SetThreadPriority(self._thread_handle, self.previous_priority)
        except Exception:
            pass

    def describe(self) -> str:
        if not self.enabled:
            return "关闭"
        if self.applied:
            return f"已启用 THREAD_PRIORITY_HIGHEST，原优先级={self.previous_priority}"
        return f"未启用: {self.error or '未知原因'}"


def _initialize_com_for_uia_worker() -> Callable[[], None]:
    try:
        comtypes = require_module("comtypes", "comtypes")
        co_initialize = getattr(comtypes, "CoInitialize", None)
        co_uninitialize = getattr(comtypes, "CoUninitialize", None)
        if callable(co_initialize):
            co_initialize()
            if callable(co_uninitialize):
                return co_uninitialize
    except Exception:
        pass
    return lambda: None


def _is_missing_sns_list_error(exc: Exception) -> bool:
    return isinstance(exc, UIAListItemUnavailable) and "sns_list" in str(exc)


def resolve_comment_target_post(
    backend: WindowBackend,
    input_backend: InputBackend,
    win: Any,
    requested_user: str,
    window_rect: Rect,
    rounds: int,
    include_text: bool = True,
    uia_mode: str = "当前线程",
) -> tuple[MomentPostResolution, Rect]:
    post: Optional[MomentPostResolution] = None
    last_error: Optional[Exception] = None
    missing_sns_list_failures = 0
    refresh_offset = load_point(CONFIG_REFRESH)
    refresh_button_center = (
        Point(window_rect.left + refresh_offset.x, window_rect.top + refresh_offset.y)
        if refresh_offset is not None
        else None
    )
    for round_index in range(1, rounds + 1):
        try:
            print(f"[{current_timestamp_ms()}] UIA用户匹配({uia_mode}): 第 {round_index}/{rounds} 轮")
            list_item = resolve_second_uia_list_item_post(
                backend,
                win,
                requested_user,
                item_index=1,
                settle_ms=int(os.environ.get("EASYMONEY_UIA_USER_SETTLE_MS", "220")),
                include_text=include_text,
            )
            post = MomentPostResolution(
                body_frame=list_item.body_frame,
                action_point=list_item.action_point,
                text=list_item.text,
                source=(
                    f"UIA:ListItem #{list_item.item_index + 1} "
                    f"prefix={list_item.detected_prefix or '(空)'} total={list_item.elapsed_ms}ms thread={uia_mode}"
                ),
                inline_image_count=list_item.inline_image_count,
            )
            print(
                f"  UIA用户匹配成功({uia_mode}): "
                f"user={requested_user} "
                f"item=#{list_item.item_index + 1} "
                f"prefix={list_item.detected_prefix or '(空)'} "
                f"frame={list_item.body_frame.describe()} "
                f"耗时={list_item.elapsed_ms}ms"
            )
            return post, window_rect
        except WindowPositionUnavailable:
            raise
        except Exception as exc:
            last_error = exc
            if _is_missing_sns_list_error(exc):
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
            if round_index >= rounds:
                print(f"  UIA用户匹配失败({uia_mode}): {exc}")
                break
            if refresh_offset is None:
                raise EasyMoneyError("UIA用户匹配失败且未找到刷新按钮坐标配置，请先运行 locate") from exc
            print(f"  UIA用户匹配失败({uia_mode})，执行刷新后继续: {exc}")
            if refresh_button_center is None:
                raise EasyMoneyError("UIA用户匹配失败且未找到刷新按钮坐标配置，请先运行 locate") from exc
            input_backend.click(refresh_button_center)
            time.sleep(COMMENT_REFRESH_WAIT_SECONDS)
            print(
                f"  已点击 locate 保存的刷新坐标: ({int(refresh_button_center.x)}, {int(refresh_button_center.y)})，"
                f"等待 {int(COMMENT_REFRESH_WAIT_SECONDS * 1000)}ms"
            )

    raise EasyMoneyError(f"{last_error or 'UIA用户匹配失败'}；已尝试 {rounds} 轮，可用 --rounds N 调整")


def resolve_comment_target_post_worker_entry(
    requested_user: str,
    rounds: int,
    include_text: bool,
    uia_mode: str,
) -> tuple[MomentPostResolution, Rect]:
    uninitialize_com = _initialize_com_for_uia_worker()
    try:
        backend = WindowBackend()
        input_backend = InputBackend()
        win = backend.moments_window()
        backend.activate(win)
        window_rect = backend.rect(win)
        if window_rect is None:
            raise WindowPositionUnavailable("无法读取朋友圈窗口位置")
        return resolve_comment_target_post(
            backend,
            input_backend,
            win,
            requested_user,
            window_rect,
            rounds,
            include_text=include_text,
            uia_mode=uia_mode,
        )
    finally:
        uninitialize_com()


def resolve_comment_target_post_via_worker(requested_user: str, rounds: int, include_text: bool = True) -> tuple[MomentPostResolution, Rect]:
    if not uia_worker_thread_enabled():
        return resolve_comment_target_post_worker_entry(requested_user, rounds, include_text, "主线程")
    with concurrent.futures.ThreadPoolExecutor(max_workers=1, thread_name_prefix="easymoney-uia") as executor:
        return executor.submit(resolve_comment_target_post_worker_entry, requested_user, rounds, include_text, "worker线程").result()


def print_resolved_comment_post(requested_user: str, post: MomentPostResolution) -> None:
    print(f"已匹配用户: {requested_user}")
    print(f"动态定位: {post.source} frame={post.body_frame.describe()}")
    if post.text:
        print("正文内容开始")
        print(post.text)
        print("正文内容结束")


def save_comment_post_image(post: MomentPostResolution, window_rect: Rect, save_path: Optional[Path]) -> int:
    image = capture_post_image(post, window_rect)
    output = save_path or DEBUG_DIR / f"wechat_post_image_{time.strftime('%Y%m%d_%H%M%S')}.png"
    ensure_parent(output)
    image.save(output)
    print(f"动态图片/区域已保存: {output}")
    return 0


def click_comment_post_image(
    input_backend: InputBackend,
    post: MomentPostResolution,
    window_rect: Rect,
    debug: bool = False,
) -> int:
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


def resolve_comment_text(
    options: CommentOptions,
    post: MomentPostResolution,
    window_rect: Rect,
    input_backend: Optional[InputBackend] = None,
) -> str:
    final_text = (options.comment_text or "").strip()
    if not options.solve_question:
        return final_text

    context = post.text.strip()
    if not context:
        raise EasyMoneyError("需要自动答题但未能读取朋友圈正文")

    if not options.use_llm:
        solved = solve_post_question_by_rules(context)
        final_text = solved.answer
        print_ts(f"本地规则命中: {solved.answer} ({solved.evidence})")
        return final_text

    image_urls: list[str] = []
    if options.use_vision:
        vision_save_path = None
        if options.save_vision_image:
            vision_save_path = options.vision_save_path or DEBUG_DIR / f"wechat_vision_image_{time.strftime('%Y%m%d_%H%M%S')}.png"
        image_urls = capture_vision_image_data_urls(post, window_rect, save_path=vision_save_path, input_backend=input_backend)
        print_ts(f"已附带视觉截图: {len(image_urls)} 张")
        if vision_save_path is not None:
            print_ts(f"视觉截图已保存: {vision_save_path}")
    config = load_local_llm_config()
    if config:
        answer = request_llm_answer(
            config,
            question_solve_system_prompt(),
            build_question_solve_prompt(context),
            image_data_urls=image_urls,
        )
        if answer:
            final_text = answer
            print_ts(f"LLM 命中: {answer}")
    else:
        print_ts("LLM 配置缺失：请检查 .easyMoney.env")
    if not final_text and options.comment_text:
        final_text = options.comment_text.strip()
        print_ts("自动答题未命中，回退到 --text")
    if not final_text:
        raise EasyMoneyError("未能生成评论内容，请补充 --text 作为回退")
    return final_text


def should_type_comment_text_directly(text: str, input_backend: InputBackend) -> bool:
    if not text:
        return False
    can_type_text_directly = getattr(input_backend, "can_type_text_directly", None)
    if not callable(can_type_text_directly):
        can_type_text_directly = getattr(input_backend, "can_type_directly", None)
    return bool(callable(can_type_text_directly) and can_type_text_directly(text))


def prewarm_comment_text_input(options: CommentOptions, input_backend: InputBackend) -> None:
    text = (options.comment_text or "").strip()
    if not should_type_comment_text_directly(text, input_backend):
        return
    prepare_text_input = getattr(input_backend, "prepare_text_input", None)
    if callable(prepare_text_input):
        prepare_text_input(text)


def build_comment_send_plan(
    options: CommentOptions,
    post: MomentPostResolution,
    window_rect: Rect,
    config: CommentConfig,
    final_text: str,
) -> CommentSendPlan:
    open_comment_keys = parse_key_sequence_text(os.environ.get("EASYMONEY_OPEN_COMMENT_KEYS", "tab,enter"), "EASYMONEY_OPEN_COMMENT_KEYS")
    submit_comment_keys = parse_key_sequence_text(
        os.environ.get("EASYMONEY_SUBMIT_KEYS", "tab,tab,tab,enter"),
        "EASYMONEY_SUBMIT_KEYS",
    )
    submit_mode = normalize_comment_mode(options.submit_mode, "--submit-mode")
    send_point, send_point_method = resolve_send_point(post.action_point, window_rect, config)
    submit_method = f"点击发送按钮[{send_point_method}]" if submit_mode == "click" else format_key_sequence(submit_comment_keys)
    return CommentSendPlan(
        text=final_text,
        action_point=post.action_point,
        send_point=send_point,
        send_point_method=send_point_method,
        submit_mode=submit_mode,
        open_comment_keys=open_comment_keys,
        submit_comment_keys=submit_comment_keys,
        submit_method=submit_method,
        timing_detail=options.timing_detail,
    )


def print_comment_debug(plan: CommentSendPlan, input_backend: InputBackend) -> None:
    input_backend.move_to(plan.action_point)
    print(f"DEBUG: 操作按钮点 ({int(plan.action_point.x)}, {int(plan.action_point.y)})")
    print(f"DEBUG: 打开评论方式: {format_key_sequence(plan.open_comment_keys)}")
    print(f"DEBUG: 发送方式: {plan.submit_method}")
    print(f"DEBUG: 发送点参考 [{plan.send_point_method}] ({int(plan.send_point.x)}, {int(plan.send_point.y)})")
    print(f"DEBUG: 评论内容: {plan.text}")


def execute_comment_send_plan(
    plan: CommentSendPlan,
    input_backend: InputBackend,
) -> CommentSendResult:
    can_direct_type_text = should_type_comment_text_directly(plan.text, input_backend)
    input_backend.prepare_key_sequence(plan.open_comment_keys)
    if plan.submit_mode == "keys":
        input_backend.prepare_key_sequence(plan.submit_comment_keys)
    prepare_mouse_click = getattr(input_backend, "prepare_mouse_click", None)
    if callable(prepare_mouse_click):
        prepare_mouse_click(1)
    prepare_text_input = getattr(input_backend, "prepare_text_input", None)
    if can_direct_type_text and callable(prepare_text_input):
        prepare_text_input(plan.text)

    set_input_timing_enabled = getattr(input_backend, "set_input_timing_enabled", None)
    set_input_timing_context = getattr(input_backend, "set_input_timing_context", None)
    if callable(set_input_timing_enabled):
        set_input_timing_enabled(plan.timing_detail)

    def set_timing_context(context: str) -> None:
        if callable(set_input_timing_context):
            set_input_timing_context(context)

    priority_boost = CurrentThreadPriorityBoost(comment_input_priority_boost_enabled())
    with priority_boost:
        send_flow_start = time.perf_counter()
        step_start = time.perf_counter()
        set_timing_context("点操作")
        input_backend.click(plan.action_point, interval=0.0)
        action_click_ms = int((time.perf_counter() - step_start) * 1000)

        step_start = time.perf_counter()
        set_timing_context("打开评论")
        input_backend.press_sequence_atomic(plan.open_comment_keys)
        open_comment_ms = int((time.perf_counter() - step_start) * 1000)

        step_start = time.perf_counter()
        set_timing_context("输入")
        if can_direct_type_text:
            text_input_method = input_backend.type_text_directly(plan.text)
        else:
            text_input_method = input_backend.paste_text(
                plan.text,
                restore_clipboard=False,
                before_paste_delay=0.0,
                after_paste_delay=0.012,
            )
        text_input_ms = int((time.perf_counter() - step_start) * 1000)

        send_step_label = "发送点击" if plan.submit_mode == "click" else "发送快捷键"
        step_start = time.perf_counter()
        if plan.submit_mode == "click":
            set_timing_context("发送点击")
            input_backend.click(plan.send_point, interval=0.0)
            send_submit_ms = int((time.perf_counter() - step_start) * 1000)
        else:
            set_timing_context("发送快捷键")
            input_backend.press_sequence_atomic(plan.submit_comment_keys)
            send_submit_ms = int((time.perf_counter() - step_start) * 1000)
        total_send_ms = int((time.perf_counter() - send_flow_start) * 1000)
    input_timings_fn = getattr(input_backend, "input_timings", None)
    input_timings = input_timings_fn() if callable(input_timings_fn) else ()
    return CommentSendResult(
        text_input_method=text_input_method,
        action_click_ms=action_click_ms,
        open_comment_ms=open_comment_ms,
        text_input_ms=text_input_ms,
        send_submit_ms=send_submit_ms,
        total_send_ms=total_send_ms,
        send_step_label=send_step_label,
        input_timings=input_timings,
        priority_boost_detail=priority_boost.describe(),
    )


def print_comment_send_result(plan: CommentSendPlan, result: CommentSendResult) -> None:
    print(
        f"已执行评论发送: {result.text_input_method} | 打开评论={format_key_sequence(plan.open_comment_keys)} | "
        f"发送方式={plan.submit_method} | 发送点参考=({int(plan.send_point.x)}, {int(plan.send_point.y)})"
    )
    print(
        f"发送流程耗时: 总计={result.total_send_ms}ms | "
        f"点操作={result.action_click_ms}ms | 打开评论={result.open_comment_ms}ms | "
        f"输入={result.text_input_ms}ms | {result.send_step_label}={result.send_submit_ms}ms"
    )
    if plan.timing_detail and result.priority_boost_detail:
        print(f"输入线程优先级: {result.priority_boost_detail}")
    if result.input_timings:
        print("底层输入耗时:")
        for label, elapsed_ns in result.input_timings:
            print(f"  {label}: {elapsed_ns / 1_000_000:.3f}ms")


def cmd_comment(args: list[str]) -> int:
    options = parse_comment_options(args)
    config = load_comment_config()
    if not config and comment_requires_config(options):
        raise EasyMoneyError("未找到评论配置，请先运行 comment-locate")

    input_backend = InputBackend()
    prewarm_comment_text_input(options, input_backend)
    post, window_rect = resolve_comment_target_post_via_worker(
        options.requested_user,
        options.rounds,
        include_text=options.solve_question,
    )
    print_resolved_comment_post(options.requested_user, post)

    if options.save_post_image or options.test_image_crop:
        return save_comment_post_image(post, window_rect, options.save_path)
    if options.click_post_image:
        return click_comment_post_image(input_backend, post, window_rect, debug=options.debug)

    final_text = resolve_comment_text(options, post, window_rect, input_backend=input_backend)
    if config is None:
        raise EasyMoneyError("未找到评论配置，请先运行 comment-locate")
    plan = build_comment_send_plan(options, post, window_rect, config, final_text)
    if options.debug:
        print_comment_debug(plan, input_backend)
        return 0

    result = execute_comment_send_plan(plan, input_backend)
    print_comment_send_result(plan, result)
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
    solved = solve_post_question_by_rules(args[1])
    print(f"答案: {solved.answer}")
    if solved.evidence:
        print(f"证据: {solved.evidence}")
    print(f"置信度: {solved.confidence:.2f}")
    return 0


def main(argv: Optional[list[str]] = None) -> int:
    configure_timestamped_logging()
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            try:
                reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass
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
        "capture-info": cmd_capture_info,
        "locate": cmd_locate,
        "run": cmd_run,
        "comment-locate": cmd_comment_locate,
        "comment-fixed-send-locate": cmd_comment_fixed_send_locate,
        "post-image-locate": lambda a: cmd_post_image_locate(a, x_only=False),
        "post-image-x-locate": lambda a: cmd_post_image_locate(a, x_only=True),
        "input-bench": cmd_input_bench,
        "comment": cmd_comment,
        "llm": cmd_llm,
        "doubao": cmd_doubao,
    }
    handler = dispatch.get(mode)
    if handler is None:
        print_usage()
        raise EasyMoneyError(f"未知命令: {mode}")
    return handler(rest)
